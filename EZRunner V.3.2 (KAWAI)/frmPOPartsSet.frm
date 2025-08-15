VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPOPartsSet 
   BackColor       =   &H00FDDFE3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase Order (Set)"
   ClientHeight    =   11025
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15240
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   11025
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   645
      Left            =   60
      TabIndex        =   91
      Top             =   840
      Width           =   7305
      Begin VB.CommandButton CmdCari 
         BackColor       =   &H0080FFFF&
         Caption         =   "Search"
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
         Left            =   5790
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   180
         Width           =   1125
      End
      Begin VB.CheckBox ChkUpdate 
         BackColor       =   &H00FDDFE3&
         Caption         =   "Update"
         Height          =   225
         Left            =   1170
         TabIndex        =   93
         Top             =   255
         Width           =   885
      End
      Begin VB.CheckBox ChkCreate 
         BackColor       =   &H00FDDFE3&
         Caption         =   "Create"
         Height          =   225
         Left            =   210
         TabIndex        =   92
         Top             =   255
         Width           =   795
      End
      Begin MSForms.ComboBox CboSetCode 
         Height          =   315
         Left            =   4170
         TabIndex        =   96
         Top             =   210
         Width           =   1485
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "2619;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase Set Code"
         Height          =   195
         Index           =   0
         Left            =   2700
         TabIndex        =   95
         Top             =   270
         Width           =   1635
      End
   End
   Begin VB.TextBox TxtPOLOT 
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
      MaxLength       =   35
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   8190
      Width           =   4650
   End
   Begin VB.TextBox txtMarking 
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
      Left            =   12990
      MaxLength       =   25
      TabIndex        =   23
      Top             =   7305
      Width           =   2085
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H0080FFFF&
      Caption         =   "Find [F3]"
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
      Left            =   5490
      Style           =   1  'Graphical
      TabIndex        =   86
      TabStop         =   0   'False
      Top             =   6480
      Width           =   1125
   End
   Begin VB.TextBox txtSearch 
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
      Left            =   2985
      MaxLength       =   25
      TabIndex        =   85
      TabStop         =   0   'False
      Top             =   6510
      Width           =   2430
   End
   Begin VB.TextBox txtRevisi 
      Alignment       =   2  'Center
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
      Left            =   5370
      MaxLength       =   2
      TabIndex        =   7
      Top             =   2625
      Width           =   450
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1080
      Left            =   62
      TabIndex        =   71
      Top             =   1425
      Width           =   15105
      Begin VB.TextBox LblMat 
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
         Height          =   210
         Left            =   12000
         Locked          =   -1  'True
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   705
         Width           =   2670
      End
      Begin VB.TextBox txtAddress 
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
         Height          =   210
         Left            =   9360
         Locked          =   -1  'True
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   277
         Width           =   5355
      End
      Begin VB.TextBox txtSupplier 
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
         Height          =   210
         Left            =   3030
         Locked          =   -1  'True
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   277
         Width           =   4995
      End
      Begin VB.TextBox txtDeliverTo 
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
         Height          =   210
         Left            =   12015
         Locked          =   -1  'True
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   2790
      End
      Begin VB.TextBox txtWHTo 
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
         Height          =   210
         Left            =   5895
         Locked          =   -1  'True
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   675
         Width           =   3345
      End
      Begin MSComCtl2.DTPicker dtpPeriod 
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Top             =   615
         Width           =   1305
         _ExtentX        =   2302
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
         CustomFormat    =   "MMM yyyy"
         Format          =   141230083
         UpDown          =   -1  'True
         CurrentDate     =   37798
      End
      Begin VB.Line Line3 
         X1              =   12000
         X2              =   14655
         Y1              =   960
         Y2              =   960
      End
      Begin MSForms.ComboBox CboMat 
         Height          =   315
         Left            =   10545
         TabIndex        =   3
         Top             =   660
         Width           =   1305
         VariousPropertyBits=   612386843
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2302;556"
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
         Index           =   29
         Left            =   9420
         TabIndex        =   90
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier CD"
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
         Left            =   240
         TabIndex        =   80
         Top             =   285
         Width           =   1035
      End
      Begin VB.Line Line1 
         X1              =   3030
         X2              =   8070
         Y1              =   525
         Y2              =   525
      End
      Begin MSForms.ComboBox cboSupplier 
         Height          =   315
         Left            =   1560
         TabIndex        =   0
         Top             =   225
         Width           =   1305
         VariousPropertyBits=   612386843
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2302;556"
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
         Height          =   195
         Index           =   1
         Left            =   8355
         TabIndex        =   79
         Top             =   285
         Width           =   690
      End
      Begin VB.Line Line2 
         X1              =   9360
         X2              =   14760
         Y1              =   525
         Y2              =   525
      End
      Begin VB.Label lblCaption 
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
         Index           =   2
         Left            =   240
         TabIndex        =   78
         Top             =   675
         Width           =   540
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Deliver To"
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
         Left            =   9555
         TabIndex        =   77
         Top             =   255
         Visible         =   0   'False
         Width           =   885
      End
      Begin MSForms.ComboBox cboDeliverTo 
         Height          =   315
         Left            =   10560
         TabIndex        =   5
         Top             =   195
         Visible         =   0   'False
         Width           =   1305
         VariousPropertyBits=   612386843
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2302;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Line Line9 
         Visible         =   0   'False
         X1              =   12015
         X2              =   14790
         Y1              =   495
         Y2              =   495
      End
      Begin VB.Line Line10 
         X1              =   5895
         X2              =   9225
         Y1              =   930
         Y2              =   930
      End
      Begin MSForms.ComboBox cboWHTo 
         Height          =   315
         Left            =   4440
         TabIndex        =   2
         Top             =   630
         Width           =   1305
         VariousPropertyBits=   612386843
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2302;556"
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
         Caption         =   "Warehouse To"
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
         Left            =   3030
         TabIndex        =   76
         Top             =   690
         Width           =   1230
      End
   End
   Begin VB.TextBox txtMarking 
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
      Left            =   7455
      MaxLength       =   25
      TabIndex        =   19
      Top             =   7290
      Width           =   2085
   End
   Begin VB.TextBox txtMarking 
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
      Left            =   10260
      MaxLength       =   25
      TabIndex        =   21
      Top             =   7305
      Width           =   2085
   End
   Begin VB.TextBox txtPacking 
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
      Height          =   210
      Left            =   2985
      Locked          =   -1  'True
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   7575
      Width           =   3585
   End
   Begin VB.TextBox txtInsurance 
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
      Height          =   210
      Left            =   2985
      Locked          =   -1  'True
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   7905
      Visible         =   0   'False
      Width           =   3585
   End
   Begin VB.TextBox txtRemarks 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   7710
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   24
      Top             =   7800
      Width           =   7470
   End
   Begin VB.TextBox txtMarking 
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
      Left            =   12990
      MaxLength       =   25
      TabIndex        =   22
      Top             =   6915
      Width           =   2085
   End
   Begin VB.TextBox txtTransport 
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
      Height          =   210
      Left            =   2985
      Locked          =   -1  'True
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   7890
      Width           =   3585
   End
   Begin VB.TextBox txtPaymentTerm 
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
      Height          =   210
      Left            =   2985
      Locked          =   -1  'True
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   6900
      Width           =   3585
   End
   Begin VB.TextBox txtPriceCondition 
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
      Height          =   210
      Left            =   2985
      Locked          =   -1  'True
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   7230
      Width           =   3570
   End
   Begin VB.TextBox txtMarking 
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
      Left            =   7455
      MaxLength       =   25
      TabIndex        =   18
      Top             =   6915
      Width           =   2085
   End
   Begin VB.TextBox txtMarking 
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
      Left            =   10260
      MaxLength       =   25
      TabIndex        =   20
      Top             =   6915
      Width           =   2085
   End
   Begin VB.TextBox txtPONo 
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
      Height          =   255
      Left            =   2055
      MaxLength       =   25
      TabIndex        =   6
      Top             =   2655
      Width           =   2430
   End
   Begin VB.CommandButton cmdReport 
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
      Left            =   10397
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   10095
      Width           =   1125
   End
   Begin VB.TextBox txtPPn 
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
      Height          =   285
      Left            =   10277
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   8955
      Width           =   2355
   End
   Begin VB.TextBox txtAmount 
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
      Height          =   285
      Left            =   7847
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   8955
      Width           =   2355
   End
   Begin VB.TextBox txtPONo2 
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
      Height          =   285
      Left            =   167
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   8955
      Width           =   2490
   End
   Begin VB.CommandButton command1 
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
      Index           =   3
      Left            =   11612
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   10095
      Width           =   1125
   End
   Begin VB.CommandButton command1 
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
      Index           =   2
      Left            =   12825
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2595
      Width           =   1125
   End
   Begin VB.TextBox txtGrandTotal 
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
      Height          =   285
      Left            =   12707
      Locked          =   -1  'True
      MaxLength       =   35
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   8955
      Width           =   2355
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FDDFE3&
      Height          =   533
      Left            =   62
      TabIndex        =   43
      Top             =   9375
      Width           =   15105
      Begin VB.Label lblErrMsg 
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
         Height          =   225
         Left            =   105
         TabIndex        =   44
         Top             =   195
         Width           =   14880
      End
   End
   Begin VB.CommandButton command1 
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
      Index           =   0
      Left            =   14042
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   10095
      Width           =   1125
   End
   Begin VB.CommandButton cmdSubMenu 
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
      Left            =   62
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   10095
      Width           =   1125
   End
   Begin VB.CommandButton command2 
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
      Left            =   6722
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   10095
      Width           =   1125
   End
   Begin VB.CommandButton command2 
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
      Left            =   5462
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   10095
      Width           =   1125
   End
   Begin VB.CommandButton command2 
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
      Left            =   4202
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   10095
      Width           =   1125
   End
   Begin VB.CommandButton command2 
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
      Left            =   2942
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   10095
      Width           =   1125
   End
   Begin VB.CommandButton command1 
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
      Index           =   1
      Left            =   12827
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   10095
      Width           =   1125
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   13335
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   465
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   714
   End
   Begin MSComCtl2.DTPicker dtpPODate 
      Height          =   315
      Left            =   6720
      TabIndex        =   8
      Top             =   2625
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
      Format          =   141230083
      CurrentDate     =   37798
   End
   Begin MSComCtl2.DTPicker dtpDeliveryDate 
      Height          =   315
      Left            =   9630
      TabIndex        =   9
      Top             =   2625
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
      Format          =   141230083
      CurrentDate     =   37798
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   3390
      Left            =   90
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   3030
      Width           =   15105
      _cx             =   26644
      _cy             =   5980
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
      Editable        =   1
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
      Begin MSForms.ComboBox cboCurr 
         Height          =   285
         Left            =   11625
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   -30
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
      Begin MSForms.ComboBox cboPrice 
         Height          =   285
         Left            =   9585
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   -30
         Visible         =   0   'False
         Width           =   2055
         VariousPropertyBits=   746604571
         MaxLength       =   16
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
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PO LOT"
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
      Left            =   120
      TabIndex        =   88
      Top             =   8250
      Width           =   630
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Line6"
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
      Left            =   12465
      TabIndex        =   87
      Top             =   7365
      Width           =   450
   End
   Begin MSForms.ComboBox cboSearch 
      Height          =   315
      Left            =   825
      TabIndex        =   84
      TabStop         =   0   'False
      Top             =   6510
      Width           =   2085
      VariousPropertyBits=   612386843
      MaxLength       =   6
      DisplayStyle    =   7
      Size            =   "3678;556"
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
      Caption         =   "Search"
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
      Left            =   135
      TabIndex        =   83
      Top             =   6570
      Width           =   600
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rev."
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
      Left            =   4890
      TabIndex        =   82
      Top             =   2685
      Width           =   390
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Order (Set)"
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
      Left            =   120
      TabIndex        =   81
      Top             =   240
      Width           =   15105
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00A6D2FF&
      Height          =   915
      Left            =   6705
      Top             =   6810
      Width           =   8475
   End
   Begin MSForms.ComboBox cboPacking 
      Height          =   315
      Left            =   1935
      TabIndex        =   14
      Top             =   7515
      Width           =   975
      VariousPropertyBits=   612386843
      MaxLength       =   6
      DisplayStyle    =   3
      Size            =   "1720;556"
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
      Caption         =   "Line2 "
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
      Left            =   6810
      TabIndex        =   70
      Top             =   7350
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
      Index           =   21
      Left            =   9720
      TabIndex        =   69
      Top             =   7365
      Width           =   450
   End
   Begin VB.Line Line8 
      Visible         =   0   'False
      X1              =   2985
      X2              =   6555
      Y1              =   8145
      Y2              =   8145
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
      Index           =   8
      Left            =   6810
      TabIndex        =   66
      Top             =   7830
      Width           =   765
   End
   Begin VB.Line Line7 
      X1              =   2985
      X2              =   6555
      Y1              =   8145
      Y2              =   8145
   End
   Begin VB.Line Line6 
      X1              =   2985
      X2              =   6555
      Y1              =   7815
      Y2              =   7815
   End
   Begin VB.Line Line5 
      X1              =   3000
      X2              =   6555
      Y1              =   7155
      Y2              =   7155
   End
   Begin VB.Line Line4 
      X1              =   3000
      X2              =   6555
      Y1              =   7485
      Y2              =   7485
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PO Marking"
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
      Left            =   6795
      TabIndex        =   62
      Top             =   6555
      Width           =   975
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Line5"
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
      Left            =   12465
      TabIndex        =   61
      Top             =   6975
      Width           =   450
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Insurance Covered"
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
      Left            =   135
      TabIndex        =   60
      Top             =   7905
      Visible         =   0   'False
      Width           =   1650
   End
   Begin MSForms.ComboBox cboInsuranceCls 
      Height          =   315
      Left            =   1920
      TabIndex        =   16
      Top             =   7860
      Visible         =   0   'False
      Width           =   975
      VariousPropertyBits=   612386843
      MaxLength       =   6
      DisplayStyle    =   3
      Size            =   "1720;556"
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
      Index           =   17
      Left            =   135
      TabIndex        =   59
      Top             =   7905
      Width           =   1245
   End
   Begin MSForms.ComboBox cboTransport 
      Height          =   315
      Left            =   1935
      TabIndex        =   15
      Top             =   7845
      Width           =   975
      VariousPropertyBits=   612386843
      MaxLength       =   6
      DisplayStyle    =   3
      Size            =   "1720;556"
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
      Caption         =   "Line1"
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
      Left            =   6810
      TabIndex        =   58
      Top             =   6975
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
      Index           =   20
      Left            =   9720
      TabIndex        =   57
      Top             =   6975
      Width           =   450
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Packing"
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
      Left            =   135
      TabIndex        =   56
      Top             =   7575
      Width           =   660
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
      Left            =   135
      TabIndex        =   55
      Top             =   6915
      Width           =   1260
   End
   Begin MSForms.ComboBox cboPaymentTerm 
      Height          =   315
      Left            =   1935
      TabIndex        =   12
      Top             =   6855
      Width           =   975
      VariousPropertyBits=   612386843
      MaxLength       =   6
      DisplayStyle    =   3
      Size            =   "1720;556"
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
      Caption         =   "Price Condition"
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
      Left            =   135
      TabIndex        =   54
      Top             =   7245
      Width           =   1290
   End
   Begin MSForms.ComboBox cboPriceCondition 
      Height          =   315
      Left            =   1935
      TabIndex        =   13
      Top             =   7185
      Width           =   975
      VariousPropertyBits=   612386843
      MaxLength       =   6
      DisplayStyle    =   3
      Size            =   "1720;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00A6D2FF&
      Height          =   495
      Left            =   75
      Top             =   8850
      Width           =   15105
   End
   Begin MSForms.ComboBox cboStatus 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2625
      Width           =   1215
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "2143;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cboPONo 
      Height          =   315
      Left            =   2025
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   2625
      Width           =   2775
      VariousPropertyBits=   612386843
      DisplayStyle    =   3
      Size            =   "4895;556"
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
      Caption         =   "PO Date"
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
      Left            =   5910
      TabIndex        =   51
      Top             =   2685
      Width           =   705
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PO No"
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
      Left            =   1425
      TabIndex        =   50
      Top             =   2685
      Width           =   525
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
      Index           =   6
      Left            =   8340
      TabIndex        =   49
      Top             =   2685
      Width           =   1185
   End
   Begin MSForms.ComboBox cboAlarm 
      Height          =   315
      Left            =   11865
      TabIndex        =   10
      Top             =   2625
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
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alarm"
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
      Left            =   11250
      TabIndex        =   48
      Top             =   2685
      Width           =   510
   End
   Begin VB.Label lblFix 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status Fix"
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
      Height          =   195
      Left            =   14085
      TabIndex        =   46
      Top             =   2655
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grand Total"
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
      Left            =   13395
      TabIndex        =   45
      Top             =   8610
      Width           =   1005
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
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
      Left            =   8460
      TabIndex        =   42
      Top             =   8610
      Width           =   1140
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PO No"
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
      Left            =   180
      TabIndex        =   41
      Top             =   8610
      Width           =   525
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PPn"
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
      Left            =   11310
      TabIndex        =   40
      Top             =   8610
      Width           =   315
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00A6D2FF&
      BackStyle       =   1  'Opaque
      Height          =   315
      Left            =   6705
      Top             =   6510
      Width           =   8475
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00A6D2FF&
      BackStyle       =   1  'Opaque
      Height          =   300
      Left            =   75
      Top             =   8565
      Width           =   15105
   End
End
Attribute VB_Name = "frmPOPartsSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sql As String, sqlGrid As String
Dim RS As New ADODB.Recordset
Dim rsGrid As New ADODB.Recordset
Dim i As Long
Dim actrow As Long, isippn As Long
Dim ubah As Boolean, ubahgrid As Boolean, ada As Boolean, adaim As Boolean, statusprice As Boolean
Dim statusfix As String, kodeitem As String
Dim jmlpage As Integer, intpage As Integer
Dim sqlitem As String
Dim RsItem As New ADODB.Recordset
Dim sqlcurstock As String, sqlreq As String, sqlfixord As String, sqlpo As String, sqlrec As String
Dim sqlcomp1 As String, sqlcomp2 As String
Dim rscurstock As New Recordset
Dim rsreq As New Recordset
Dim rsfixord As New Recordset
Dim rsPO As New Recordset
Dim rsrec As New Recordset
Dim kodepo As String
Dim orderawal As Double, comp As Double
Dim rscomp1 As New Recordset
Dim rscomp2 As New Recordset
Dim isipodate As Date, countrycls As Byte, statuskunci As Boolean
Dim tempperiod2 As String, tempdeldate As Date, temptgl As Byte

Dim bteColSelect As Byte
Dim bteColProdCode As Byte
Dim bteColDesc As Byte
Dim bteColUnitCls As Byte
Dim bteColUnit As Byte
Dim bteColQty As Byte
Dim bteColStock As Byte
Dim bteColOrderPoint As Byte
Dim bteColFixOrder As Byte
Dim bteColReq As Byte
Dim bteColMinOrder As Byte ' Add for KAWAI 20090501
Dim bteColSuggestion As Byte
Dim bteColOrder As Byte
Dim bteColEndQty As Byte
Dim bteColLotQty As Byte
Dim bteColCurrCode As Byte
Dim bteColCurr As Byte
Dim bteColPrice As Byte
Dim bteColAmount As Byte
Dim btecolReqNext As Byte 'add by edi 20100920
Dim bteColSafe As Byte
Dim bteColSafePercen As Byte
Dim TotCol As Byte

Public popanggil As String

Dim bteHakPrice As Byte
Dim SeqNoSet As Double
Dim ListSupplier As String

Sub Header()

    bteColSelect = 0
    bteColProdCode = 1
    bteColDesc = 2
    bteColUnitCls = 3
    bteColUnit = 4
    bteColQty = 5
    bteColStock = 6
    bteColOrderPoint = 7
    bteColFixOrder = 8
    bteColReq = 9
    bteColMinOrder = 10
    bteColSuggestion = 11
    bteColOrder = 12
    bteColEndQty = 13
    bteColLotQty = 14
    bteColCurrCode = 15
    bteColCurr = 16
    bteColPrice = 17
    bteColAmount = 18
    btecolReqNext = 19
    bteColSafe = 20
    bteColSafePercen = 21
    TotCol = 22
    
    With grid
        .clear
        
        .Rows = 2
        .ColS = TotCol
        
        .TextMatrix(0, bteColSelect) = " "
        .TextMatrix(0, bteColProdCode) = "Item Code"
        .TextMatrix(0, bteColDesc) = "Description"
        .TextMatrix(0, bteColUnitCls) = "UnitCls"
        .TextMatrix(0, bteColUnit) = "Qty Unit"
        .TextMatrix(0, bteColQty) = "Qty / Box"
        '.TextMatrix(0, bteColStock) = "Current Stock"
        .TextMatrix(0, bteColStock) = "Pre Month Stock"
        .TextMatrix(0, bteColOrderPoint) = "Order Point Qty"
        .TextMatrix(0, bteColFixOrder) = "Fix Order (Receipt Schd)"
        .TextMatrix(0, bteColReq) = "Req"
        .TextMatrix(0, bteColMinOrder) = "Min Order" ' Add for KAWAI 20090501
        .TextMatrix(0, bteColSuggestion) = "Suggestion"
        .TextMatrix(0, bteColOrder) = "Order"
        .TextMatrix(0, bteColEndQty) = "End Stock"
        .TextMatrix(0, bteColLotQty) = "Lot Qty"
        .TextMatrix(0, bteColCurrCode) = "CurrCode"
        .TextMatrix(0, bteColCurr) = "Curr"
        .TextMatrix(0, bteColPrice) = "Price"
        .TextMatrix(0, bteColAmount) = "Amount"
        
        .TextMatrix(1, bteColSelect) = " "
        .TextMatrix(1, bteColProdCode) = "Item Code"
        .TextMatrix(1, bteColDesc) = "Description"
        .TextMatrix(1, bteColUnitCls) = "UnitCls"
        .TextMatrix(1, bteColUnit) = "Qty Unit"
        .TextMatrix(1, bteColQty) = "Qty / Box"
        '.TextMatrix(1, bteColStock) = "Current Stock"
        .TextMatrix(1, bteColStock) = "Pre Month Stock"
        .TextMatrix(1, bteColOrderPoint) = "Order Point Qty"
        .TextMatrix(1, bteColFixOrder) = "Fix Order (Receipt Schd)"
        .TextMatrix(1, bteColReq) = "Req"
        .TextMatrix(1, bteColMinOrder) = "Min Order" ' Add for KAWAI 20090501
        .TextMatrix(1, bteColSuggestion) = "Suggestion"
        .TextMatrix(1, bteColOrder) = "Order"
        .TextMatrix(1, bteColEndQty) = "End Stock"
        .TextMatrix(1, bteColLotQty) = "Lot Qty"
        .TextMatrix(1, bteColCurrCode) = "CurrCode"
        .TextMatrix(1, bteColCurr) = "Curr"
        .TextMatrix(1, bteColPrice) = "Price"
        .TextMatrix(1, bteColAmount) = "Amount"
        
        .ColWidth(bteColSelect) = 300
        .ColWidth(bteColProdCode) = 1600
        .ColWidth(bteColDesc) = 2300
        .ColWidth(bteColUnit) = 700
        .ColWidth(bteColQty) = 1100
        .ColWidth(bteColStock) = 1450
        .ColWidth(bteColOrderPoint) = 1500
        .ColWidth(bteColFixOrder) = 1400
        .ColWidth(bteColReq) = 1000
        .ColWidth(bteColMinOrder) = 1000        ' Add for KAWAI 200090501
        .ColWidth(bteColSuggestion) = 1100
        .ColWidth(bteColOrder) = 1000
        .ColWidth(bteColEndQty) = 1450
        .ColWidth(bteColLotQty) = 1000
        .ColWidth(bteColCurr) = 900
        .ColWidth(bteColPrice) = 2000
        .ColWidth(bteColAmount) = 2500
        
        .ColHidden(bteColUnitCls) = True
        .ColHidden(bteColCurrCode) = True
        .ColHidden(btecolReqNext) = True
        .ColHidden(bteColSafe) = True
        .ColHidden(bteColSafePercen) = True
        
        .ColHidden(bteColCurr) = (bteHakPrice = 0)
        .ColHidden(bteColPrice) = (bteHakPrice = 0)
        .ColHidden(bteColAmount) = (bteHakPrice = 0)
        
        
        .MergeRow(bteColSelect) = True
        .MergeRow(bteColProdCode) = True
        
        
        For i = 0 To bteColAmount
        .MergeCol(i) = True
        Next i
        .MergeCells = flexMergeFixedOnly
        
        .Cell(flexcpAlignment, bteColSelect, bteColSelect, bteColProdCode, bteColAmount) = flexAlignCenterCenter
        .ColAlignment(bteColSelect) = flexAlignCenterCenter
        .ColAlignment(bteColProdCode) = flexAlignLeftCenter
        .ColAlignment(bteColDesc) = flexAlignLeftCenter
        .ColAlignment(bteColUnit) = flexAlignCenterCenter
        For i = bteColQty To bteColLotQty
        .ColAlignment(i) = flexAlignRightCenter
        Next i
        .ColAlignment(bteColCurr) = flexAlignCenterCenter
        .ColAlignment(bteColPrice) = flexAlignRightCenter
        .ColAlignment(bteColAmount) = flexAlignRightCenter
        
        .RowHeight(0) = 225
        .RowHeight(1) = 225
        
        .FrozenCols = bteColUnitCls
        
    End With
    
    jmlpage = 0
    intpage = 0
End Sub

Sub Browse()
Dim p As String
Dim a As Double

    LblErrMsg = ""

    sql = "select * from purchaseorder_master where po_no='" & txtPoNo.Text & "' and sheetcoil_cls=0"
    If RS.State <> adStateClosed Then RS.Close
    RS.Open sql, Db, adOpenKeyset, adLockOptimistic

    If Not (RS.BOF And RS.EOF) Then
        ada = True
        ubah = True

        tempperiod2 = IIf(IsNull(RS("period")), " ", Left(Trim(RS("period")), 4) & "-" & Right(Trim(RS("period")), 2) & "-01")
        tempdeldate = IIf(IsNull(RS("delivery_date")), " ", Trim(RS("delivery_date")))
        
        txtremarks.Text = IIf(IsNull(RS("remarks")), " ", Trim(RS("remarks")))
        statusfix = IIf(IsNull(RS("fix_cls")), 0, RS("fix_cls"))
        
        cboPriceCondition.Text = RS.Fields("PriceCondition_Cls") & ""
        cboPaymentTerm.Text = RS.Fields("PaymentTerm_Cls") & ""
        CboPacking.Text = RS.Fields("POPacking_Cls") & ""
        cboInsuranceCls.Text = RS.Fields("Insurance_Cls") & ""
        cboTransport.Text = RS.Fields("Transportation_Cls") & ""
        
        TxtPOLOT.Text = Trim(RS.Fields("PO_LOT") & "")
        
        txtMarking(0).Text = Trim(RS.Fields("POMarking1") & "")
        txtMarking(1).Text = Trim(RS.Fields("POMarking2") & "")
        txtMarking(2).Text = Trim(RS.Fields("POMarking3") & "")
        txtMarking(3).Text = Trim(RS.Fields("POMarking4") & "")
        txtMarking(4).Text = Trim(RS.Fields("POMarking5") & "")
        txtMarking(5).Text = Trim(RS.Fields("POMarking6") & "")
                
        browseitem
        BrowseGrid
        formatprice
        
        For i = 2 To grid.Rows - 1
            If grid.Cell(flexcpChecked, i, bteColSelect) = flexChecked Then
                a = a + grid.TextMatrix(i, bteColAmount)
            End If
        Next i
        txtamount.Text = Format(a, gs_formatAmount)
        If isippn = 0 Then
            txtPPN.Text = Format(0, gs_formatAmount)
        Else
            txtPPN.Text = Format(CDbl(isippn / 100) * CDbl(txtamount.Text), gs_formatAmount)
        End If
        txtGrandTotal = Format(CDbl(txtPPN.Text) + CDbl(txtamount.Text), gs_formatAmount)

        If statusfix = 1 Then
            kunci (True)
        Else
            kunci (False)
        End If

    Else
        ada = False
    End If

End Sub

Sub BrowseGrid()
    Dim R As Double, g As Double, Pos As Double
    sqlGrid = "select * from purchaseorder_detail where po_no='" & txtPoNo.Text & "' order by item_code"
    If rsGrid.State <> adStateClosed Then rsGrid.Close
    rsGrid.Open sqlGrid, Db, adOpenKeyset, adLockOptimistic

    R = 1
    Pos = 1
    With grid
    Do While Not rsGrid.EOF
        
        For g = 2 To .Rows - 1
            If .TextMatrix(g, 1) = Trim(rsGrid("Item_Code")) Then
                
                .Cell(flexcpChecked, g, bteColSelect) = flexChecked
                .TextMatrix(g, bteColOrder) = Format(Val(rsGrid("qty") & ""), gs_formatQty)
                
                If (Format(tempperiod2, "MMM yyyy") <> Format(dtpPeriod.Value, "MMM yyyy")) Or (Format(tempdeldate, "01 MMM yyyy") <> Format(dtpDeliveryDate.Value, "01 MMM yyyy")) Then
                    If Year(dtpPeriod) = Year(dtpDeliveryDate) And Month(dtpPeriod) = Month(dtpDeliveryDate) Then
                        .TextMatrix(g, bteColFixOrder) = Format((CDbl(.TextMatrix(g, bteColFixOrder)) + CDbl(.TextMatrix(g, bteColOrder))), gs_formatQty)
                    ElseIf Format(dtpPeriod, "yyyy-mm-01") > Format(dtpDeliveryDate, "yyyy-mm-01") Then
                        .TextMatrix(g, bteColStock) = Format((CDbl(.TextMatrix(g, bteColStock)) + CDbl(.TextMatrix(g, bteColOrder))), gs_formatQty)
                    ElseIf Format(dtpPeriod, "yyyy-mm-01") < Format(dtpDeliveryDate, "yyyy-mm-01") Then
                    
                    End If
                    .TextMatrix(g, bteColEndQty) = Format((CDbl(.TextMatrix(g, bteColStock)) + CDbl(.TextMatrix(g, bteColFixOrder)) - CDbl(.TextMatrix(g, bteColReq))), gs_formatQty)
                Else
                    'Penambahan utk menghitung End Stock
                    .TextMatrix(g, bteColFixOrder) = Format(CDbl(.TextMatrix(g, bteColFixOrder)), gs_formatQty)
                    .TextMatrix(g, bteColEndQty) = Format((CDbl(.TextMatrix(g, bteColStock)) + CDbl(.TextMatrix(g, bteColFixOrder)) - CDbl(.TextMatrix(g, bteColReq))), gs_formatQty)
                End If
                
                If Format(tempdeldate, "dd MMM yyyy") = Format(dtpDeliveryDate.Value, "dd MMM yyyy") Then
                    .TextMatrix(g, bteColUnitCls) = Trim(rsGrid("Unit_cls"))
                    .TextMatrix(g, bteColUnit) = uf_GetUnitDescription(Trim(rsGrid("Unit_Cls")))
                    
                    .TextMatrix(g, bteColCurrCode) = Trim(rsGrid("currency_code"))
                    .TextMatrix(g, bteColCurr) = uf_GetCurrencyDescription(Trim(rsGrid("Currency_code")))
                    .TextMatrix(g, bteColPrice) = Format(Val(rsGrid("price")), gs_formatPrice)

                    .TextMatrix(g, bteColAmount) = IIf(IsNull(rsGrid("amount")), 0, Trim(rsGrid("amount")))
                    If .TextMatrix(g, bteColAmount) <> 0 Then .TextMatrix(g, bteColAmount) = Format(.TextMatrix(g, bteColAmount), gs_formatAmount)
                Else
                    .TextMatrix(g, bteColAmount) = IIf(.TextMatrix(g, bteColPrice) = "", 0, Format(CDbl(.TextMatrix(g, bteColOrder)) * CDbl(.TextMatrix(g, bteColPrice)), gs_formatAmount))
                End If
                
                Pos = Pos + 1
                .RowPosition(g) = Pos
                R = R + 1
            End If
        Next g

        rsGrid.MoveNext
    Loop
    End With
End Sub

Sub browseitem()
    
    Header
    If ubah = False Then
        txtamount.Text = Format(0, gs_formatAmount)
        txtPPN.Text = Format(0, gs_formatAmount)
        txtGrandTotal.Text = Format(0, gs_formatAmount)
    End If
    kodeitem = ""
    adaim = False
    i = 2
   
    
    Call Item(cboSupplier.Text, 1)
'    Call Item(cboSupplier.Text, 0)
'    Call Item("000000", 1)
'    Call Item("000000", 0)
    
    adaim = True
    Call Item(cboSupplier.Text)
    
    jmlpage = Int((grid.Rows - 2) / 16) + 1
    If grid.Rows > 2 Then
        intpage = 1
    Else
        intpage = 0
    End If
    
End Sub

Sub browseprice()
Dim sql2 As String
Dim rs2 As New Recordset

'    sql2 = "select trade_code, priority_cls, currency_code, price, unit_cls from price_master where " & _
'           "item_code='" & grid.TextMatrix(actrow, bteColProdCode) & "' and price_cls='01' and (trade_code='" & cboSupplier.Text & _
'           "' or trade_code='000000') and start_date<='" & Format(dtpDeliveryDate.Value, "yyyymmdd") & "' and end_date>='" & _
'           Format(dtpDeliveryDate.Value, "yyyymmdd") & "' order by trade_code desc, priority_cls desc"
    
    sql2 = "select trade_code, priority_cls, currency_code, price, unit_cls from price_master where " & _
           "item_code='" & grid.TextMatrix(actrow, bteColProdCode) & "' and price_cls='01' and (trade_code='" & cboSupplier.Text & _
           "' or trade_code='000000') and month(start_date)='" & Month(dtpPeriod) & "' and Year(start_date)='" & _
           Year(dtpPeriod) & "' order by trade_code desc, priority_cls desc"
    
    Set rs2 = Db.Execute(sql2)


    With cboprice
        .clear
        .columnCount = 4
        .ColumnWidths = "70pt;70pt;0pt;0pt"
        .ListWidth = 140
        .ListRows = 4

        i = 0
        Do While Not rs2.EOF
            .AddItem
            .List(i, 0) = Format(Trim(rs2("price")), gs_formatPrice)
            If rs2("trade_code") = "000000" Then
              .List(i, 1) = "Common " & Trim(rs2("priority_cls"))
            Else
              .List(i, 1) = "Priority " & Trim(rs2("priority_cls"))
            End If
            .List(i, 2) = Trim(rs2("Currency_Code"))
            .List(i, 3) = Trim(rs2("unit_cls"))

            rs2.MoveNext
            i = i + 1

        Loop
    End With
End Sub

Function Item(ByVal C As String, Optional ByVal p As Integer)
Dim sqlitdesc As String, sqlinvcon As String
Dim cs As Double
Dim rsinvcon As New Recordset
Dim tempnow As Date, closingmonth As Date, tempdtpPeriod As Date
Dim moq, spq, req, reqN, lastMth, safe, safePer As Double
Dim temptgl2 As String

tempnow = Format(Now, "yyyy-mm") & "-01"
tempdtpPeriod = Format(dtpPeriod, "yyyy-mm-dd")
        
sqlinvcon = "select * from inventory_control where fix_cls=1"
If rsinvcon.State <> adStateClosed Then rsinvcon.Close
rsinvcon.Open sqlinvcon, Db, adOpenKeyset, adLockOptimistic
If Not (rsinvcon.BOF And rsinvcon.EOF) Then
    rsinvcon.MoveLast
    closingmonth = Trim(rsinvcon("inventory_year")) & "-" & Format(Trim(rsinvcon("inventory_month")), "0#") & "-01"
End If

    temptgl2 = DateAdd("m", 1, dtpPeriod.Value)
    If adaim = False Then
' command by edi 20100920
        sqlitem = "select *, (curstock + fixorder - requirement) endstock " & _
                    vbLf & "From ( " & _
                    vbLf & "select price_master.item_code, price_master.trade_code, price_master.priority_cls, price_master.unit_cls ,currency_code, isnull(price_master.price,0) Price, " & _
                    vbLf & "item_master.item_name , finishgoodpart_cls, number_entering, number_box, lot_qty, orderpoint_qty, MinOrder, control_cls " & _
                    vbLf & ", isnull( " & _
                    vbLf & "(select sisaPOQty from " & _
                    vbLf & "( select item_code, sum(sisaQty)SisaPoQty from " & _
                    vbLf & "    (select pr.qty recQty,SisaQty =Pod.qty--case when pod.qty-isnull(pr.Qty,0)<0 then 0 " & _
                    vbLf & "        --else pod.qty-isnull(pr.Qty,0) end ,pod.* " & _
                    vbLf & "        ,pod.* " & _
                    vbLf & "        from purchaseOrder_detail pod left join purchaseOrder_master pom on pod.po_no=pom.po_no " & _
                    vbLf & "        left join ( " & _
                    vbLf & "        select po_no,item_code,sum(case receipt_cls when 'R1' then -qty else qty end)qty " & _
                    vbLf & "        from part_receipt pr group by po_no,item_code " & _
                    vbLf & "        ) pr " & _
                    vbLf & "        on pod.po_no=pr.po_no and pod.item_code=pr.item_code " & _
                    vbLf & "        where --(pod.complete_cls<>'1' or pod.complete_cls is null ) and " & _
                    vbLf & "             year(pod.delivery_date)='" & Year(dtpPeriod.Value) & "' and month(pod.delivery_date)='" & Month(dtpPeriod.Value) & "' "
 
        If (Format(tempperiod2, "MMM yyyy") <> Format(dtpPeriod.Value, "MMM yyyy")) Or (Format(tempdeldate, "01 MMM yyyy") <> Format(dtpDeliveryDate.Value, "01 MMM yyyy")) Then _
            sqlitem = sqlitem & " and pom.po_no<>'" & Trim(txtPoNo.Text) & "' "
                
        sqlitem = sqlitem & "    ) tbE group by item_code " & _
                        vbLf & ")tbF where tbF.item_code=price_master.item_code) " & _
                        vbLf & ",0) as fixorder " & _
                        vbLf & ", isnull( " & _
                        vbLf & "(select sisaReqQty from " & _
                        vbLf & "(select childItem_code, sum(sisaReqQty)sisaReqQty " & _
                        vbLf & "  from ( " & _
                        vbLf & "    select childItem_code,sum(childRequirement_qty)plans,sum(childRequirementResult_qty)Results, " & _
                        vbLf & "        case when sum(childRequirement_qty)-sum(childRequirementResult_qty)-sum(offchildrequirement_qty)<0 then 0 else " & _
                        vbLf & "        Sum (childRequirement_qty) - Sum(childRequirementResult_qty)-sum(offchildrequirement_qty) end As SisaReqQty "

        sqlitem = sqlitem & "From requirement " & _
                        vbLf & "    where year(childrequirement_date)='" & Year(dtpPeriod.Value) & "' and month(childrequirement_date)='" & Month(dtpPeriod.Value) & "' " & _
                        vbLf & "    and (complete_cls is null or complete_cls<>'1') " & _
                        vbLf & "    group by parentitem_code, factory_code, line_code, production_date, " & _
                        vbLf & "    cast(year(childrequirement_date) as varchar(4))+'-'+cast(month(childrequirement_date)as varchar(4)) ,childItem_code " & _
                        vbLf & "    )tbC group by childItem_code " & _
                        vbLf & ")tbD where tbD.childitem_code=price_master.item_code) " & _
                        vbLf & ",0) requirement " & vbLf
    
       sqlitem = sqlitem & ", isnull( " & _
                        vbLf & "(select sisaReqQty from " & _
                        vbLf & "(select childItem_code, sum(sisaReqQty)sisaReqQty " & _
                        vbLf & "  from ( " & _
                        vbLf & "    select childItem_code,sum(childRequirement_qty)plans,sum(childRequirementResult_qty)Results, " & _
                        vbLf & "        case when sum(childRequirement_qty)-sum(childRequirementResult_qty)-sum(offchildrequirement_qty)<0 then 0 else " & _
                        vbLf & "        Sum (childRequirement_qty) - Sum(childRequirementResult_qty)-sum(offchildrequirement_qty) end As SisaReqQty " & _
                        vbLf & "From requirement " & _
                        vbLf & "    where year(childrequirement_date)='" & Year(temptgl2) & "' and month(childrequirement_date)='" & Month(temptgl2) & "' " & _
                        vbLf & "    and (complete_cls is null or complete_cls<>'1') " & _
                        vbLf & "    group by parentitem_code, factory_code, line_code, production_date, " & _
                        vbLf & "    cast(year(childrequirement_date) as varchar(4))+'-'+cast(month(childrequirement_date)as varchar(4)) ,childItem_code " & _
                        vbLf & "    )tbC group by childItem_code " & _
                        vbLf & ")tbD where tbD.childitem_code=price_master.item_code) " & _
                        vbLf & ",0) requirementNextMonth "
                        
            sqlitem = sqlitem & _
                        vbLf & ", isnull( " & _
                        vbLf & "(select isnull(stockMaster_stock,0) + isnull(tbPO.sisaPOqty,0) - isnull(tbReq.sisaReqQty,0) + isnull(tbPOAC.POQty,0) StockMaster_Stock1 " & _
                        vbLf & "From item_master " & _
                        vbLf & "Left Join " & _
                        vbLf & "( select isnull(case when datediff(month,ClosingDate,StartDate)=0 then sum(lm_premonth) " & _
                        vbLf & "   when datediff(month,ClosingDate,StartDate) >=1 then sum(tm_premonth) " & _
                        vbLf & "   --when datediff(month,ClosingDate,StartDate) >=2 then sum(nm_premonth) " & _
                        vbLf & "   end,0) StockMaster_Stock,ClosingDate,Item_code ,startDate " & _
                        vbLf & "  From " & _
                        vbLf & "  (select " & _
                        vbLf & "      (select cast (cast(year as varchar(4) ) +case when month <10 then '0' else'' end +cast (month as varchar(2) )+'01' " & _
                        vbLf & "       as dateTime)ClosingDate " & _
                        vbLf & "       from ( select top 1 max(inventory_month)month,inventory_year year  from inventory_control " & _
                        vbLf & "              where fix_cls='1' group by inventory_year  order by inventory_year desc )tbA " & _
                        vbLf & "      )ClosingDate,StartDate='" & Format(tempdtpPeriod, "yyyy-mm-dd") & "',SM.*  " & _
                        vbLf & "   from stock_master SM "
        
        'Jika NG Cls = No Maka tidak diperhitungkan
        sqlitem = sqlitem & _
                        vbLf & "   left join Warehouse_Master WM " & _
                        vbLf & "        ON SM.Warehouse_Code = WM.WH_Code " & _
                        vbLf & "   left join (Select Trade_Code,isnull(NG_Cls,0) NG_Cls from Trade_Master where trade_cls = '1') TM " & _
                        vbLf & "        ON SM.Warehouse_Code = TM.Trade_Code " & _
                        vbLf & "   Where WM.NG_Cls = '02' or TM.NG_Cls = 0 "
        
        sqlitem = sqlitem & _
                        vbLf & " )tbA " & _
                        vbLf & " group by ClosingDate,Item_code,StartDate " & _
                        vbLf & " )tbStock on tbstock.item_code=item_master.item_code " & _
                        vbLf & "Left Join " & _
                        vbLf & "( select * from " & _
                        vbLf & "   ( select item_code,sum(sisaQty)SisaPoQty from " & _
                        vbLf & "    ( select pr.qty recQty, SisaQty = case when pod.qty-isnull(pr.Qty,0)<0 then 0 " & _
                        vbLf & "        else pod.qty-isnull(pr.Qty,0) end,pod.* " & _
                        vbLf & "        from purchaseOrder_detail pod left join purchaseOrder_master pom on pod.po_no=pom.po_no " & _
                        vbLf & "        left join ( " & _
                        vbLf & "        select po_no,item_code,sum(case receipt_cls when 'R1' then -qty else qty end)qty " & _
                        vbLf & "        from part_receipt pr group by po_no,item_code ) pr " & _
                        vbLf & "        on pod.po_no=pr.po_no and pod.item_code=pr.item_code " & _
                        vbLf & "        --where (pod.complete_cls<>'1' or pod.complete_cls is null ) and pod.delivery_date >='" & Format(CDate(closingmonth), "yyyy-mm-dd") & "' " & _
                        vbLf & "        where (pod.complete_cls<>'1' or pod.complete_cls is null ) and pod.delivery_date <'" & Format(DateAdd("m", 1, CDate(closingmonth)), "yyyy-mm-dd") & "' " & _
                        vbLf & "        --and pod.delivery_date <'" & Format(tempdtpPeriod, "yyyy-mm-dd") & "' " & vbLf
                
        If (Format(tempperiod2, "MMM yyyy") <> Format(dtpPeriod.Value, "MMM yyyy")) Or (Format(tempdeldate, "01 MMM yyyy") <> Format(dtpDeliveryDate.Value, "01 MMM yyyy")) Then _
            sqlitem = sqlitem & " and pom.po_no<>'" & Trim(txtPoNo.Text) & "' "
            
        sqlitem = sqlitem & "    ) tbA group by item_code " & _
                        vbLf & "   )tbB where tbB.item_code=item_master.item_code " & _
                        vbLf & ")tbPo on tbPo.item_code=item_master.item_code " & vbLf

    'PO Qty setelah closing diperhitungkan utuh
        sqlitem = sqlitem & "Left Join " & _
                  vbLf & "  ( select  pod.item_code, sum(pod.qty) POqty     " & _
                  vbLf & "    from purchaseOrder_detail pod left join purchaseOrder_master pom on pod.po_no=pom.po_no " & _
                  vbLf & "    where pod.delivery_date >='" & Format(DateAdd("m", 1, CDate(closingmonth)), "yyyy-mm-dd") & "' " & _
                  vbLf & "           and pod.delivery_date <'" & Format(tempdtpPeriod, "yyyy-mm-dd") & "' " & _
                  vbLf & "    group by pod.item_code                     " & _
                  vbLf & "  )tbPOAC on tbPoAC.item_code=item_master.item_code "

          sqlitem = sqlitem & _
                        vbLf & "Left Join " & _
                        vbLf & "( select childItem_code, sum(sisaReqQty)sisaReqQty " & _
                        vbLf & "  from ( select childItem_code,sum(childRequirement_qty)plans,sum(childRequirementResult_qty)Results, " & _
                        vbLf & "     case when sum(childRequirement_qty)-sum(childRequirementResult_qty)-sum(offchildrequirement_qty)<0 then 0 else " & _
                        vbLf & "         sum(childRequirement_qty)-sum(childRequirementResult_qty)-sum(offchildrequirement_qty) end as SisaReqQty " & _
                        vbLf & "     from requirement where childrequirement_date >='" & Format(DateAdd("m", 1, CDate(closingmonth)), "yyyy-mm-dd") & "' and childrequirement_date <'" & Format(tempdtpPeriod, "yyyy-mm-dd") & "' " & _
                        vbLf & "     and (complete_cls is null or complete_cls<>'1') " & _
                        vbLf & "     group by parentitem_code, factory_code, line_code, production_date, " & _
                        vbLf & "     cast(year(childrequirement_date) as varchar(4))+'-'+cast(month(childrequirement_date)as varchar(4)), childItem_code " & _
                        vbLf & "    )tbA group by childItem_code " & _
                        vbLf & ")tbReq on item_master.item_code=tbReq.childItem_code "
                        
         sqlitem = sqlitem & "where item_master.item_code=price_master.item_code ) " & _
                       vbLf & ",0) curstock, " & _
                       vbLf & "isnull(safety_stock,0)safety_stock,isnull(safety_stock_percentage,0)safety_stock_percentage" & _
                       vbLf & "From price_master " & _
                       vbLf & "inner join (Select * from [PO_Set_Master] Where Parent_SetCode = '" & Trim(CboSetCode) & "') AS Set_master on price_master.item_code=Set_master.item_code " & _
                       vbLf & "inner join item_master on set_master.item_code=item_master.item_code " & _
                       vbLf & "Left Join (Select Trade_Code, Item_Code, Price_Cls, Priority_cls, Price from Price_Master where Priority_Cls = 1 and Left(start_date,6)<='201101' And Left(End_Date,6)>='201101') Price1 ON Price_Master.Item_Code = Price1.Item_Code and Price_Master.Price_Cls = Price1.Price_Cls and price_master.trade_code = price1.trade_code " & _
                       vbLf & "Left join (Select Trade_Code, Item_Code, Price_Cls, Priority_cls, Price from Price_Master where Priority_Cls = 0 and Left(start_date,6)<='201101' And Left(End_Date,6)>='201101') Price0 ON Price_Master.Item_Code = Price0.Item_Code and Price_Master.Price_Cls = Price0.Price_Cls and price_master.trade_code = price0.trade_code "

    ' Price Berdasarkan Periode (Kawai -20100116)
    sqlitem = sqlitem & vbLf & " where (Price_master.trade_code='" & C & "')  And " & _
                        vbLf & "        Left(start_date,6)<='" & Format(dtpPeriod, "yyyyMM") & "' And Left(End_Date,6)>='" & Format(dtpPeriod, "yyyyMM") & "' " & _
                    vbLf & "and Price_master.price_cls='01' and price_master.priority_cls= coalesce(Price1.Priority_cls,Coalesce(Price0.Priority_cls,'0')) and price_master.item_code not in ('" & kodeitem & "') " & _
                    vbLf & "--and (rtrim(sheetcoil_cls) is null or rtrim(sheetcoil_cls)='') " & vbLf

    ' Add Material Filter For Kawai 20090421
    If CboMat <> strAll Then sqlitem = sqlitem & " And Item_Master.Material_Cls='" & CboMat & "'"
    sqlitem = sqlitem & ") PO "
 
    If cboAlarm.Text = "Yes" Then _
    sqlitem = sqlitem & vbLf & " Where (curstock + fixorder ) < (case control_cls when '03' then orderpoint_qty else 0 end) "
    sqlitem = sqlitem & vbLf & " order by item_code, trade_code desc , priority_cls desc"

    Else
        sqlitem = "select *, (curstock + fixorder - requirement) endstock " & _
                  vbLf & "From ( " & _
                  vbLf & "      select item_master.item_code, supplier_code, unit_cls, item_master.item_name, finishgoodpart_cls, number_entering, number_box, lot_qty, orderpoint_qty, MinOrder, control_cls " & _
                  vbLf & "      , isnull( " & _
                  vbLf & "                (select sisaPOQty  " & _
                  vbLf & "                 from   ( select item_code, sum(sisaQty)SisaPoQty  " & _
                  vbLf & "                          from (select pr.qty recQty,SisaQty =case when pod.qty-isnull(pr.Qty,0)<0 then 0 " & _
                  vbLf & "                                      else pod.qty-isnull(pr.Qty,0) end ,pod.* " & _
                  vbLf & "                                from purchaseOrder_detail pod " & _
                  vbLf & "                                left join purchaseOrder_master pom on pod.po_no=pom.po_no " & _
                  vbLf & "                                Left Join ( " & _
                  vbLf & "                                            select po_no,item_code,sum(case receipt_cls when 'R1' then -qty else qty end)qty " & _
                  vbLf & "                                            from part_receipt pr group by po_no,item_code " & _
                  vbLf & "                                           ) pr " & _
                  vbLf & "                                       on pod.po_no=pr.po_no and pod.item_code=pr.item_code " & _
                  vbLf & "                                where (pod.complete_cls<>'1' or pod.complete_cls is null ) " & _
                  vbLf & "                                       and year(pod.delivery_date)='" & Year(dtpPeriod.Value) & "' and month(pod.delivery_date)='" & Month(dtpPeriod.Value) & "' "

        If (Format(tempperiod2, "MMM yyyy") <> Format(dtpPeriod.Value, "MMM yyyy")) Or (Format(tempdeldate, "01 MMM yyyy") <> Format(dtpDeliveryDate.Value, "01 MMM yyyy")) Then _
            sqlitem = sqlitem & " and pom.po_no<>'" & Trim(txtPoNo.Text) & "' "
                  
        sqlitem = sqlitem & "                             ) tbE group by item_code " & _
                 vbLf & "                           )tbF where tbF.item_code=item_master.item_code) " & _
                 vbLf & "               ,0) as fixorder " & _
                 vbLf & "       , isnull( " & _
                  vbLf & "               (select sisaReqQty  " & _
                  vbLf & "                from  (select childItem_code, sum(sisaReqQty)sisaReqQty " & _
                  vbLf & "      from ( Select ACC.*From ( " & _
                  vbLf & "      select parentitem_code, childItem_code,sum(childRequirement_qty)plans,sum(childRequirementResult_qty)Results, " & _
                  vbLf & "          case when sum(childRequirement_qty)-sum(childRequirementResult_qty)-sum(offchildrequirement_qty)<0 then 0 else " & _
                  vbLf & "          Sum (childRequirement_qty) - Sum(childRequirementResult_qty)-sum(offchildrequirement_qty) "
        sqlitem = sqlitem & "  end As SisaReqQty " & _
                  vbLf & "      From requirement " & _
                  vbLf & "      where year(childrequirement_date)='" & Year(dtpPeriod.Value) & "' and month(childrequirement_date)='" & Month(dtpPeriod.Value) & "' " & _
                  vbLf & "      and (complete_cls is null or complete_cls<>'1') " & _
                  vbLf & "      group by parentitem_code, factory_code, line_code, production_date, " & _
                  vbLf & "      cast(year(childrequirement_date) as varchar(4))+'-'+cast(month(childrequirement_date)as varchar(4)) , childItem_code )ACC " & _
                  vbLf & "      Inner join BOM_Master BMM on BMM.Parent_ItemCode = ACC.ParentItem_Code and BMM.Item_Code=ACC.ChildItem_Code " & _
                  vbLf & "      )tbC group by childItem_code--, parentitem_code, " & _
                  vbLf & "  )tbD where tbD.childitem_code=item_master.item_code) " & _
                  vbLf & "  ,0) requirement "
                  
         sqlitem = sqlitem & "       , isnull( " & _
                  vbLf & "               (select sisaReqQty  " & _
                  vbLf & "                from  (select childItem_code, sum(sisaReqQty)sisaReqQty " & _
                  vbLf & "      from ( Select ACC.*From ( " & _
                  vbLf & "      select parentitem_code, childItem_code,sum(childRequirement_qty)plans,sum(childRequirementResult_qty)Results, " & _
                  vbLf & "          case when sum(childRequirement_qty)-sum(childRequirementResult_qty)-sum(offchildrequirement_qty)<0 then 0 else " & _
                  vbLf & "          Sum (childRequirement_qty) - Sum(childRequirementResult_qty)-sum(offchildrequirement_qty) " & _
                  vbLf & "  end As SisaReqQty " & _
                  vbLf & "      From requirement " & _
                  vbLf & "      where year(childrequirement_date)='" & Year(temptgl2) & "' and month(childrequirement_date)='" & Month(temptgl2) & "' " & _
                  vbLf & "      and (complete_cls is null or complete_cls<>'1') " & _
                  vbLf & "      group by parentitem_code, factory_code, line_code, production_date, " & _
                  vbLf & "      cast(year(childrequirement_date) as varchar(4))+'-'+cast(month(childrequirement_date)as varchar(4)) , childItem_code )ACC " & _
                  vbLf & "      Inner join BOM_Master BMM on BMM.Parent_ItemCode = ACC.ParentItem_Code and BMM.Item_Code=ACC.ChildItem_Code "
        sqlitem = sqlitem & "      )tbC group by childItem_code--, parentitem_code, " & _
                  vbLf & "  )tbD where tbD.childitem_code=item_master.item_code) " & _
                  vbLf & "  ,0) requirementNextMonth "
 
        sqlitem = sqlitem & vbLf & "  , isnull( " & _
                  vbLf & "  (select isnull(stockMaster_stock,0) + isnull(tbPO.sisaPOqty,0) - isnull(tbReq.sisaReqQty,0) + isnull(tbPOAC.POQty,0) StockMaster_Stock1 " & _
                  vbLf & "  from item_master im " & _
                  vbLf & "  Left Join " & _
                  vbLf & "  ( select isnull(case when datediff(month,ClosingDate,StartDate)=0 then sum(lm_premonth) " & _
                  vbLf & "     when datediff(month,ClosingDate,StartDate) =1 then sum(tm_premonth) " & _
                  vbLf & "     when datediff(month,ClosingDate,StartDate) >=2 then sum(nm_premonth) " & _
                  vbLf & "     end,0) StockMaster_Stock,ClosingDate,Item_code ,startDate " & _
                  vbLf & "    From " & _
                  vbLf & "    (select " & _
                  vbLf & "        (select cast (cast(year as varchar(4) ) +case when month <10 then '0' else'' end +cast (month as varchar(2) )+'01' " & _
                  vbLf & "         as dateTime)ClosingDate " & _
                  vbLf & "         from ( select top 1 max(inventory_month)month,inventory_year year  from inventory_control " & _
                  vbLf & "                where fix_cls='1' group by inventory_year  order by inventory_year desc )tbA " & _
                  vbLf & "        )ClosingDate,StartDate='" & Format(tempdtpPeriod, "yyyy-mm-dd") & "',SM.*  " & vbLf
                  
        'Jika NG Cls = No Maka tidak diperhitungkan
        sqlitem = sqlitem & _
                  vbLf & "    from stock_master SM " & _
                  vbLf & "    left join Warehouse_master WM " & _
                  vbLf & "         ON SM.Warehouse_Code = WM.WH_Code " & _
                  vbLf & "    left join (Select Trade_Code,isnull(NG_Cls,0) NG_Cls from Trade_Master where trade_cls = '1') TM " & _
                  vbLf & "          ON SM.Warehouse_Code = TM.Trade_Code " & _
                  vbLf & "    Where WM.NG_Cls = '02' or TM.NG_Cls = 0 " & vbLf
                  
        sqlitem = sqlitem & " )tbA " & _
                  vbLf & " group by ClosingDate,Item_code,StartDate " & _
                  vbLf & "  )tbStock on tbstock.item_code=im.item_code " & _
                  vbLf & "  Left Join " & _
                  vbLf & "  ( select * from " & _
                  vbLf & "     ( select item_code,sum(sisaQty)SisaPoQty from " & _
                  vbLf & "      ( select pr.qty recQty, SisaQty = case when pod.qty-isnull(pr.Qty,0)<0 then 0 " & _
                  vbLf & "          else pod.qty-isnull(pr.Qty,0) end,pod.* " & _
                  vbLf & "          from purchaseOrder_detail pod left join purchaseOrder_master pom on pod.po_no=pom.po_no " & _
                  vbLf & "          left join ( " & _
                  vbLf & "          select po_no,item_code,sum(case receipt_cls when 'R1' then -qty else qty end)qty " & _
                  vbLf & "          from part_receipt pr group by po_no,item_code ) pr " & _
                  vbLf & "          on pod.po_no=pr.po_no and pod.item_code=pr.item_code " & _
                  vbLf & "          --where (pod.complete_cls<>'1' or pod.complete_cls is null ) and pod.delivery_date >='" & Format(CDate(closingmonth), "yyyy-mm-dd") & "' " & _
                  vbLf & "          where (pod.complete_cls<>'1' or pod.complete_cls is null ) and pod.delivery_date <'" & Format(DateAdd("m", 1, CDate(closingmonth)), "yyyy-mm-dd") & "' " & _
                  vbLf & "          --and pod.delivery_date <'" & Format(tempdtpPeriod, "yyyy-mm-dd") & "' " & vbLf ' Line hanya PO setelah closing yg dihitung sisanya
                  
        If (Format(tempperiod2, "MMM yyyy") <> Format(dtpPeriod.Value, "MMM yyyy")) Or (Format(tempdeldate, "01 MMM yyyy") <> Format(dtpDeliveryDate.Value, "01 MMM yyyy")) Then _
            sqlitem = sqlitem & " and pom.po_no<>'" & Trim(txtPoNo.Text) & "' "
                  
        sqlitem = sqlitem & ") tbA group by item_code " & _
                  vbLf & "     )tbB where tbB.item_code=item_master.item_code " & _
                  vbLf & "  )tbPo on tbPo.item_code=im.item_code " & vbLf
                                    
        'PO Qty setelah closing diperhitungkan utuh
        sqlitem = sqlitem & "Left Join " & _
                  vbLf & "  ( select  pod.item_code, sum(pod.qty) POqty     " & _
                  vbLf & "    from purchaseOrder_detail pod left join purchaseOrder_master pom on pod.po_no=pom.po_no " & _
                  vbLf & "    where pod.delivery_date >='" & Format(DateAdd("m", 1, CDate(closingmonth)), "yyyy-mm-dd") & "' " & _
                  vbLf & "           and pod.delivery_date <'" & Format(tempdtpPeriod, "yyyy-mm-dd") & "' " & _
                  vbLf & "    group by pod.item_code                     " & _
                  vbLf & "  )tbPOAC on tbPoAC.item_code=item_master.item_code "

        sqlitem = sqlitem & _
                  vbLf & "  Left Join " & _
                  vbLf & "  ( select childItem_code, sum(sisaReqQty)sisaReqQty " & _
                  vbLf & "    from ( select childItem_code,sum(childRequirement_qty)plans,sum(childRequirementResult_qty)Results, " & _
                  vbLf & "       case when sum(childRequirement_qty)-sum(childRequirementResult_qty)-sum(offchildrequirement_qty)<0 then 0 else " & _
                  vbLf & "           sum(childRequirement_qty)-sum(childRequirementResult_qty)-sum(offchildrequirement_qty) end as SisaReqQty " & _
                  vbLf & "       from requirement where childrequirement_date >='" & Format(DateAdd("m", 1, CDate(closingmonth)), "yyyy-mm-dd") & "' and childrequirement_date <'" & Format(tempdtpPeriod, "yyyy-mm-dd") & "' " & _
                  vbLf & "       and (complete_cls is null or complete_cls<>'1') " & _
                  vbLf & "       group by parentitem_code, factory_code, line_code, production_date, " & _
                  vbLf & "       cast(year(childrequirement_date) as varchar(4))+'-'+cast(month(childrequirement_date)as varchar(4)), childItem_code " & _
                  vbLf & "      )tbA group by childItem_code "
                  
        sqlitem = sqlitem & ")tbReq on im.item_code=tbReq.childItem_code " & _
                  vbLf & "  where im.item_code=item_master.item_code ) " & _
                  vbLf & "  ,0) curstock, " & _
                  vbLf & "isnull(safety_stock,0)safety_stock,isnull(safety_stock_percentage,0)safety_stock_percentage" & _
                  vbLf & "from item_master " & _
                  vbLf & "inner join (Select * from [PO_Set_Master] Where Parent_SetCode = 'A001') AS Set_master on item_master.item_code=Set_master.item_code  " & _
                  vbLf & "where supplier_code='" & C & "' and " & _
                  vbLf & "      item_master.item_code not in ('" & kodeitem & "')  " & _
                  vbLf & "      --and (rtrim(sheetcoil_cls) is null or rtrim(sheetcoil_cls)='') " & vbLf
                  
    ' Add Material Filter For Kawai 20090421
    If CboMat <> strAll Then sqlitem = sqlitem & " And Item_Master.Material_Cls='" & CboMat & "'"
    sqlitem = sqlitem & ") PO "

    
    If cboAlarm.Text = "Yes" Then _
    sqlitem = sqlitem & " Where (curstock + fixorder) < (case control_cls when '03' then orderpoint_qty else 0 end) "
    sqlitem = sqlitem & " order by PO.item_code"

    End If
        
    Set RsItem = Db.Execute(sqlitem)
    
    If Not (RsItem.BOF And RsItem.EOF) Then
    With grid
    Do While Not RsItem.EOF
        
        .Rows = .Rows + 1
        
        If i > 2 Then
            kodeitem = kodeitem & "','" & Trim(RsItem("Item_Code"))
        Else
            kodeitem = Trim(RsItem("Item_Code"))
        End If
        .TextMatrix(i, bteColProdCode) = Trim(RsItem("Item_Code"))
        
        .TextMatrix(i, bteColDesc) = IIf(IsNull(RsItem("item_name")), "", Trim(RsItem("item_name")))
        If RsItem("finishgoodpart_cls") = "01" Then
            .TextMatrix(i, bteColQty) = Format(Val(RsItem("number_entering") & ""), gs_formatBox)
            spq = .TextMatrix(i, bteColQty)
        Else
            .TextMatrix(i, bteColQty) = Format(Val(RsItem("number_box") & ""), gs_formatBox)
            spq = .TextMatrix(i, bteColQty)
        End If
        
        .TextMatrix(i, bteColOrderPoint) = Format(Val(RsItem("orderpoint_qty") & ""), gs_formatQty)
        .TextMatrix(i, bteColMinOrder) = Format(Val(RsItem("MinOrder") & ""), gs_formatQty) 'Add for KAWAI 20090501
        moq = .TextMatrix(i, bteColMinOrder)
        
        .TextMatrix(i, bteColLotQty) = Format(Val(RsItem("lot_qty") & ""), gs_formatQty)
        
        If IsNull(RsItem("unit_cls")) Then
          .TextMatrix(i, bteColUnitCls) = " "
          .TextMatrix(i, 4) = " "
        Else
          .TextMatrix(i, bteColUnitCls) = Trim(RsItem("Unit_cls"))
          .TextMatrix(i, bteColUnit) = uf_GetUnitDescription(Trim(RsItem("Unit_Cls")))
        End If
                
        If adaim = True Then
            .TextMatrix(i, bteColCurrCode) = ""
            .TextMatrix(i, bteColCurr) = ""
            .TextMatrix(i, bteColPrice) = Format(0, gs_formatPrice)
        Else
            If IsNull(RsItem("currency_code")) Then
               .TextMatrix(i, bteColCurrCode) = ""
               .TextMatrix(i, bteColCurr) = ""
            Else
              .TextMatrix(i, bteColCurrCode) = Trim(RsItem("currency_code"))
              .TextMatrix(i, bteColCurr) = uf_GetCurrencyDescription(Trim(RsItem("Currency_code")))
            End If
            .TextMatrix(i, bteColPrice) = Format(Trim(RsItem("price")), gs_formatPrice)
        End If
        
        .TextMatrix(i, bteColReq) = Format(IIf(IsNull(RsItem("requirement")), 0, RsItem("requirement")), gs_formatQty)
        req = IIf(IsNull(RsItem("requirement")), 0, RsItem("requirement"))
        
'        If Trim(RsItem("Item_Code")) = "708843" Then
'            Stop
'        End If
        
        .TextMatrix(i, bteColFixOrder) = Format(IIf(IsNull(RsItem("fixorder")), 0, RsItem("fixorder")), gs_formatQty)
        
        .TextMatrix(i, bteColStock) = Format(IIf(IsNull(RsItem("curstock")), 0, RsItem("curstock")), gs_formatQty)
        lastMth = IIf(IsNull(RsItem("curstock")), 0, RsItem("curstock"))
        
        .TextMatrix(i, btecolReqNext) = Format(IIf(IsNull(RsItem("requirementnextmonth")), 0, RsItem("requirementnextmonth")), gs_formatQty)
        reqN = IIf(IsNull(RsItem("requirementnextmonth")), 0, RsItem("requirementnextmonth"))
        
        .TextMatrix(i, bteColSafe) = Format(IIf(IsNull(RsItem("safety_stock")), 0, RsItem("safety_stock")), gs_formatQty)
        safe = IIf(IsNull(RsItem("safety_stock")), 0, RsItem("safety_stock"))
        
        .TextMatrix(i, bteColSafePercen) = Format(IIf(IsNull(RsItem("safety_stock_percentage")), 0, RsItem("safety_stock_percentage")), gs_formatQty)
        safePer = IIf(IsNull(RsItem("safety_stock_percentage")), 0, RsItem("safety_stock_percentage"))
        
        .TextMatrix(i, bteColSuggestion) = Format(suggestionOrder(CDbl(lastMth), CDbl(req), CDbl(reqN), CDbl(safe), CDbl(safePer), CDbl(moq), CDbl(spq)), gs_formatQty)
        .TextMatrix(i, bteColOrder) = Format(0, gs_formatQty)
        '.TextMatrix(i, bteColFixOrder) = Format(.TextMatrix(i, bteColOrder), gs_formatQty)
        '.TextMatrix(i, bteColEndQty) = Format(CDbl(lastMth) + CDbl(.TextMatrix(i, bteColOrder)) - CDbl(req), gs_formatQty)
        .TextMatrix(i, bteColEndQty) = Format(CDbl(lastMth) + CDbl(.TextMatrix(i, bteColFixOrder)) - CDbl(req), gs_formatQty)
        .TextMatrix(i, bteColAmount) = Format(CDbl(.TextMatrix(i, bteColPrice)) * CDbl(.TextMatrix(i, bteColOrder)), gs_formatAmount)

        .Cell(flexcpBackColor, i, bteColSelect) = &HFFFFFF
        .Cell(flexcpChecked, i, bteColSelect) = flexUnchecked
        .Cell(flexcpBackColor, i, bteColOrder) = &HFFFFFF
'        .Cell(flexcpBackColor, i, bteColCurr) = &HFFFFFF
'        .Cell(flexcpBackColor, i, bteColPrice) = &HFFFFFF
        RsItem.MoveNext
        i = i + 1
    Loop
    End With
    End If

End Function

Function suggestionOrder(lastM As Double, req As Double, reqNext As Double, safety As Double, safetyPercen As Double, _
moq As Double, spq As Double)
    Dim safetyStock As Double
    Dim rndSuggestion As Double

    If safetyPercen > 0 Then
        safetyStock = ((safetyPercen / 100) * reqNext) + req
    Else
        safetyStock = safety + req
    End If


    If safetyStock <= lastM Then
        suggestionOrder = 0
    Else
        If (safetyStock - lastM) < moq Then
            'suggestionOrder = moq
            'Exit Function
            rndSuggestion = moq
        Else
            rndSuggestion = spq
        End If
        
        If rndSuggestion = 0 Then
            suggestionOrder = 0
        Else
            suggestionOrder = (RoundUp((safetyStock - lastM) / rndSuggestion)) * rndSuggestion
            
        End If
    End If
        
End Function

Sub Kosong()
    cboWHTo.Text = ""
    CboMat = strAll
    cboDeliverTo.Text = ""
    txtsupplier.Text = ""
    txtAddress.Text = ""
    cboSupplier.Text = ""
    dtpPeriod.Value = Format(Now, "MMM yyyy")
    temptgl = dtpPeriod.Month
    txtPoNo.Text = ""
    txtPONo2.Text = ""
    dtpPODate.Value = Format(Now, "dd MMM yyyy")
    isipodate = Format(dtpPODate, "yyyy-mm-dd")
    dtpDeliveryDate.Value = Format(Now + 1, "dd MMM yyyy")
    Call ppn(dtpDeliveryDate.Value)
    grid.FocusRect = flexFocusNone
    cboAlarm.ListIndex = 1
    txtSearch = ""
    
    ubah = False
    ada = False
    LblErrMsg = ""
    statusfix = 0
    kodeitem = ""

    kunci (False)
    GetDefaultValue
    kosongBwh
    Header
End Sub

Sub kosongBwh()
    txtremarks.Text = ""
    txtamount.Text = Format(0, gs_formatAmount)
    txtPPN.Text = Format(0, gs_formatAmount)
    txtGrandTotal.Text = Format(0, gs_formatAmount)
    TxtPOLOT.Text = ""
    txtSearch = ""
End Sub


Private Sub AddtoCboMat()

    Dim adoRs As New ADODB.Recordset
    
    sql = "Select * From Material_Cls"
    adoRs.Open sql, Db, adOpenKeyset, adLockOptimistic
    
    CboMat.clear
    CboMat.columnCount = 2
    CboMat.TextColumn = 1
    
    CboMat.AddItem ""
    CboMat.List(CboMat.ListCount - 1, 0) = strAll
    CboMat.List(CboMat.ListCount - 1, 1) = strAll
    
    While Not adoRs.EOF
        CboMat.AddItem ""
        CboMat.List(CboMat.ListCount - 1, 0) = Trim(adoRs!Material_Cls)
        CboMat.List(CboMat.ListCount - 1, 1) = Trim$(adoRs!Description)
        adoRs.MoveNext
    Wend
    CboMat.ColumnWidths = "40 pt; 160 pt"
    CboMat.ListWidth = 200
    CboMat.ListRows = 4

End Sub


Sub AddToComboSupplier()
    
    Dim sqlcust As String
    Dim RsCust As New Recordset

    sqlcust = "select trade_code, trade_name, address1, country_cls, po_cls, Epte_Cls " & _
        "from trade_master where trade_cls='2' or Trade_Cls='3'"
        
    Set RsCust = Db.Execute(sqlcust)
    With cboSupplier
        .clear
        .columnCount = 6
        .ColumnWidths = "50pt;300pt;0pt;50pt;0pt;0pt"
        .ListWidth = 350
        .ListRows = 15
        i = 0
        Do While Not RsCust.EOF
            .AddItem
            .List(i, 0) = Trim(RsCust("Trade_code"))
            .List(i, 1) = IIf(IsNull(RsCust("trade_name")), " ", Trim(RsCust("Trade_Name")))
            .List(i, 2) = IIf(IsNull(RsCust("address1")), " ", Trim(RsCust("Address1")))
            .List(i, 3) = IIf(IsNull(RsCust("country_cls")), 0, Trim(RsCust("country_cls")))
            .List(i, 4) = IIf(IsNull(RsCust("po_cls")), 0, Trim(RsCust("po_cls")))
            .List(i, 5) = IIf(IsNull(RsCust("Epte_Cls")), 0, Trim(RsCust("Epte_Cls")))
            RsCust.MoveNext
            i = i + 1
        Loop
        RsCust.Close
    End With
    
End Sub

Sub AddToDeliveryPlace()
    
    Dim sqlcust As String
    Dim adoRs As New Recordset

    sqlcust = "Select Location_Code, Location_Name From Delivery_Place Where Trade_Code = '999'"
        
    Set adoRs = Db.Execute(sqlcust)
    With cboDeliverTo
        .clear
        .columnCount = 2
        .ColumnWidths = "50pt;200pt"
        .ListWidth = 250
        .ListRows = 5
        i = 0
        Do While Not adoRs.EOF
            .AddItem
            .List(i, 0) = Trim(adoRs("Location_Code"))
            .List(i, 1) = IIf(IsNull(adoRs("Location_Name")), " ", Trim(adoRs("Location_Name")))
            adoRs.MoveNext
            i = i + 1
        Loop
        adoRs.Close
    End With
    
End Sub

Sub AddToComboPONo(p As String)
Dim sqlno As String
Dim rsno As New Recordset
    
    sqlno = "select po_no from purchaseorder_master where sheetcoil_cls=0 " & _
            "and whto in (" & _
            "       select code from (select trade_code code from trade_master union select  wh_code code from warehouse_Master) a " & _
            ") " & _
            "and year(po_date)='" & Year(dtpPODate) & "' " & _
            "and month(po_date)='" & Month(dtpPODate) & "' " & p
    Set rsno = Db.Execute(sqlno)

    With CboPOnO
        .clear
        .ColumnWidths = "130pt"
        .ListWidth = 130
        .ListRows = 15

        i = 0
        Do While Not rsno.EOF
            .AddItem Trim(rsno("PO_No"))
            rsno.MoveNext
            i = i + 1
        Loop
    End With

End Sub

Private Sub GetDefaultValue()
    
    Dim RsCust As New Recordset

    sql = "select Price_Condition, POPayment_Terms, Insurance_Cls, Transportation_Cls, " & _
        "POMarking1, POMarking2, POMarking3, POMarking4, POMarking5, POMarking6 " & _
        "from trade_master where trade_code = '" & cboSupplier.Text & "'"
        
    RsCust.Open sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not RsCust.EOF Then
        
        CboPacking.Text = ""
        cboPriceCondition.Text = Trim(RsCust.Fields("Price_Condition") & "")
        cboPaymentTerm.Text = Trim(RsCust.Fields("POPayment_Terms") & "")
        cboInsuranceCls.Text = Trim(RsCust.Fields("Insurance_Cls") & "")
        cboTransport.Text = Trim(RsCust.Fields("Transportation_Cls") & "")
        txtMarking(0).Text = Trim(RsCust.Fields("POMarking1") & "")
        txtMarking(1).Text = Trim(RsCust.Fields("POMarking2") & "")
        txtMarking(2).Text = Trim(RsCust.Fields("POMarking3") & "")
        txtMarking(3).Text = Trim(RsCust.Fields("POMarking4") & "")
        txtMarking(4).Text = Trim(RsCust.Fields("POMarking5") & "")
        txtMarking(5).Text = Trim(RsCust.Fields("POMarking6") & "")
    
    Else
    
        CboPacking.Text = ""
        cboPriceCondition.Text = ""
        cboPaymentTerm.Text = ""
        cboInsuranceCls.Text = ""
        cboTransport.Text = ""
        txtMarking(0).Text = ""
        txtMarking(1).Text = ""
        txtMarking(2).Text = ""
        txtMarking(3).Text = ""
        txtMarking(4).Text = ""
        txtMarking(5).Text = ""
        
    End If
    RsCust.Close
    
    txtSearch = ""
    
End Sub

Sub PONO(ByVal Tgl As String, ByVal supp As String)
    
'    Dim rsno As New Recordset, LO As Integer, LM As String
'
'    If cboSupplier.Column(3) = "1" Then
'        LO = 2
'    Else
'        LO = 1
'    End If
'
'    If CboMat = strAll Then
'        LM = "3"
'    Else
'        LM = Right(CboMat, 1)
'    End If
'
'    Sql = "select isnull(max(Right(rtrim(po_no), 3)), 0) + 1 new_po from purchaseorder_master Where Year(PO_Date) = " & dtpPODate.year & " And month(PO_Date) = " & dtpPODate.month
'    Set rsno = Db.Execute(Sql)
'
'    If Not (rsno.BOF And rsno.EOF) Then
'        txtpono.Text = "KI3-" & Format(dtpPODate.Value, "YYMM") & "/" & Trim(UserInitPO) & LO & LM & "/" '& Format(rsno.Fields("new_po"), "000")
'        txtpono.SetFocus
'        SendKeys "{end}"
'    Else
'        txtpono.Text = "KI3-" & Format(dtpPODate.Value, "YYMM") & "/" & Trim(UserInitPO) & LO & LM & "/001"
'    End If
'
'    txtpono.locked = True

End Sub

Sub kunci(l As Boolean)
    dtpPODate.Enabled = Not l
    dtpDeliveryDate.Enabled = Not l
    dtpPeriod.Enabled = Not l
    grid.Editable = Not l
    Command1(0).Enabled = Not l
    lblFix.Visible = l
    statuskunci = l
End Sub

Sub ppn(ByVal d As Date)
Dim sqlppn As String
Dim rs3 As New ADODB.Recordset
    
    sqlppn = "select rate from tax_cls where tax_code='PPN' and start_date<='" & _
             Format(d, "yyyymmdd") & "' and end_date>='" & Format(d, "yyyymmdd") & "' "
    Set rs3 = Db.Execute(sqlppn)
    
    If Not (rs3.BOF And rs3.EOF) Then
        isippn = IIf(IsNull(rs3(0)), 0, CDbl(rs3(0)))
    Else
        isippn = 0
    End If
End Sub

Sub cekprice(ByVal Baris As Integer)
Dim sqlcp As String
Dim rsCP As New Recordset
statusprice = False

'    sqlcp = "select price from price_master where " & _
'           "item_code='" & grid.TextMatrix(Baris, bteColProdCode) & "' and price_cls='01' and (trade_code='" & cboSupplier.Text & _
'           "' or trade_code='000000') and start_date<='" & Format(dtpDeliveryDate.Value, "yyyymmdd") & "' and end_date>='" & _
'           Format(dtpDeliveryDate.Value, "yyyymmdd") & "' "
           
    sqlcp = "select price from price_master where " & _
           "item_code='" & grid.TextMatrix(Baris, bteColProdCode) & "' and price_cls='01' and (trade_code='" & cboSupplier.Text & _
           "' or trade_code='000000') and month(start_date)='" & Month(dtpPeriod) & "' and year(Start_date)='" & _
           Year(dtpPeriod) & "' "
           
    Set rsCP = Db.Execute(sqlcp)

    If Not (rsCP.BOF And rsCP.EOF) Then
        Do While Not rsCP.EOF
            If rsCP(0) = 0 Then statusprice = True: Exit Sub
            rsCP.MoveNext
        Loop
    End If

End Sub

Sub formatprice()
Dim p1 As Byte, p2 As String, p0 As String
Dim jmldigit As Byte, jmldigit0 As Byte, j As Integer

jmldigit = 0
    With grid
        For i = 1 To .Rows - 1
            If InStr(1, .TextMatrix(i, bteColPrice), ".") > 0 Then _
                jmldigit0 = Len(Trim(.TextMatrix(i, bteColPrice))) - InStr(1, Trim(.TextMatrix(i, bteColPrice)), ".")
            If jmldigit0 > jmldigit Then jmldigit = jmldigit0
        Next i

        For i = 1 To .Rows - 1
            p0 = Trim(.TextMatrix(i, bteColPrice))
            If InStr(1, p0, ".") > 0 Then
                p1 = Len(p0) - InStr(1, p0, ".")
                For j = 1 To jmldigit - p1
                    p2 = p0 & " "
                    p0 = p2
                Next j
            End If
            .TextMatrix(i, bteColPrice) = p0
        Next i
    
    End With
End Sub

Function Stock(ByVal Item As String, ByVal p As Date, cur As Double) As Double
Dim F, R, e, X As Double
Dim kodeno As String

     sqlreq = "select childrequirement_qty, childrequirementresult_qty from requirement_master where childitem_code='" & _
              Trim(Item) & "' and childrequirement_month='" & Month(p) & "' and childrequirement_year='" & _
              Year(p) & "' "
     Set rsreq = Db.Execute(sqlreq)
     If Not (rsreq.BOF And rsreq.EOF) Then
         R = IIf(IsNull(rsreq(0)), 0, rsreq(0)) - IIf(IsNull(rsreq(1)), 0, rsreq(1))
     Else
         R = 0
     End If
     If R < 0 Then R = 0
     

    If (Format(tempperiod2, "MMM yyyy") <> Format(dtpPeriod.Value, "MMM yyyy")) Or (Format(tempdeldate, "01 MMM yyyy") <> Format(dtpDeliveryDate.Value, "01 MMM yyyy")) Then
        sqlpo = "select po_no from purchaseorder_master where year(delivery_date)=" & _
                Year(p) & " and month(delivery_date)=" & Month(p) & " and po_no<>'" & Trim(txtPoNo.Text) & "' "
    Else
        sqlpo = "select po_no from purchaseorder_master where year(delivery_date)=" & _
                Year(p) & " and month(delivery_date)=" & Month(p)   '& " and po_no<>'" & Trim(txtpono.Text) & "' "
    End If
    Set rsPO = Db.Execute(sqlpo)
    X = 0
    kodeno = ""
    If Not (rsPO.BOF And rsPO.EOF) Then
        Do While Not rsPO.EOF
            If X = 0 Then
                kodeno = Trim(rsPO(0))
            Else
                kodeno = kodeno & "','" & Trim(rsPO(0))
            End If
            rsPO.MoveNext
            X = X + 1
        Loop
    End If
    
     sqlfixord = "select sum(qty) as qty from purchaseorder_detail " & _
                 "where po_no in ('" & kodeno & "') and item_code = '" & Trim(Item) & "' " & _
                 "and (complete_cls<>1 or complete_cls is null)"
     Set rsfixord = Db.Execute(sqlfixord)
     If Not (rsfixord.BOF And rsfixord.EOF) Then

        sqlrec = "select sum(case when part_receipt.receipt_cls='r1' then -part_receipt.qty " & _
                 "else part_receipt.qty end) as qty from part_receipt, purchaseorder_detail " & _
                 "where part_receipt.po_no=purchaseorder_detail.po_no and part_receipt.item_code=purchaseorder_detail.item_code and " & _
                 "part_receipt.po_no in ('" & kodeno & "') and part_receipt.item_code = '" & Trim(Item) & "' and " & _
                 "(purchaseorder_detail.complete_cls<>1 or purchaseorder_detail.complete_cls is null)"
         Set rsrec = Db.Execute(sqlrec)

         If Not (rsrec.BOF And rsrec.EOF) Then
             F = CDbl(IIf(IsNull(rsfixord(0)), 0, rsfixord(0))) - CDbl(IIf(IsNull(rsrec(0)), 0, rsrec(0)))
         Else
             F = IIf(IsNull(rsfixord(0)), 0, rsfixord(0))
         End If
     Else
         F = 0
     End If
     If F < 0 Then F = 0
    
     Stock = CDbl(cur) + CDbl(F) - CDbl(R)
     cur = Format(Stock, gs_formatQty)
End Function

Private Sub cboDeliverTo_Change()
    
    If cboDeliverTo.MatchFound Then txtDeliverTo.Text = cboDeliverTo.Column(1) Else txtDeliverTo.Text = ""
    
End Sub

Private Sub cboInsuranceCls_Change()

    If cboInsuranceCls.MatchFound Then txtInsurance.Text = cboInsuranceCls.Column(1) Else txtInsurance.Text = ""

End Sub

Private Sub CboMat_Change()
Dim t As String
If CboMat.ListIndex >= 0 Then LblMat = CboMat.Column(1) Else LblMat = strAll
    If cboStatus.Text = "Create" Then
        t = Format(Month(dtpPODate), "0#") & "/" & Year(dtpPODate)
        Call PONO(t, cboSupplier.Text)
    End If

End Sub

Private Sub cboPacking_Change()
    
    If CboPacking.MatchFound Then TxtPacking.Text = CboPacking.Column(1) Else TxtPacking.Text = ""
    
End Sub

Private Sub cboPaymentTerm_Change()
    
    If cboPaymentTerm.MatchFound Then txtPaymentTerm.Text = cboPaymentTerm.Column(1) Else txtPaymentTerm.Text = ""
    
End Sub

Private Sub CboPOnO_Change()

    cboWHTo.locked = CboPOnO.MatchFound

End Sub

Private Sub cbopricecondition_Change()
    
    If cboPriceCondition.MatchFound Then txtPriceCondition.Text = cboPriceCondition.Column(1) Else txtPriceCondition.Text = ""
    
End Sub

Private Sub cboTransport_Change()
    
    If cboTransport.MatchFound Then TxtTransport.Text = cboTransport.Column(1) Else TxtTransport.Text = ""
    
End Sub

Private Sub cboWHTo_Change()
    Dim t As String
    If cboWHTo.MatchFound Then txtWHTo.Text = cboWHTo.Column(1) Else txtWHTo.Text = ""
    If cboStatus.Text = "Create" Then
        t = Format(Month(dtpPODate), "0#") & "/" & Year(dtpPODate)
        'Call PoNo(t, cboSupplier.Text)
    End If
End Sub


Private Sub ChkCreate_Click()

If ChkCreate.Value = 1 Then
    Call ChkCreateAktif
Else
    ChkUpdate.Value = 1
    Call ChkUPdateAktif
End If

End Sub
Private Sub ChkUPdateAktif()
    cboStatus.ListIndex = 1
    cboStatus.Enabled = False
    Command1(2).Caption = "&Update"
    txtPoNo.locked = True
    SeqNoSet = 0
    Kosong
    AddCboSetCode
    Call CekPOSet
    
End Sub
Private Sub ChkCreateAktif()
    ChkUpdate.Value = 0
    cboStatus.ListIndex = 0
    cboStatus.Enabled = True
    Command1(2).Caption = "&Create"
    txtPoNo.locked = False
    Kosong
    AddCboSetCode
    Call CekPOSet
End Sub


Private Sub ChkUpdate_Click()
If ChkUpdate.Value = 1 Then
    ChkCreate.Value = 0
    Call ChkUPdateAktif
Else
    ChkCreate.Value = 1
    Call ChkCreateAktif
End If
End Sub

Private Sub CmdCari_Click(Index As Integer)
Call CekPOSet
'Call browseitem
End Sub

Private Sub cmdReport_Click()

    Me.MousePointer = vbHourglass
'    If cboSupplier.Column(3) = "0" Then
'        POLocal txtPONo.Text, bteHakPrice
'    Else
        POImport txtPoNo.Text, bteHakPrice, 1, "", dtpDeliveryDate.Value
'    End If
    Me.MousePointer = vbDefault

End Sub

Private Sub cmdSearch_Click()
    
    Dim i As Double
    
    LblErrMsg = ""
    
    If txtSearch = "" Or grid.Rows = 2 Then txtSearch.SetFocus: Exit Sub
    If grid.Row = grid.Rows - 1 Then i = 2 Else i = grid.Row + 1
    
    Do
        Select Case cboSearch.ListIndex
        Case 0
            grid.Col = bteColProdCode
            If UCase(Mid(grid.TextMatrix(i, bteColProdCode), 1, Len(txtSearch))) = UCase(txtSearch) Then
                Exit Do
            End If
        Case 1
            grid.Col = bteColDesc
            If InStr(UCase(grid.TextMatrix(i, bteColDesc)), UCase(txtSearch)) <> 0 Then
                Exit Do
            End If
        End Select
        i = i + 1
        If i = grid.Rows - 1 Then
            txtSearch = ""
            i = 2
            LblErrMsg = DisplayMsg(8012)
            Exit Do
        End If
    Loop
    
    grid.Row = i
    grid.TopRow = i
    grid.SetFocus
    
End Sub

Private Sub ComboBox1_Change()
If CboMat.ListIndex >= 0 Then LblMat = CboMat.Column(1) Else LblMat = strAll
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        cmdSearch_Click
    End If
End Sub

Private Sub Form_Load()
    If gb_Simulation = True Then Call up_InitSimulation(Me)
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
    
    sql = "select * from purchaseorder_master"
    If RS.State <> adStateClosed Then RS.Close
    RS.Open sql, Db, adOpenKeyset, adLockOptimistic

    sqlGrid = "select * from purchaseorder_detail order by item_code"
    If rsGrid.State <> adStateClosed Then rsGrid.Close
    rsGrid.Open sqlGrid, Db, adOpenKeyset, adLockOptimistic
    
    bteHakPrice = (hakPrice(Me.Name))
    
    lblCaption(10).Visible = (bteHakPrice = 1)
    lblCaption(11).Visible = (bteHakPrice = 1)
    lblCaption(12).Visible = (bteHakPrice = 1)
    
    txtamount.Visible = (bteHakPrice = 1)
    txtPPN.Visible = (bteHakPrice = 1)
    txtGrandTotal.Visible = (bteHakPrice = 1)
    
    'AddToComboSupplier
    AddToDeliveryPlace
    AddtoCboMat
    
    cboStatus.AddItem "Create"
    cboStatus.AddItem "Update"
    
    up_FillCombo cboPriceCondition, "PriceCondition_Cls"
    cboPriceCondition.ListWidth = 225
    cboPriceCondition.ColumnWidths = "50pt;175"
    
    up_FillCombo cboPaymentTerm, "PaymentTerm_Cls"
    cboPaymentTerm.ListWidth = 225
    cboPaymentTerm.ColumnWidths = "50pt;175"
    
    up_FillCombo CboPacking, "POPacking_Cls"
    CboPacking.ListWidth = 225
    CboPacking.ColumnWidths = "50pt;175"
    
    up_FillCombo cboInsuranceCls, "Insurance_Cls"
    cboInsuranceCls.ListWidth = 225
    cboInsuranceCls.ColumnWidths = "50pt;175"
    
    up_FillCombo cboTransport, "Transportation_Cls"
    cboTransport.ListWidth = 225
    cboTransport.ColumnWidths = "50pt;175"
    
    cboAlarm.AddItem "Yes"
    cboAlarm.AddItem "No"
    
    Call up_FillCombo(cbocurr, "curr_cls")
    cbocurr.TextColumn = 2
    
    SetComboWHTo
    Kosong
    cboStatus.ListIndex = 1
    
    With cboSearch
        .AddItem "Item Code"
        .AddItem "Description"
        .ListIndex = 0
    End With
    
    ChkCreate.Value = 1
    
    Call AddCboSetCode

End Sub

Private Sub cboStatus_Click()
Dim ketemu As Boolean
Dim t As String


    ketemu = False
    LblErrMsg = ""

    kunci (False)
    GetDefaultValue
    kosongBwh
    Header

    If cboStatus.ListIndex = 0 Then
        Command1(2).Caption = "Create"
        ClearPO
        ubah = False
        CboPOnO.locked = True
        txtPoNo.Text = "KI3-"
        dtpDeliveryDate.Value = Format(Now + 1, "dd MMM yyyy")
        If cboSupplier.Text <> "" Then
            t = Format(Month(dtpPODate), "0#") & "/" & Year(dtpPODate)
            Call PONO(t, cboSupplier.Text)
        End If
        cboWHTo.locked = False
    Else
        If cboSupplier.Text = "" Then
            CboPOnO.clear
            txtPoNo.Text = ""
        Else
            sql = " and supplier_Code='" & cboSupplier.Text & "' "
            AddToComboPONo (sql)
        End If

        ubah = True
        Command1(2).Caption = "Update"
        CboPOnO.locked = False
        'txtPONo.locked = False

        For i = 0 To CboPOnO.ListCount - 1
            If txtPoNo.Text = CboPOnO.List(i) Then
                ketemu = True
                CboPOnO.ListIndex = i
                Exit For
            End If
        Next
        If ketemu = False Then
            txtPoNo.Text = ""
            dtpDeliveryDate.Value = Format(Now + 1, "dd MMM yyyy")
        End If
        cboWHTo.locked = CboPOnO.MatchFound
    End If

End Sub

Private Sub cboStatus_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
  If KeyCode = 13 Then cboStatus_Click
End Sub

Private Sub cbopono_Click()
    LblErrMsg = ""
    txtPoNo.Text = CboPOnO.Text
    Header
    GetDefaultValue
    kosongBwh
    
    Dim p As String
    
    sql = "select * from purchaseorder_master where po_no='" & txtPoNo.Text & "' and sheetcoil_cls=0"
    If RS.State <> adStateClosed Then RS.Close
    RS.Open sql, Db, adOpenKeyset, adLockOptimistic

    If Not (RS.BOF And RS.EOF) Then
        isipodate = Format(dtpPODate, "yyyy-mm-dd")
        dtpPODate.Value = IIf(IsNull(RS("po_date")), " ", Format(Trim(RS("po_date")), "dd MMM yyyy"))
        Call dtpPODate_Change
        p = IIf(IsNull(RS("period")), "", Left(Trim(RS("period")), 4) & "-" & Right(Trim(RS("period")), 2) & "-01")
        dtpPeriod.Value = Format(p, "MMM yyyy")
        temptgl = dtpPeriod.Month
        cboSupplier.Text = Trim(RS("Supplier_code"))
        dtpDeliveryDate.Value = IIf(IsNull(RS("delivery_date")), " ", Format(Trim(RS("delivery_date")), "dd MMM yyyy"))
        cboWHTo.Text = Trim(RS("WHTo") & "")
        cboDeliverTo.Text = Trim(RS("Deliver_To") & "")
        txtRevisi.Text = Trim(RS("Revise_No") & "")
    End If
    
End Sub

Private Sub cbopono_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
  If KeyCode = 13 Then cbopono_Click
End Sub

Private Sub lblCaption_DblClick(Index As Integer)
    If txtPoNo.locked Then txtPoNo.locked = False Else txtPoNo.locked = True
End Sub


Private Sub TxtPOLOT_KeyPress(KeyAscii As Integer)
'If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtpono_Change()
    
    Dim ketemu As Boolean
    
    txtPONo2.Text = txtPoNo.Text
    If cboStatus.ListIndex = 1 Then
        For i = 0 To CboPOnO.ListCount - 1
            If txtPoNo.Text = CboPOnO.List(i) Then
                ketemu = True
                CboPOnO.ListIndex = i
                Exit For
            End If
        Next
        If ketemu = False Then CboPOnO.ListIndex = -1
    End If

End Sub

Private Sub txtPONo_GotFocus()
SendKeys "{End}"
End Sub

Private Sub txtpono_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
      If cboStatus.ListIndex = 0 Then
         SendKeys vbTab
      Else
        Header
        GetDefaultValue
        kosongBwh
        Dim p As String
        LblErrMsg = ""
        sql = "select * from purchaseorder_master where po_no='" & txtPoNo.Text & "' and sheetcoil_cls=0"
        If RS.State <> adStateClosed Then RS.Close
        RS.Open sql, Db, adOpenKeyset, adLockOptimistic
    
        If Not (RS.BOF And RS.EOF) Then
            isipodate = Format(dtpPODate, "yyyy-mm-dd")
            dtpPODate.Value = IIf(IsNull(RS("po_date")), " ", Format(Trim(RS("po_date")), "dd MMM yyyy"))
            Call dtpPODate_Change
            p = IIf(IsNull(RS("period")), " ", Left(Trim(RS("period")), 4) & "-" & Right(Trim(RS("period")), 2) & "-01")
            dtpPeriod.Value = Format(p, "MMM yyyy")
            temptgl = dtpPeriod.Month
            cboSupplier.Text = Trim(RS("Supplier_code"))
            dtpDeliveryDate.Value = IIf(IsNull(RS("delivery_date")), " ", Format(Trim(RS("delivery_date")), "dd MMM yyyy"))
        End If
      End If
    End If
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = Asc("'") Then KeyAscii = 0
    
End Sub

Private Sub cbosupplier_Click()
Dim ketemu As Boolean
Dim t As String

Call cboStatus_Click

ketemu = False
LblErrMsg = ""
kunci (False)
ClearPO
    If cboSupplier.ListIndex <> -1 Then
        cboWHTo.Text = ""
        cboDeliverTo.Text = ""
        txtsupplier.Text = cboSupplier.Column(1)
        txtAddress.Text = cboSupplier.Column(2)
        countrycls = cboSupplier.Column(3)
        GetDefaultValue
        If cboStatus.ListIndex = 1 Then
            sql = " and supplier_Code='" & cboSupplier.Text & "' "
            AddToComboPONo (sql)

            For i = 0 To CboPOnO.ListCount - 1
                If txtPoNo.Text = CboPOnO.List(i) Then
                    ketemu = True
                    CboPOnO.ListIndex = i
'                    Browse
                    Exit For
                End If
            Next
            If ketemu = False Then
                txtPoNo.Text = ""
                dtpDeliveryDate.Value = Format(Now + 1, "dd MMM yyyy")
            End If
            GetDefaultValue
            kosongBwh
            Header
        Else
            t = Format(Month(dtpPODate), "0#") & "/" & Year(dtpPODate)
            Call PONO(t, cboSupplier.Text)
        End If

    Else
        cboWHTo.Text = ""
        cboDeliverTo.Text = ""
        txtsupplier.Text = ""
        txtAddress.Text = ""
        countrycls = 0
        CboPOnO.clear
        If cboStatus.ListIndex = 1 Then
            txtPoNo.Text = ""
            dtpDeliveryDate.Value = Format(Now + 1, "dd MMM yyyy")
            GetDefaultValue
            kosongBwh
            Header
        Else
            txtPoNo.Text = ""
        End If
        LblErrMsg.Caption = DisplayMsg(4050) '"Record with this Supplier Code not Exist"
        cboSupplier.SetFocus
        Exit Sub

    End If
    
    If countrycls = 1 Then
        isippn = 0
    Else
        Call ppn(dtpDeliveryDate.Value)
    End If

End Sub

Private Sub cbosupplier_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
  If KeyCode = 13 Then
    For i = 0 To cboSupplier.ListCount - 1
        If cboSupplier.Text = cboSupplier.List(i) Then
            cboSupplier.ListIndex = i
            Exit For
        End If
    Next
    cbosupplier_Click
  End If
End Sub

Private Sub cboSupplier_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub dtpPODate_Change()
Dim ketemu As Boolean
Dim isidtpPO2 As Date
Dim t As String

isidtpPO2 = Format(dtpPODate, "yyyy-mm-dd")
If DateDiff("m", isipodate, isidtpPO2) <> 0 Then
    If cboStatus.ListIndex = 1 Then
        If cboSupplier.Text = "" Then
            CboPOnO.clear
            txtPoNo.Text = ""
        Else
            sql = " and supplier_Code='" & cboSupplier.Text & "' "
            AddToComboPONo (sql)
        End If

        For i = 0 To CboPOnO.ListCount - 1
            If txtPoNo.Text = CboPOnO.List(i) Then
                ketemu = True
                CboPOnO.ListIndex = i
                Exit For
            End If
        Next
        If ketemu = False Then
            txtPoNo.Text = ""
            dtpDeliveryDate.Value = Format(Now + 1, "dd MMM yyyy")
            GetDefaultValue
            kosongBwh
            Header
        End If
    Else
        If cboSupplier.Text <> "" Then
            t = Format(Month(dtpPODate), "0#") & "/" & Year(dtpPODate)
            Call PONO(t, cboSupplier.Text)
        End If
    End If
End If
isipodate = Format(dtpPODate, "yyyy-mm-dd")
End Sub

Private Sub dtpPeriod_Change()
Call dtpPeriod_Click
temptgl = dtpPeriod.Month

    If cboStatus.ListIndex = 1 Then
        Header
        GetDefaultValue
        kosongBwh
    End If

End Sub

Private Sub dtpPeriod_Click()
    If dtpPeriod.Month = 1 And Val(temptgl) = 12 Then dtpPeriod.Year = dtpPeriod.Year + 1
    If dtpPeriod.Month = 12 And Val(temptgl) = 1 Then dtpPeriod.Year = dtpPeriod.Year - 1
End Sub

Private Sub dtpDeliveryDate_Change()
    If countrycls = 1 Then
        isippn = 0
    Else
        Call ppn(dtpDeliveryDate.Value)
    End If
    
    If cboStatus.ListIndex = 1 Then
        Header
        GetDefaultValue
        kosongBwh
    End If
End Sub

Private Sub cbocurr_Click()
    If cbocurr.ListIndex <> -1 Then
        grid.TextMatrix(actrow, bteColCurrCode) = cbocurr.Column(0)
        grid.TextMatrix(actrow, bteColCurr) = cbocurr.Column(1)
    End If
End Sub

Private Sub cbocurr_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then cbocurr_Click
End Sub

Private Sub cbocurr_LostFocus()
    cbocurr.Visible = False
End Sub

Private Sub cboprice_Change()
If InStr(1, cboprice.Text, ",") = 1 Then cboprice.Text = Right(cboprice, Len(cboprice) - 1)
End Sub

Private Sub cboprice_Click()
    If cboprice.ListIndex <> -1 Then
        grid.TextMatrix(actrow, bteColCurrCode) = cboprice.Column(2)
        For i = 0 To cbocurr.ListCount - 1
            If Trim(grid.TextMatrix(actrow, bteColCurrCode)) = Trim(cbocurr.List(i)) Then
                cbocurr.ListIndex = i
                Exit For
            End If
        Next i
        grid.TextMatrix(actrow, bteColCurr) = uf_GetCurrencyDescription(Trim(cboprice.Column(2)))
        grid.TextMatrix(actrow, bteColUnitCls) = cboprice.Column(3)
        grid.TextMatrix(actrow, bteColUnit) = uf_GetUnitDescription(Trim(cboprice.Column(3)))
    End If
End Sub

Private Sub cboprice_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
  If KeyCode = 13 Then cboprice_Click
End Sub

Private Sub CboPrice_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then
          KeyAscii = 0
    End If
    If InStr(1, cboprice.Text, ".") > 1 Then If KeyAscii = Asc(".") Then KeyAscii = 0
    If Trim(cboprice.Text) = "" Then cboprice.Text = Format(0, gs_formatPrice)
    If CDbl(cboprice.Text & Chr(KeyAscii)) > gd_MaxPrice Then KeyAscii = 0
End Sub

Private Sub cboPrice_LostFocus()
Dim sql3 As String
Dim rs3 As New Recordset
Dim ketemu As Boolean
    
    If cboprice.Text = "" Then cboprice.Text = Format(0, gs_formatPrice)
    
    Dim z As Double
    z = CDbl(cboprice.Text)
    If z > gd_MaxPrice Then
        cboprice.Text = Left(z, 10)
    End If
        
    grid.TextMatrix(actrow, bteColPrice) = Format(cboprice.Text, gs_formatPrice)
    Call Grid_AfterEdit(actrow, bteColPrice)
    
    cboprice.Text = Format(cboprice.Text, gd_MaxPrice)    'If cboprice.Text <> 0 Then
    
    For i = 0 To cboprice.ListCount - 1
        If Trim(cboprice.Text) = Trim(cboprice.List(i)) Then
            ketemu = True
            cboprice.ListIndex = i
            Exit For
        End If
    Next i
    
    If ketemu = False Then
        sql3 = "select unit_cls from item_master where item_code='" & grid.TextMatrix(actrow, bteColProdCode) & "' "
        Set rs3 = Db.Execute(sql3)
        
        If Not (rs3.BOF And rs3.EOF) Then
            grid.TextMatrix(actrow, bteColUnitCls) = rs3(0)
            grid.TextMatrix(actrow, bteColUnit) = uf_GetUnitDescription(Trim(rs3(0)))
        End If
    End If
    cboprice.Visible = False
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim a As Double

a = 0
With grid
    If .Col = bteColOrder Then
        txtSearch = ""
        If .TextMatrix(Row, bteColOrder) = "" Then .TextMatrix(Row, bteColOrder) = Format(0, gs_formatQty)
        If IsNumeric(.TextMatrix(Row, bteColOrder)) = False Then .TextMatrix(Row, bteColOrder) = Format(0, gs_formatQty)
        If CDbl(.TextMatrix(Row, bteColOrder)) > gd_MaxQty Then LblErrMsg = DisplayMsg(4045) & " " & gd_MaxQty: .TextMatrix(Row, bteColOrder) = Format(orderawal, gs_formatQty): .SetFocus: Exit Sub  '"Quantity must be lower or equal than 9,999,999.99"
        LblErrMsg = ""
        If CDbl(.TextMatrix(Row, bteColOrder)) < CDbl(.TextMatrix(Row, bteColMinOrder)) Then
            
            LblErrMsg = "[9000]-Quantity must be greater or equal to Minimum Order"
            .TextMatrix(Row, bteColOrder) = Format(orderawal, gs_formatQty)
            .SetFocus
            Exit Sub
        End If
            
        .TextMatrix(Row, bteColOrder) = Format(.TextMatrix(Row, bteColOrder), gs_formatQty)
        
        If Year(dtpPeriod) = Year(dtpDeliveryDate) And Month(dtpPeriod) = Month(dtpDeliveryDate) Then
            .TextMatrix(Row, bteColFixOrder) = Format((CDbl(.TextMatrix(Row, bteColFixOrder)) + CDbl(.TextMatrix(Row, bteColOrder)) - orderawal), gs_formatQty)
        ElseIf Format(dtpPeriod, "yyyy-mm-01") > Format(dtpDeliveryDate, "yyyy-mm-01") Then
            .TextMatrix(Row, bteColStock) = Format((CDbl(.TextMatrix(Row, bteColStock)) + CDbl(.TextMatrix(Row, bteColOrder)) - orderawal), gs_formatQty)
        ElseIf Format(dtpPeriod, "yyyy-mm-01") < Format(dtpDeliveryDate, "yyyy-mm-01") Then

        End If
        .TextMatrix(Row, bteColEndQty) = Format((CDbl(.TextMatrix(Row, bteColStock)) + CDbl(.TextMatrix(Row, bteColFixOrder)) - CDbl(.TextMatrix(Row, bteColReq))), gs_formatQty)
    End If
    If .Col = bteColSelect Or .Col = bteColOrder Then 'Or .Col = bteColPrice Then
            txtSearch = ""
            formatprice
            .TextMatrix(Row, bteColAmount) = Format(CDbl(.TextMatrix(Row, bteColOrder)) * CDbl(.TextMatrix(Row, bteColPrice)), gs_formatAmount)
            For i = 2 To .Rows - 1
                If .Cell(flexcpChecked, i, bteColSelect) = flexChecked Then
                    a = a + .TextMatrix(i, bteColAmount)
                End If
            Next i
            txtamount.Text = Format(a, gs_formatAmount)
            If isippn = 0 Then
                txtPPN.Text = Format(0, gs_formatAmount)
            Else
                txtPPN.Text = Format(CDbl(isippn / 100) * CDbl(txtamount.Text), gs_formatAmount)
            End If
            txtPPN.Text = Format(txtPPN.Text, gs_formatAmount)
            txtGrandTotal = Format(CDbl(txtPPN.Text) + CDbl(txtamount.Text), gs_formatAmount)
    End If
    
End With
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
  actrow = Row
  If grid.Cell(flexcpChecked, Row, bteColSelect) <> flexChecked Then
    If grid.Col <> bteColSelect Then
        Cancel = True
    End If
  Else
    If grid.Col <> bteColSelect And grid.Col <> bteColOrder Then 'And Grid.Col <> bteColCurr And Grid.Col <> bteColPrice Then
      Cancel = True
    End If
    If grid.Col = bteColOrder Then orderawal = CDbl(grid.TextMatrix(Row, bteColOrder))
  End If
End Sub

Private Sub Grid_CellChanged(ByVal Row As Long, ByVal Col As Long)
If grid.Col = bteColOrder Then _
If InStr(1, grid.TextMatrix(Row, Col), ",") = 1 Then grid.TextMatrix(Row, Col) = Right(grid.TextMatrix(Row, Col), Len(grid.TextMatrix(Row, Col)) - 1)
End Sub

Private Sub grid_Click()
  If statuskunci = False Then
  If grid.Row > 1 Then
  If grid.Cell(flexcpChecked, grid.Row, bteColSelect) = flexChecked Then
    With grid
        If .Col = bteColCurr Then
'            cboCurr.top = .Cell(flexcpTop, .Row, bteColCurr)
'            cboCurr.Left = .Cell(flexcpLeft, .Row, bteColCurr)
'            cboCurr.Width = .CellWidth + 30
'            Call up_FillCombo(cboCurr, "curr_cls")
'            cboCurr.TextColumn = 2
'            If Grid.TextMatrix(.Row, bteColCurr) <> "" Then
'                cboCurr.Text = Grid.TextMatrix(.Row, bteColCurr)
'                For i = 0 To cboCurr.ListCount - 1
'                    If Trim(Grid.TextMatrix(.Row, bteColCurrCode)) = Trim(cboCurr.List(i)) Then
'                        cboCurr.ListIndex = i
'                        Exit For
'                    End If
'                Next i
'            End If
'            cboCurr.Visible = True
'            cboCurr.SetFocus
'            cboPrice.Visible = False
        ElseIf .Col = bteColPrice Then
'            cboPrice.top = .Cell(flexcpTop, .Row, bteColPrice)
'            cboPrice.Left = .Cell(flexcpLeft, .Row, bteColPrice)
'            cboPrice.Width = .CellWidth + 30
'            cboPrice.Text = ""
'            BrowsePrice
'            If Grid.TextMatrix(.Row, bteColPrice) <> 0 Then
'                cboPrice.Text = Grid.TextMatrix(.Row, bteColPrice)
'                For i = 0 To cboPrice.ListCount - 1
'                    If Trim(Grid.TextMatrix(.Row, bteColPrice)) = Trim(cboPrice.List(i)) Then
'                        cboPrice.ListIndex = i
'                        Exit For
'                    End If
'                Next i
'            End If
'            cboPrice.Visible = True
'            cboPrice.SetFocus
'            cboCurr.Visible = False
'        Else
'            cboCurr.Visible = False
'            cboPrice.Visible = False
        End If
        
        If .Col = bteColOrder Then
            .FocusRect = flexFocusInset
        Else
            .FocusRect = flexFocusNone
        End If
    End With
  End If
  End If
  End If
End Sub

Private Sub Grid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
LblErrMsg = ""
  If grid.Col = bteColOrder Then
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then
      KeyAscii = 0
    End If
  End If
End Sub

Private Sub Grid_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    cbocurr.Visible = False
    cboprice.Visible = False
End Sub

Private Sub Grid_DblClick()
Dim date1 As Date
Dim diff As Integer
Dim startm As Date
Dim endm As Date

If grid.Rows > 2 Then
    If hakAkses("frm_ReceiptSupplyScheculeInquiry") = 0 Then LblErrMsg = DisplayMsg(3007):   Exit Sub
    date1 = DateAdd("m", 1, dtpPeriod.Value)
    diff = DateDiff("d", Format(dtpPeriod, "yyyy-mm-01"), Format(date1, "yyyy-mm-01"))
    startm = CDate(Year(dtpPeriod) & "-" & Month(dtpPeriod) & "-01")
    endm = CDate(Year(dtpPeriod) & "-" & Month(dtpPeriod) & "-" & diff)
    With grid
        popanggil = "poparts"
        frm_ReceiptSupplyScheculeInquiry.CboItemCD = .TextMatrix(.Row, bteColProdCode)
        frm_ReceiptSupplyScheculeInquiry.DMonth(0) = Format(startm, "dd MMM yyyy")
        frm_ReceiptSupplyScheculeInquiry.DMonth(1) = Format(endm, "dd MMM yyyy")
        frm_ReceiptSupplyScheculeInquiry.ClickSearch
        frm_ReceiptSupplyScheculeInquiry.grid.LeftCol = 12
        frm_ReceiptSupplyScheculeInquiry.Show
        frm_ReceiptSupplyScheculeInquiry.Cmd_save(8).Caption = "&Back"
    End With
End If
End Sub

Private Sub Command1_Click(Index As Integer)

Dim sql3 As String, sql4 As String
Dim a As Double, R As Double
Dim sqldetil As String, t As String
Dim rsdetil As New Recordset, rs4 As New Recordset
Dim strTempCurr As String
Dim VSeq As Long

Me.MousePointer = vbHourglass
LblErrMsg = ""
a = 0

Select Case Index
Case 0

    If hakUpdate(Me.Name) = 0 Then LblErrMsg = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
    
    If txtPoNo.Text = "" Then
        txtPoNo.SetFocus
        LblErrMsg = DisplayMsg(1048) '"Please Select PO No"
        Me.MousePointer = vbDefault
        Exit Sub
    ElseIf cboSupplier.Text = "" Then
        cboSupplier.SetFocus
        LblErrMsg = DisplayMsg(1054) '"Please Select Supplier Code"
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    
    'Sql = "select * from purchaseorder_master where left(po_no, 16) ='" & Mid(txtPONo.Text, 1, 16) & "' and sheetcoil_cls=0"
    
    sql = "select * from purchaseorder_master where po_no ='" & Trim(txtPoNo.Text) & "'"
    
    If RS.State <> adStateClosed Then RS.Close
    RS.Open sql, Db, adOpenKeyset, adLockOptimistic

    If RS.BOF And RS.EOF Then
      LblErrMsg.Caption = DisplayMsg(4015)
      txtPoNo.SetFocus
      Me.MousePointer = vbDefault
      Exit Sub
    End If

    If ubah = True Then
        
        RS("period") = Year(dtpPeriod.Value) & Format(Month(dtpPeriod.Value), "0#")
        RS("po_date") = Format(dtpPODate.Value, "YYYY-MM-DD")
        RS("delivery_date") = Format(dtpDeliveryDate.Value, "YYYY-MM-DD")
        RS("WHTo") = cboWHTo.Text
        RS("Deliver_To") = cboDeliverTo.Text
        RS("Revise_No") = txtRevisi.Text
        RS("PriceCondition_Cls") = cboPriceCondition.Text
        RS("PaymentTerm_Cls") = cboPaymentTerm.Text
        RS("POPacking_Cls") = CboPacking.Text
        RS("Insurance_Cls") = cboInsuranceCls.Text
        RS("Transportation_Cls") = cboTransport.Text
        
        RS("PO_LOT") = TxtPOLOT.Text
        
        RS("POMarking1") = txtMarking(0).Text
        RS("POMarking2") = txtMarking(1).Text
        RS("POMarking3") = txtMarking(2).Text
        RS("POMarking4") = txtMarking(3).Text
        RS("POMarking5") = txtMarking(4).Text
        RS("POMarking6") = txtMarking(5).Text
        RS("remarks") = txtremarks.Text
        RS("amount") = txtamount.Text
        RS("ppn") = txtPPN.Text
        RS("total_amount") = txtGrandTotal.Text
        RS("remarks") = txtremarks.Text
        RS("Last_Update") = Now
        RS("Last_User") = userLogin
        RS.update
                
        strTempCurr = ""
        With grid
            For i = 2 To .Rows - 1
                If .Cell(flexcpChecked, i, bteColSelect) = flexChecked Then
                    If strTempCurr = "" Then strTempCurr = .TextMatrix(i, bteColCurrCode)
                    If CDbl(.TextMatrix(i, bteColOrder)) = 0 Then
                        .Col = bteColOrder
                        .Row = i
                        .SetFocus
                        actrow = i
                        LblErrMsg = DisplayMsg(1012) '"Please Input Order Quantity"
                        Me.MousePointer = vbDefault
                        Exit Sub
                    ElseIf CDbl(.TextMatrix(i, bteColOrder)) > gd_MaxQty Then
                        LblErrMsg = DisplayMsg(4045) & " " & gd_MaxQty '"Quantity must be lower or equal than 9,999,999.99"
                        .Col = bteColOrder
                        .Row = i
                        .SetFocus
                        actrow = i
                        Me.MousePointer = vbDefault
                        Exit Sub
                    ElseIf .TextMatrix(i, bteColCurr) = "" Then
                        If bteHakPrice = 0 Then
                            .TextMatrix(i, bteColCurrCode) = gs_DefaultCurrencyCode
                            .TextMatrix(i, bteColPrice) = 0
                        Else
                            .Col = bteColCurr
                            .Row = i
                            actrow = i
                            .SetFocus
                            grid_Click
                            LblErrMsg = DisplayMsg(1028)
                            Me.MousePointer = vbDefault
                            Exit Sub
                        End If
                    ElseIf CDbl(.TextMatrix(i, bteColPrice)) = 0 Then
                        If bteHakPrice = 0 Then
                            .TextMatrix(i, bteColPrice) = Format(0, gs_formatPrice)
                        Else
                            Call cekprice(i)
                            'Info dari Pak Teguh harga 0 tidak divalidasi 20 Juli 2010
'                            If Not statusprice Then
'                                .Col = bteColPrice
'                                .Row = i
'                                .SetFocus
'                                actrow = i
'                                grid_Click
'                                LblErrMsg = DisplayMsg(1029) '"Please Input Price"
'                                Me.MousePointer = vbDefault
'                                Exit Sub
'                            End If
                        End If
                    ElseIf .TextMatrix(i, bteColCurrCode) <> strTempCurr Then
                        .Col = bteColCurr
                        .Row = i
                        actrow = i
                        .SetFocus
                        grid_Click
                        LblErrMsg = DisplayMsg(4084)
                        Me.MousePointer = vbDefault
                        Exit Sub
                    End If
                
                Else
                    
                    sql4 = "select * from part_receipt where po_no='" & txtPoNo.Text & "' and item_code='" & .TextMatrix(i, bteColProdCode) & "' "
                    Set rs4 = Db.Execute(sql4)
                    If Not (rs4.BOF And rs4.EOF) Then
                        .Row = i
                        .SetFocus
                        actrow = i
                        LblErrMsg = DisplayMsg(1204)
                        Me.MousePointer = vbDefault
                        Exit Sub
                    End If
                        
                End If
                
            Next i
            
            If .Rows > 2 Then
            
                sql3 = "delete from purchaseorder_detail where po_no='" & txtPoNo.Text & "' "
                Db.Execute sql3
                R = 1
                        
                For i = 2 To .Rows - 1
                    If .Cell(flexcpChecked, i, bteColSelect) = flexChecked Then
                        rsGrid.AddNew
                        VSeq = SeqN
                        rsGrid("Seq_no") = VSeq
                        rsGrid("po_no") = txtPoNo.Text
                        rsGrid("item_Code") = .TextMatrix(i, bteColProdCode)
                        rsGrid("price") = .TextMatrix(i, bteColPrice)
                        rsGrid("currency_code") = .TextMatrix(i, bteColCurrCode)
                        rsGrid("unit_cls") = .TextMatrix(i, bteColUnitCls)
                        rsGrid("qty") = .TextMatrix(i, bteColOrder)
                        rsGrid("delivery_date") = dtpDeliveryDate
                        rsGrid("amount") = .TextMatrix(i, bteColAmount)
                        rsGrid("Last_Update") = Now
                        rsGrid("Last_User") = userLogin
                        rsGrid.update
                        R = R + 1
                    End If
                Next i
                
            End If
            
            LblErrMsg = DisplayMsg(1101)
            
        End With
        
    End If

Case 1
    Kosong
    cboStatus.ListIndex = 1
    Call cboStatus_Click
    cboSupplier.SetFocus
Case 2

    If cboStatus.ListIndex = 0 Then
        
        If hakUpdate(Me.Name) = 0 Then LblErrMsg = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
        If cboSupplier.Text = "" Then
            cboSupplier.SetFocus
            LblErrMsg = DisplayMsg(1054) '"Please Select Supplier Code"
            Me.MousePointer = vbDefault
            Exit Sub
        ElseIf cboWHTo.MatchFound = False Then
            cboWHTo.SetFocus
            LblErrMsg = DisplayMsg("0031") '"Please Select Warehouse"
            Me.MousePointer = vbDefault
            Exit Sub
        Else
            If cboSupplier.Text <> "" Then
                cboSupplier.MatchEntry = 1
                cboSupplier.Text = cboSupplier.Text
                If cboSupplier.MatchFound = False Then
                    LblErrMsg = DisplayMsg(4050)
                    cboSupplier.SetFocus
                    cboSupplier.MatchEntry = 2
                    Me.MousePointer = vbDefault
                    Exit Sub
                End If
                cboSupplier.MatchEntry = 2
            End If

            If txtPoNo.Text = "" Then
                txtPoNo.SetFocus
                LblErrMsg = DisplayMsg(1046) '"Please Input PO No"
                Me.MousePointer = vbDefault
                Exit Sub
            End If

            On Error Resume Next
            If ubah = False Then

                'Sql = "select * from purchaseorder_master where left(po_no, 16) ='" & Mid(txtPONo.Text, 1, 16) & "' and sheetcoil_cls=0"

                sql = "select * from purchaseorder_master where po_no ='" & Trim(txtPoNo.Text) & "'"
                

                If RS.State <> adStateClosed Then RS.Close
                RS.Open sql, Db, adOpenKeyset, adLockOptimistic

                If Not (RS.BOF And RS.EOF) Then
                    LblErrMsg.Caption = DisplayMsg(1023)
                    txtPoNo.SetFocus
                    Me.MousePointer = vbDefault
                    Exit Sub
                Else
                    If ChkCreate.Value = 1 And SeqNoSet = 0 Then
                        SeqNoSet = POSetSeqNo
                    End If
                    
                    RS.AddNew
                    RS("po_no") = txtPoNo.Text
                    RS("supplier_code") = cboSupplier.Text
                End If

            End If

            RS("period") = Year(dtpPeriod.Value) & Format(Month(dtpPeriod.Value), "0#")
            RS("po_date") = Format(dtpPODate.Value, "YYYY-MM-DD")
            RS("WHTo") = cboWHTo.Text
            RS("Deliver_To") = cboDeliverTo.Text
            RS("Revise_No") = txtRevisi.Text
            RS("delivery_date") = Format(dtpDeliveryDate.Value, "YYYY-MM-DD")
            RS("amount") = txtamount.Text
            RS("ppn") = txtPPN.Text
            RS("total_amount") = txtGrandTotal.Text
            RS("remarks") = txtremarks.Text
            RS("sheetcoil_cls") = 0
            RS("Last_Update") = Now
            RS("Last_User") = userLogin
            RS("POSet_Code") = CboSetCode.Text
            RS("POSet_SeqNo") = SeqNoSet
            RS.update

            If InStr(1, err.Description, "Violation of PRIMARY KEY constraint") > 0 Then
                t = Format(Month(dtpPODate), "0#") & "/" & Year(dtpPODate)
                Call PONO(t, cboSupplier.Text)
                txtPONo2.Text = txtPoNo.Text
                RS("po_No") = txtPoNo.Text
                RS("Last_Update") = Now
                RS("Last_User") = userLogin
                RS.update
            End If

            'cboStatus.Text = "Update"
            If cboSupplier.Text <> "" Then browseitem: formatprice
            LblErrMsg.Caption = DisplayMsg(1000)
            ubah = True
        End If

    Else

        If txtPoNo.Text = "" Then
            txtPoNo.SetFocus
            LblErrMsg = DisplayMsg(1048) '"Please Select PO No"
            Me.MousePointer = vbDefault
            Exit Sub
        Else
            Browse
            If ada = False Then
                dtpDeliveryDate.Value = Format(Now + 1, "dd MMM yyyy")
                txtamount.Text = Format(0, gs_formatAmount)
                txtPPN.Text = Format(0, gs_formatAmount)
                txtGrandTotal.Text = Format(0, gs_formatAmount)
                txtremarks.Text = ""

                LblErrMsg.Caption = DisplayMsg(4015)
                txtPoNo.SetFocus
                Me.MousePointer = vbDefault
                Exit Sub
            End If
        End If

    End If
    
Case 3

    If txtPoNo.Text <> "" And cboSupplier.Text <> "" Then
        
        Dim p As String
        
        sql = "select * from purchaseorder_master where po_no='" & txtPoNo.Text & "' and sheetcoil_cls=0"
        If RS.State <> adStateClosed Then RS.Close
        RS.Open sql, Db, adOpenKeyset, adLockOptimistic
            
        If Not (RS.BOF And RS.EOF) Then
            isipodate = Format(dtpPODate, "yyyy-mm-dd")
            dtpPODate.Value = IIf(IsNull(RS("po_date")), " ", Format(Trim(RS("po_date")), "dd MMM yyyy"))
            p = IIf(IsNull(RS("period")), " ", Left(Trim(RS("period")), 4) & "-" & Right(Trim(RS("period")), 2) & "-01")
            dtpPeriod.Value = Format(p, "MMM yyyy")
            temptgl = dtpPeriod.Month
            cboSupplier.Text = Trim(RS("Supplier_code"))
            dtpDeliveryDate.Value = IIf(IsNull(RS("delivery_date")), " ", Format(Trim(RS("delivery_date")), "dd MMM yyyy"))
        End If
        Browse
    
    End If
            
End Select
Me.MousePointer = vbDefault

End Sub


Private Sub command2_Click(Index As Integer)
Dim Atas As String
LblErrMsg.Caption = ""

Select Case Index
    Case 1:
            If intpage = 1 Then
               LblErrMsg.Caption = DisplayMsg(4020) '"This is the first page !"
            ElseIf jmlpage > 1 Then
               intpage = 1
               LblErrMsg.Caption = ""
            End If

            On Error Resume Next
            grid.TopRow = 1

    Case 2:
            If intpage = 1 Then
               LblErrMsg = DisplayMsg(4020) '"This is the first page !"
            Else
               intpage = intpage - 1
               If intpage < 0 Then intpage = 0
               LblErrMsg = ""
            End If
            On Error Resume Next
            Atas = grid.TopRow

            grid.TopRow = grid.TopRow - 16
            If Atas = grid.TopRow Then grid.TopRow = 1

    Case 3:
            If intpage < jmlpage Then
              intpage = intpage + 1
              If intpage > jmlpage Then intpage = jmlpage
              LblErrMsg.Caption = ""
            Else
              LblErrMsg.Caption = DisplayMsg(4021) '"This is the last page !"
            End If

            On Error Resume Next
            grid.TopRow = grid.TopRow + 16

    Case 4:
            If intpage = jmlpage Then
              LblErrMsg.Caption = DisplayMsg(4021) '"This is the last page !"
            ElseIf intpage < jmlpage Then
              intpage = jmlpage
              LblErrMsg.Caption = ""
            End If

            On Error Resume Next
            grid.TopRow = grid.Rows

End Select
End Sub

Private Sub CmdSubMenu_Click()
        
    If ChkCreate.Value = 1 Then
        If Cek1Set = False Then
            If MsgBox("PO with supplier " & ListSupplier & " not created yet" & vbCrLf & "Do You want to Create it? ", vbYesNo + vbExclamation, "Confirmation") = vbYes Then
                Exit Sub
            End If
        End If
    End If

    ClearPO
    Unload Me
    frmMainMenu.Show

End Sub
Function Cek1Set() As Boolean
Dim sql As String
Dim RS As New Recordset
Dim i As Long

Cek1Set = True
ListSupplier = ""

For i = 0 To cboSupplier.ListCount - 1
    sql = " SELECT * FROM PurchaseOrder_Master " & vbCrLf & _
          " Where Supplier_Code = '" & Trim(cboSupplier.List(i, 0)) & "' " & vbCrLf & _
          "       and POSet_Code = '" & Trim(CboSetCode.Text) & "' and POSet_SeqNo = '" & SeqNoSet & "' "
    
    If RS.State = 1 Then RS.Close
    RS.Open sql, Db, adOpenKeyset, adLockReadOnly
    
    If RS.RecordCount = 0 Then
        Cek1Set = False
        ListSupplier = ListSupplier & Trim(cboSupplier.List(i, 0)) & ","
    End If
Next i

If Cek1Set = False Then
 ListSupplier = Left(ListSupplier, Len(ListSupplier) - 1)
End If
 
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
    Set RsItem = Nothing
    Set rscurstock = Nothing
    Set rsreq = Nothing
    Set rsfixord = Nothing
    Set rsPO = Nothing
    Set rsrec = Nothing
    Set rscomp1 = Nothing
    Set rscomp2 = Nothing
End Sub

Private Sub SetComboWHTo()

    Dim adoRs As New ADODB.Recordset
    
    sql = "Select rtrim(wh_code) as WC,wh_name as WN from warehouse_master " & _
            " where wh_code in ( " & _
            " select code from (select  trade_code code from trade_Master union select wh_code code from warehouse_master) a " & _
            " ) " & _
            " order by wh_code"
    adoRs.Open sql, Db, adOpenKeyset, adLockOptimistic
    
    cboWHTo.clear
    cboWHTo.columnCount = 2
    cboWHTo.TextColumn = 1
    While Not adoRs.EOF
        cboWHTo.AddItem ""
        cboWHTo.List(cboWHTo.ListCount - 1, 0) = Trim(adoRs!wC)
        cboWHTo.List(cboWHTo.ListCount - 1, 1) = Trim$(adoRs!wn)
        adoRs.MoveNext
    Wend
    cboWHTo.ColumnWidths = "60 pt; 180 pt"
    cboWHTo.ListWidth = 240
    cboWHTo.ListRows = 15

End Sub

Private Sub txtRevisi_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
End Sub

Private Sub ClearPO()
    
    sql = "delete from purchaseorder_master where not exists(select po_no from purchaseorder_detail where po_no = purchaseorder_master.po_no) and po_no = '" & txtPoNo & "'"
    Db.Execute sql
    
End Sub

Function SeqN() As Double
Dim rsmax As New ADODB.Recordset
    
    sql = "Select ISNULL(Max(Seq_No),0) + 1  SeqNo " & _
        "From PurchaseOrder_Detail"
    Set rsmax = Db.Execute(sql)
    SeqN = rsmax!seqNo
    rsmax.Close
End Function

Sub AddCboSetCode()
    
    Dim sql As String
    Dim RS As New Recordset

    sql = "SELECT DISTINCT Parent_SetCode, Description " & _
              "from [PO_Set_Master] "
        
    Set RS = Db.Execute(sql)
    With CboSetCode
        .clear
        .columnCount = 2
        .ColumnWidths = "80pt;200"
        .ListWidth = 280
        .ListRows = 15
        i = 0
        Do While Not RS.EOF
            .AddItem
            .List(i, 0) = Trim(RS("Parent_SetCode"))
            .List(i, 1) = Trim(RS("Description"))
           
            RS.MoveNext
            i = i + 1
        Loop
        RS.Close
    End With
    
End Sub
        
Function POSetSeqNo()
    Dim sql As String
    Dim rsPOSetSeqNo As New Recordset
    
    sql = "SElECT isnull(Max(POSet_SeqNo),0) POSet_SeqNo FROM PurchaseOrder_Master WHERE POSet_Code = '" & Trim(CboSetCode.Text) & "' "
    If rsPOSetSeqNo.State = 1 Then rsPOSetSeqNo.Close
    rsPOSetSeqNo.Open sql, Db, adOpenKeyset, adLockReadOnly
    
    If rsPOSetSeqNo.Fields("POSet_SeqNo") > 0 Then
        POSetSeqNo = rsPOSetSeqNo.Fields("POSet_SeqNo") + 1
    Else
        POSetSeqNo = 1
    End If

End Function

Private Sub CekPOSet()
Dim sql As String
Dim rsJmlSupplier As New Recordset

'Get Count of Supplier
sql = _
" SELECT SUPPLIER.*, Isnull(TM.trade_name,'') Trade_Name, Isnull(TM.address1,'') Address1, Isnull(TM.country_cls,'') Country_Cls " & vbCrLf & _
"        , isnull(TM.po_cls,'') PO_Cls, isnull(TM.Epte_Cls,'') EPTE_Cls FROM " & vbCrLf & _
" (             " & vbCrLf & _
"   select COALESCE(PM1.Trade_Code, PM0.TRade_Code,IM.Supplier_Code) Supplier_Code " & vbCrLf & _
"   from dbo.PO_Set_Master POS  " & vbCrLf & _
"   INNER JOIN Item_Master IM   " & vbCrLf & _
"       ON POS.Item_Code = IM.item_code " & vbCrLf & _
"   LEFT JOIN Price_Master PM1  " & vbCrLf & _
"       ON POS.Item_Code = PM1.Item_Code AND convert(char,Getdate(),112) between PM1.Start_Date and PM1.End_Date AND PM1.Price_Cls = '01' AND PM1.Priority_Cls = 1  " & vbCrLf & _
"   LEFT JOIN Price_master PM0  " & vbCrLf & _
"       ON POS.Item_Code = PM0.Item_Code AND convert(char,Getdate(),112) between PM0.Start_Date and PM0.End_Date AND PM0.Price_Cls = '01' AND PM0.Priority_Cls = 0  " & vbCrLf & _
"   WHERE POS.Parent_SetCode = '" & Trim(CboSetCode.Text) & "' " & _
" )SUPPLIER     " & vbCrLf & _
" INNER JOIN Trade_Master TM " & vbCrLf & _
"   ON SUPPLIER.Supplier_Code = TM.Trade_Code " & vbCrLf & _
" GROUP BY SUPPLIER_CODE, Trade_Name,Address1, Country_Cls, PO_Cls, EPTE_Cls    "

If rsJmlSupplier.State = 1 Then rsJmlSupplier.Close
rsJmlSupplier.Open sql, Db, adOpenKeyset, adLockReadOnly

If rsJmlSupplier.RecordCount > 0 Then
    
    With cboSupplier
        .clear
        .columnCount = 6
        .ColumnWidths = "50pt;300pt;0pt;0pt;0pt;0pt"
        .ListWidth = 350
        .ListRows = 15


        Dim i As Long
        For i = 0 To rsJmlSupplier.RecordCount - 1
            .AddItem
            .List(i, 0) = Trim(rsJmlSupplier!Supplier_Code)
            .List(i, 1) = Trim(rsJmlSupplier!trade_name)
            .List(i, 2) = Trim(rsJmlSupplier!address1)
            .List(i, 3) = Trim(rsJmlSupplier!country_cls)
            .List(i, 4) = Trim(rsJmlSupplier!PO_Cls)
            .List(i, 5) = Trim(rsJmlSupplier!Epte_cls)
            rsJmlSupplier.MoveNext
        Next
        
    End With

End If

End Sub
