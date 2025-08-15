VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPOSubconScheduled 
   BackColor       =   &H00FDDFE3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase Order Scheduled (Subcon)"
   ClientHeight    =   10920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   Icon            =   "frmPOSubconScheduled.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10920
   ScaleWidth      =   15270
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
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
      Left            =   10290
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   88
      TabStop         =   0   'False
      Top             =   8820
      Width           =   2355
   End
   Begin VB.TextBox TxtDisc 
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
      Left            =   3120
      MaxLength       =   25
      TabIndex        =   85
      TabStop         =   0   'False
      Top             =   8820
      Width           =   2355
   End
   Begin VB.TextBox TxtSubAmount 
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
      Left            =   1740
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   84
      TabStop         =   0   'False
      Top             =   8820
      Visible         =   0   'False
      Width           =   2355
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
      Left            =   12720
      Locked          =   -1  'True
      MaxLength       =   35
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   8820
      Width           =   2385
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
      Left            =   90
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   8820
      Width           =   2490
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
      Left            =   5550
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   8820
      Width           =   2355
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
      Left            =   7920
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   8820
      Width           =   2355
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
      Left            =   9960
      MaxLength       =   25
      TabIndex        =   22
      Top             =   6750
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
      Index           =   0
      Left            =   7290
      MaxLength       =   25
      TabIndex        =   20
      Top             =   6750
      Width           =   2085
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
      Left            =   2925
      Locked          =   -1  'True
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   7035
      Width           =   3570
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
      Left            =   2925
      Locked          =   -1  'True
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   6705
      Width           =   3585
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
      Left            =   2925
      Locked          =   -1  'True
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   8025
      Width           =   3585
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
      Left            =   12870
      MaxLength       =   25
      TabIndex        =   24
      Top             =   6750
      Width           =   2085
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
      Left            =   7650
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   26
      Top             =   7590
      Width           =   7470
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
      Left            =   2925
      Locked          =   -1  'True
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   7710
      Width           =   3585
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
      Left            =   2925
      Locked          =   -1  'True
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   7380
      Width           =   3585
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
      Left            =   9960
      MaxLength       =   25
      TabIndex        =   23
      Top             =   7110
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
      Index           =   1
      Left            =   7290
      MaxLength       =   25
      TabIndex        =   21
      Top             =   7095
      Width           =   2085
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
      Left            =   2925
      MaxLength       =   25
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   6315
      Width           =   2430
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
      Left            =   5430
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   6285
      Width           =   1125
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
      Left            =   12870
      MaxLength       =   25
      TabIndex        =   25
      Top             =   7110
      Width           =   2085
   End
   Begin VB.TextBox txtRev 
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
      Left            =   5460
      MaxLength       =   20
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   2250
      Width           =   675
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   435
      Left            =   13343
      TabIndex        =   32
      Top             =   210
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   767
   End
   Begin VB.CommandButton cmdreport 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Preview"
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
      Left            =   10215
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   10050
      Width           =   1125
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   580
      Left            =   83
      TabIndex        =   46
      Top             =   870
      Width           =   15105
      Begin MSComCtl2.DTPicker requestdate1 
         Height          =   315
         Left            =   1920
         TabIndex        =   0
         Top             =   180
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
      Begin MSComCtl2.DTPicker requestdate2 
         Height          =   315
         Left            =   3840
         TabIndex        =   1
         Top             =   180
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
      Begin VB.Label Label9 
         Alignment       =   2  'Center
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
         Left            =   3360
         TabIndex        =   49
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Request Date From"
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
         Left            =   120
         TabIndex        =   48
         Top             =   240
         Width           =   1815
      End
      Begin MSForms.ComboBox cborequestno 
         Height          =   315
         Left            =   6720
         TabIndex        =   2
         Top             =   180
         Width           =   1815
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "3201;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Request No"
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
         Left            =   5640
         TabIndex        =   47
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton command1 
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
      Index           =   3
      Left            =   11475
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   10050
      Width           =   1125
   End
   Begin VB.TextBox txtpono 
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
      Left            =   2070
      MaxLength       =   25
      TabIndex        =   5
      Top             =   2250
      Width           =   2550
   End
   Begin VB.CommandButton command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Create"
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
      Left            =   12510
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2250
      Width           =   1125
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FDDFE3&
      Height          =   533
      Left            =   60
      TabIndex        =   36
      Top             =   9210
      Width           =   15045
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
         Height          =   255
         Left            =   90
         TabIndex        =   37
         Top             =   150
         Width           =   14865
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
      Left            =   13980
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   10050
      Width           =   1125
   End
   Begin VB.CommandButton cmdsubmenu 
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
      Left            =   210
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   10050
      Width           =   1125
   End
   Begin VB.CommandButton command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Clea&r"
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
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   10050
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   720
      Left            =   83
      TabIndex        =   38
      Top             =   1440
      Width           =   15105
      Begin VB.TextBox lblcust 
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
         Height          =   200
         Index           =   0
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   280
         Width           =   4995
      End
      Begin VB.TextBox lblcust 
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
         Index           =   1
         Left            =   9120
         Locked          =   -1  'True
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   280
         Width           =   5715
      End
      Begin VB.Line Line2 
         Index           =   1
         X1              =   9120
         X2              =   14880
         Y1              =   525
         Y2              =   525
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
         Height          =   255
         Index           =   4
         Left            =   8280
         TabIndex        =   42
         Top             =   270
         Width           =   840
      End
      Begin MSForms.ComboBox cbocust 
         Height          =   315
         Left            =   975
         TabIndex        =   3
         Top             =   240
         Width           =   1920
         VariousPropertyBits=   612386843
         MaxLength       =   10
         DisplayStyle    =   3
         Size            =   "3387;556"
         ColumnCount     =   12
         ListRows        =   15
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   3030
         X2              =   8070
         Y1              =   525
         Y2              =   525
      End
      Begin VB.Label LblCode 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier "
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
         TabIndex        =   39
         Top             =   270
         Width           =   1095
      End
   End
   Begin MSComCtl2.DTPicker podate 
      Height          =   315
      Left            =   7170
      TabIndex        =   7
      Top             =   2220
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
   Begin MSComCtl2.DTPicker DeliveryDate 
      Height          =   315
      Left            =   10110
      TabIndex        =   8
      Top             =   2220
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   3495
      Left            =   135
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2700
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   6165
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
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
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmPOSubconScheduled.frx":0E42
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "grid"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Additional"
      TabPicture(1)   =   "frmPOSubconScheduled.frx":0E5E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GridAdd"
      Tab(1).ControlCount=   1
      Begin VSFlex8Ctl.VSFlexGrid GridAdd 
         Height          =   2850
         Left            =   -74810
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   490
         Width           =   14730
         _cx             =   25982
         _cy             =   5027
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
         Rows            =   2
         Cols            =   1
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
      Begin VSFlex8Ctl.VSFlexGrid grid 
         Height          =   2850
         Left            =   200
         TabIndex        =   11
         Top             =   490
         Width           =   14730
         _cx             =   25982
         _cy             =   5027
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
         FormatString    =   $"frmPOSubconScheduled.frx":0E7A
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
         Begin MSComCtl2.DTPicker DelDate 
            Height          =   315
            Left            =   3720
            TabIndex        =   81
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
            Format          =   141230083
            CurrentDate     =   37798
         End
         Begin MSForms.ComboBox cboprice 
            Height          =   285
            Left            =   6840
            TabIndex        =   83
            TabStop         =   0   'False
            Top             =   480
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
         Begin MSForms.ComboBox cbocurr 
            Height          =   285
            Left            =   5640
            TabIndex        =   82
            TabStop         =   0   'False
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
      End
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PPh"
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
      Left            =   10350
      TabIndex        =   89
      Top             =   8520
      Width           =   315
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Discount"
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
      Left            =   3855
      TabIndex        =   87
      Top             =   8490
      Width           =   735
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Total"
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
      Left            =   2475
      TabIndex        =   86
      Top             =   8490
      Visible         =   0   'False
      Width           =   810
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
      Index           =   1
      Left            =   12330
      TabIndex        =   80
      Top             =   7110
      Width           =   450
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
      Index           =   0
      Left            =   12330
      TabIndex        =   79
      Top             =   6780
      Width           =   450
   End
   Begin VB.Label Label3 
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
      Height          =   255
      Index           =   3
      Left            =   8820
      TabIndex        =   78
      Top             =   2250
      Width           =   1365
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
      Left            =   7980
      TabIndex        =   77
      Top             =   8520
      Width           =   315
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
      TabIndex        =   76
      Top             =   8520
      Width           =   525
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
      Left            =   5940
      TabIndex        =   75
      Top             =   8520
      Width           =   1140
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
      Left            =   12750
      TabIndex        =   74
      Top             =   8520
      Width           =   1005
   End
   Begin MSForms.ComboBox cboPriceCondition 
      Height          =   315
      Left            =   1875
      TabIndex        =   16
      Top             =   6990
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
      Left            =   75
      TabIndex        =   73
      Top             =   7050
      Width           =   1290
   End
   Begin MSForms.ComboBox cboPaymentTerm 
      Height          =   315
      Left            =   1875
      TabIndex        =   15
      Top             =   6660
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
      Left            =   75
      TabIndex        =   72
      Top             =   6720
      Width           =   1260
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
      Left            =   75
      TabIndex        =   71
      Top             =   7380
      Width           =   660
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
      Left            =   9450
      TabIndex        =   70
      Top             =   6810
      Width           =   450
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
      Left            =   6750
      TabIndex        =   69
      Top             =   6780
      Width           =   450
   End
   Begin MSForms.ComboBox cboTransport 
      Height          =   315
      Left            =   1875
      TabIndex        =   19
      Top             =   7980
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
      Left            =   75
      TabIndex        =   68
      Top             =   8040
      Width           =   1245
   End
   Begin MSForms.ComboBox cboInsuranceCls 
      Height          =   315
      Left            =   1875
      TabIndex        =   18
      Top             =   7650
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
      Left            =   75
      TabIndex        =   67
      Top             =   7710
      Width           =   1650
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
      Left            =   12403
      TabIndex        =   66
      Top             =   4440
      Width           =   450
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
      Left            =   6735
      TabIndex        =   65
      Top             =   6360
      Width           =   975
   End
   Begin VB.Line Line4 
      X1              =   2940
      X2              =   6495
      Y1              =   7290
      Y2              =   7290
   End
   Begin VB.Line Line5 
      X1              =   2940
      X2              =   6495
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Line Line6 
      X1              =   2925
      X2              =   6495
      Y1              =   7620
      Y2              =   7620
   End
   Begin VB.Line Line7 
      X1              =   2940
      X2              =   6510
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comments"
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
      Left            =   6660
      TabIndex        =   64
      Top             =   7635
      Width           =   930
   End
   Begin VB.Line Line8 
      X1              =   2925
      X2              =   6495
      Y1              =   7950
      Y2              =   7950
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
      Left            =   9450
      TabIndex        =   63
      Top             =   7170
      Width           =   450
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
      Left            =   6750
      TabIndex        =   62
      Top             =   7155
      Width           =   510
   End
   Begin MSForms.ComboBox cboPacking 
      Height          =   315
      Left            =   1875
      TabIndex        =   17
      Top             =   7290
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
   Begin VB.Shape Shape5 
      BackColor       =   &H00A6D2FF&
      Height          =   915
      Left            =   6645
      Top             =   6615
      Width           =   8475
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
      Left            =   75
      TabIndex        =   61
      Top             =   6375
      Width           =   600
   End
   Begin MSForms.ComboBox cboSearch 
      Height          =   315
      Left            =   765
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   6315
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
      Left            =   12403
      TabIndex        =   60
      Top             =   4830
      Width           =   450
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Rev"
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
      Index           =   0
      Left            =   5010
      TabIndex        =   50
      Top             =   2280
      Width           =   525
   End
   Begin VB.Label lblfix 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Status Fix "
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
      Height          =   255
      Left            =   13950
      TabIndex        =   43
      Top             =   2325
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label Label3 
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
      Height          =   255
      Index           =   1
      Left            =   1470
      TabIndex        =   41
      Top             =   2310
      Width           =   720
   End
   Begin VB.Label Label3 
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
      Height          =   255
      Index           =   2
      Left            =   6300
      TabIndex        =   40
      Top             =   2280
      Width           =   825
   End
   Begin MSForms.ComboBox cbopono 
      Height          =   315
      Left            =   2070
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   2250
      Width           =   2835
      VariousPropertyBits=   612386843
      DisplayStyle    =   3
      Size            =   "5001;556"
      ListWidth       =   5291
      ListRows        =   15
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   2250
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
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Order Scheduled (Subcon)"
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
      Left            =   83
      TabIndex        =   35
      Top             =   240
      Width           =   15105
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00A6D2FF&
      BackStyle       =   1  'Opaque
      Height          =   300
      Left            =   60
      Top             =   8460
      Width           =   15045
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00A6D2FF&
      BackStyle       =   1  'Opaque
      Height          =   315
      Left            =   6645
      Top             =   6315
      Width           =   8475
   End
End
Attribute VB_Name = "frmPOSubconScheduled"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sql As String, sqlGrid As String
Dim RS As New ADODB.Recordset, rsGrid As New ADODB.Recordset
Dim i As Long, orderawal As Double, isippn As Long
Dim ubah As Boolean, ubahgrid As Boolean, ada As Boolean, sampun As Boolean
Dim statusfix As String
Dim actrow As Long, activecurrcd As String, activecurr As String
Dim countrycls As Byte
Const isiPOTerm = "after B/L Date,after Delivery Date,after Invoice Date,prior before Shipment,from Custom Clearance Date,after Receive Invoice,after Receive Goods"
Dim ColReqNo As Byte, ColCodeItem As Byte, ColDesc As Byte, ColUnit As Byte, ColQtyBox As Byte
Dim ColLot As Byte, ColReqQty As Byte, ColOrder As Byte, ColRemaining As Byte, ColDelDate As Byte, ColCurr As Byte, ColPrice As Byte
Dim ColAmount As Byte, ColReqSeqNo As Byte, colRemark As Byte, colTmpQty As Byte, colSeqNo As Byte

Dim bteColSelect  As Byte
Dim bteColAddCode  As Byte
Dim bteColAddName As Byte
Dim bteColAddQty As Byte
Dim bteColAddPrice  As Byte
Dim bteColAddAmount As Byte
Dim bteColProdCode As Byte
Dim bteColDesc As Byte
Dim bteColOrder As Byte
Dim bteColCurrCode As Byte

Private Sub ClearData()

    sql = "Delete from PurchaseOrder_Master Where PO_No = '" & Trim(txtPoNo.Text) & "' " & _
          "and PO_No not in (select PO_No from PurchaseOrder_Detail) and others_cls = '0' and period is null"
    Db.Execute sql

End Sub

Private Sub CekPONumber()
    
    Dim adoRs As New ADODB.Recordset
    
    sql = "Select * From PurchaseOrder_Master Where PO_No = '" & Trim(txtPoNo.Text) & "'"
    adoRs.Open sql, Db, adOpenKeyset, adLockOptimistic, adCmdText
    If Not adoRs.EOF Then
        Call PONO(Right(Year(PODate), 2), Format(Month(PODate), "0#"))
        adoRs.update "PO_No", Trim(txtPoNo.Text)
    End If
    adoRs.Close
    Set adoRs = Nothing
    
End Sub

Sub Kosong()
    LblErrMsg = ""
    requestdate1.Value = Format(Now, "yyyy-mm-01")
    requestdate2.Value = Format(Now, "dd MMM yyyy")
    Call adtocborequestno
    cborequestno.Text = ""
    
    cboCust.Text = ""
    lblcust(0).Text = "": lblcust(1).Text = ""
    txtPoNo.Text = "": txtPONo2.Text = ""
    PODate.Value = Format(Now, "dd MMM yyyy")
    PODate.Enabled = True
    Call ppn(PODate.Value)
    txtRev.Text = ""
    
    grid.FocusRect = flexFocusNone
    DelDate.Value = Format(Now + 1, "dd MMM yyyy")
    
    cboPriceCondition.Text = ""
    
    ubah = False: ada = False
    statusfix = 0
    
    Call kunci(False)
    Call kosongBwh: Call Header
End Sub

Sub kosongBwh()
    ' Add 20090112
    TxtSubAmount.Text = 0
    TxtDisc.Text = 0
    ' ---
    txtamount.Text = 0
    txtPPN.Text = 0
    txtGrandTotal.Text = 0
    cboPriceCondition.Text = ""
    cboPaymentTerm.Text = ""
    CboPacking.Text = ""
    cboInsuranceCls.Text = ""
    cboTransport.Text = ""
    txtMarking(0).Text = "": txtMarking(1).Text = "": txtMarking(2).Text = "": txtMarking(3).Text = ""
    txtMarking(4).Text = "": txtMarking(5).Text = "": txtremarks = ""
    
End Sub

Function adtocboCust(ByVal filter As Boolean)
Dim sqlcust As String
Dim RsCust As New Recordset

    If filter = True Then
        sqlcust = "select tm.trade_code, tm.trade_name, tm.address1, tm.country_cls, tm.po_cls, " & _
                  "tm.popayment_code, tm.popayment_day, tm.popayment_terms, tm.transportation_Cls, isnull(tm.Trade_Abbr,'') Trade_Abbr " & _
                  "from trade_master tm " & _
                  "Inner Join " & _
                  "(select distinct pm.trade_code as supplier_code From price_master pm " & _
                  "    inner join (Select item_code from porequest_detail where porequest_no = '" & Trim(cborequestno.Text) & "') prd on prd.item_code = pm.item_code " & _
                  "    where pm.price_cls = '01' " & _
                  "    Union " & _
                  "    select distinct im.supplier_code from item_master im " & _
                  "    inner join (Select item_code from porequest_detail where porequest_no = '" & Trim(cborequestno.Text) & "') prd on prd.item_code = im.item_code " & _
                  "    Union " & _
                  "    select distinct pom.supplier_code from PurchaseOrder_Master pom " & _
                  "    inner join (select PO_No from PurchaseOrder_Detail where porequest_no = '" & Trim(cborequestno.Text) & "') pod on pod.PO_No = pom.PO_No " & _
                  "    where isnull(pom.others_Cls,'0') = '0' and pom.period is null " & _
                  ") sc on sc.supplier_code = tm.trade_code " & _
                  "where (tm.trade_cls='3') " & _
                  "Order By tm.Trade_Abbr "
                  ' Untuk Supplier subcon
                  
    Else
        'sqlcust = "select trade_code, trade_name, address1, country_cls, po_cls, popayment_code, popayment_days, popayment_terms, transportation_cls, isnull(Trade_Abbr,'') Trade_Abbr " & _
                  "from trade_master where (trade_cls='2' or trade_cls='3') Order By Trade_Abbr"
          sqlcust = "select trade_code, trade_name, address1, country_cls, po_cls, isnull(Trade_Abbr,'') Trade_Abbr " & _
                  ", popayment_code, popayment_terms, transportation_Cls,Popayment_day " & _
                  "from trade_master where (trade_cls='3')"
    End If
    Set RsCust = Db.Execute(sqlcust)

    With cboCust
        .clear
        .columnCount = 8
        .ColumnWidths = "80pt;270pt;0pt;0pt;0pt;0pt;0pt;0pt"
        .ListWidth = 350
        .ListRows = 15

        i = 0
        Do While Not RsCust.EOF
            .AddItem
            
            .List(i, 0) = Trim(RsCust("Trade_code"))
            .List(i, 1) = IIf(IsNull(RsCust("trade_name")), "", Trim(RsCust("Trade_Name")))
            .List(i, 2) = IIf(IsNull(RsCust("address1")), "", Trim(RsCust("Address1")))
            .List(i, 3) = IIf(IsNull(RsCust("country_cls")), 3, Trim(RsCust("country_cls")))
            .List(i, 4) = IIf(IsNull(RsCust("po_cls")), 0, Trim(RsCust("po_cls")))
            .List(i, 5) = IIf(IsNull(RsCust("transportation_Cls")), -1, RsCust("transportation_Cls"))
            RsCust.MoveNext
            i = i + 1
        Loop
    End With
    Set RsCust = Nothing
End Function

Sub adtocbo(nmCombo, nmConst, start As Integer, col1 As Integer, Col2 As Integer, Kosong As Boolean)
Dim j As Integer, k As Integer

    With nmCombo
        .clear
        .columnCount = 2
        .TextColumn = 2
        j = 0: k = start
        
        If Kosong Then
            .AddItem
            .List(0, 0) = "Null"
            .List(0, 1) = ""
            j = 1
        End If
        
        For i = 0 To UBound(Split(nmConst, ","))
            .AddItem
            .List(j, 0) = k
            .List(j, 1) = Split(nmConst, ",")(i)
            k = k + 1: j = j + 1
        Next i
        .ListRows = 10
        .ListWidth = col1 + Col2
        .ColumnWidths = col1 & "pt;" & Col2 & "pt"
    End With
End Sub

Sub adtocombo()
Dim sql1 As String
Dim rs1 As New Recordset

    combo1.AddItem "Create"
    combo1.AddItem "Update"
    Call up_FillCombo(cbocurr, "Curr_Cls")
    'Call isiCboUnitCurr(cbocurr, isiCurr, 0, 4)
    cbocurr.TextColumn = 2
    
    Call up_FillCombo(cboPaymentTerm, "PaymentTerm_Cls")
    cboPaymentTerm.ColumnWidths = "50pt;175pt"
    cboPaymentTerm.ListWidth = 225
    
    Call up_FillCombo(cboInsuranceCls, "Insurance_Cls")
    cboInsuranceCls.ColumnWidths = "25pt;175pt"
    cboInsuranceCls.ListWidth = 200
    
    Call up_FillCombo(CboPacking, "PackingStyle_Cls")
    CboPacking.ColumnWidths = "25pt;175pt"
    CboPacking.ListWidth = 200
    
    
    Call up_FillCombo(cboTransport, "Transportation_Cls")
    cboTransport.ColumnWidths = "25pt;175pt"
    cboTransport.ListWidth = 200
    
 
    
    sql1 = "select * from PriceCondition_cls"
    If rs1.State <> adStateClosed Then rs1.Close
    rs1.Open sql1, Db, adOpenKeyset, adLockOptimistic
    If Not (rs1.BOF And rs1.EOF) Then
        i = 0
        With cboPriceCondition
            .clear
            .columnCount = 2
            .ColumnWidths = "25pt;175pt"
            .ListWidth = 200
            .ListRows = 15
            
            Do While Not rs1.EOF
                .AddItem ""
                .List(i, 0) = Trim(rs1!PriceCondition_Cls)
                .List(i, 1) = Trim(rs1!Description)
                i = i + 1
                rs1.MoveNext
            Loop
        End With
    End If
    Set rs1 = Nothing
    
     With cboSearch
        .AddItem "Item Code"
        .AddItem "Description"
        .ListIndex = 0
    End With

End Sub

Sub adtocborequestno()
Dim sqlno As String, tempcust As String
Dim rsno As New Recordset
    
    sqlno = "select porequest_no, porequest_period, isnull(fix_cls,'0') fix_cls " & _
            "from PORequest_Master where isnull(others_cls,'0')='0' and isnull(fix_cls,'0')='1' " & _
            "and porequest_date >= '" & Format(requestdate1.Value, "yyyy-mm-dd") & "' " & _
            "and porequest_date <= '" & Format(requestdate2.Value, "yyyy-mm-dd") & "' "

'    sqlno = "select porequest_no, porequest_period, isnull(fix_cls,'0') fix_cls " & _
'            "from PORequest_Master inner join "
'            "where isnull(others_cls,'0')='0' and isnull(fix_cls,'0')='1' " & _
'            "and porequest_date >= '" & Format(requestdate1.Value, "yyyy-mm-dd") & "' " & _
'            "and porequest_date <= '" & Format(requestdate2.Value, "yyyy-mm-dd") & "' "


    Set rsno = Db.Execute(sqlno)
    With cborequestno
        .clear
        .columnCount = 3
        
        If Trim(cboCust.Text) <> "" Then tempcust = Trim(cboCust.Text)
        Call adtocboCust(False): cboCust.Text = tempcust: Call cboCust_Click
        
        i = 0
        Do While Not rsno.EOF
            .AddItem
            .List(i, 0) = Trim(rsno("PORequest_No"))
            .List(i, 1) = Trim(rsno("PORequest_Period"))
            .List(i, 2) = Trim(rsno("fix_cls"))
            rsno.MoveNext
            i = i + 1
        Loop
        .ColumnWidths = "90pt;0pt;0pt"
        .ListWidth = 90
        .ListRows = 15
    End With
    Set rsno = Nothing
End Sub

Sub adtocbopono()
Dim sqlno As String
Dim rsno As New Recordset
    
'    sqlno = "select pom.PO_No from PurchaseOrder_Master pom " & _
'            "where pom.others_cls = '0' and pom.period is null " & _
'            "and pom.supplier_Code = '" & Trim(CboCust.Text) & "' " & _
'            "and pom.po_date >= '" & Format(RequestDate1, "yyyy-mm-dd") & "' " & _
'            "and pom.po_date <= '" & Format(RequestDate2, "yyyy-mm-dd") & "' "
    
    'NO PO tidak di filter berdasarkan supplier [W0-008 12 Juni 2007]
    sqlno = "select pom.PO_No from PurchaseOrder_Master pom " & _
            "where pom.others_cls = '0' and pom.period is null " & _
            "and pom.po_date >= '" & Format(requestdate1, "yyyy-mm-dd") & "' " & _
            "and pom.po_date <= '" & Format(requestdate2, "yyyy-mm-dd") & "' "
    Set rsno = Db.Execute(sqlno)
    With CboPOnO
        .clear
        Do While Not rsno.EOF
            .AddItem Trim(rsno("PO_No"))
            rsno.MoveNext
        Loop
        .ColumnWidths = "150pt"
        .ListWidth = 150
        .ListRows = 15
    End With
    Set rsno = Nothing
End Sub

Sub PONO(ByVal thn As String, ByVal bln As String)
Dim sqlno As String, SqlS As String
Dim rsno As New Recordset, rsS As New Recordset
    
    'POYYMM999
    If Trim(txtPoNo.Text) = "" Then
'        If Format(podate, "YYYY-MM-01") > "2006-07-30" Then
'            sqlno = "select top 1 rtrim(PO_No) from PurchaseOrder_Master " & _
'                    "where substring(rtrim(PO_No),3,2) = '" & thn & "' and substring(rtrim(PO_No),5,2) > '07' " & _
'                    "order by right(rtrim(PO_No),5) desc"
'        Else
            sqlno = "select top 1 rtrim(PO_No) from PurchaseOrder_Master " & _
                    "where substring(rtrim(PO_No),3,2) = '" & thn & "' " & _
                    "order by right(rtrim(PO_No),5) desc"
'        End If
    Else
'        If Format(podate, "YYYY-MM-01") > "2006-07-30" Then
'            sqlno = "select top 1 rtrim(PO_No) from PurchaseOrder_Master " & _
'                    "where substring(rtrim(PO_No),3,2) = '" & thn & "' and substring(rtrim(PO_No),5,2) > '07' " & _
'                    "and Right(RTrim(PO_No), 9)  < '" & Right(Trim(txtpono.Text), 9) & "' " & _
'                    "order by right(rtrim(PO_No),5) desc"
'        Else
            sqlno = "select top 1 rtrim(PO_No) from PurchaseOrder_Master " & _
                    "where substring(rtrim(PO_No),3,2) = '" & thn & "' " & _
                    "and Right(RTrim(PO_No), 9)  < '" & Right(Trim(txtPoNo.Text), 9) & "' " & _
                    "order by right(rtrim(PO_No),5) desc"
'        End If
    End If
    Set rsno = Db.Execute(sqlno)
    If Not (rsno.BOF And rsno.EOF) Then
        txtPoNo.Text = Left(Trim(rsno(0)), 4) & bln & Format(Right(Trim(rsno(0)), 5) + 1, "0000#")
    Else
            txtPoNo.Text = "PO" & thn & bln & "00001"
    End If
    txtPoNo.locked = True
    Set rsno = Nothing
End Sub

Function seqNo() As Long
Dim sqlseqno As String
Dim rsseqno As New Recordset

    sqlseqno = "select Seq_No from PurchaseORder_Detail order by Seq_No desc"
    If rsseqno.State <> adStateClosed Then rsseqno.Close
    rsseqno.Open sqlseqno, Db, adOpenKeyset, adLockOptimistic
    
    If Not (rsseqno.BOF And rsseqno.EOF) Then
        seqNo = rsseqno!Seq_no + 1
    Else
        seqNo = 1
    End If
    Set rsseqno = Nothing
End Function

Sub kunci(l As Boolean)
    PODate.Enabled = Not l
    grid.Editable = Not l
    Command1(0).Enabled = Not l
    cboPriceCondition.Enabled = Not l
    CboPacking.Enabled = Not l
    cboTransport.Enabled = Not l
    cboInsuranceCls.Enabled = Not l
    txtMarking(0).Enabled = Not l: txtMarking(1).Enabled = Not l: txtMarking(2).Enabled = Not l
    txtMarking(3).Enabled = Not l: txtMarking(4).Enabled = Not l: txtMarking(5).Enabled = Not l
    txtremarks.Enabled = Not l: cboPaymentTerm.Enabled = Not l
    DeliveryDate.Enabled = Not l
    lblFix.Caption = "Status Fix"
    lblFix.Visible = l
End Sub

Sub ppn(ByVal d As Date)
Dim sqlppn As String
Dim rsppn As New ADODB.Recordset
    
    sqlppn = "select rate from tax_cls where tax_code='PPN' and " & _
             "start_date <= '" & Format(d, "yyyymmdd") & "' and " & _
             "end_date >= '" & Format(d, "yyyymmdd") & "' "
    Set rsppn = Db.Execute(sqlppn)
    If Not (rsppn.BOF And rsppn.EOF) Then
        isippn = IIf(IsNull(rsppn(0)), 0, CDbl(rsppn(0)))
    Else
        isippn = 0
    End If
    Set rsppn = Nothing
End Sub

Function cekprice(ByVal Baris As Integer) As Boolean
Dim sqlcp As String
Dim rsCP As New Recordset
    
    cekprice = False
    sqlcp = "select price from price_master " & _
            "where item_code = '" & grid.TextMatrix(Baris, 2) & "' and price_cls = '01' " & _
            "and (trade_code = '" & cboCust.Text & "' or trade_code = '000000') " & _
            "and start_date <= '" & Format(grid.TextMatrix(Baris, 11), "yyyymmdd") & "' " & _
            "and end_date >= '" & Format(grid.TextMatrix(Baris, 11), "yyyymmdd") & "' "
    Set rsCP = Db.Execute(sqlcp)
    If Not (rsCP.BOF And rsCP.EOF) Then
        Do While Not rsCP.EOF
            If rsCP(0) = 0 Then cekprice = True: Exit Function
            rsCP.MoveNext
        Loop
    End If
    Set rsCP = Nothing
End Function

Sub browseprice()
Dim sql2 As String, rs2 As New Recordset
Dim tgldel As String
    
    If Trim(grid.TextMatrix(actrow, 11)) = "" Then
        tgldel = Trim(grid.TextMatrix(actrow, 18))
    Else
        tgldel = Format(grid.TextMatrix(actrow, 11), "yyyymmdd")
    End If
    
    sql2 = "select trade_code, priority_cls, isnull(currency_code,'') currency_code, price, unit_cls " & _
           "from price_master " & _
           "where item_code = '" & grid.TextMatrix(actrow, 2) & "' and price_cls = '01' " & _
           "and (trade_code = '" & cboCust.Text & "' or trade_code = '000000') " & _
           "and start_date <= '" & tgldel & "' " & _
           "and end_date >= '" & tgldel & "' " & _
           "order by trade_code desc, priority_cls desc"
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
            .List(i, 0) = Format(Trim(rs2("price")), "##,##0.00###")
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
    Set rs2 = Nothing
End Sub

Sub formatprice()
Dim p1 As Byte, p2 As String, p0 As String
Dim jmldigit As Byte, jmldigit0 As Byte, j As Integer

    jmldigit = 0
    With grid
        For i = 1 To .Rows - 1
            If InStr(1, .TextMatrix(i, 14), ".") > 0 Then _
                jmldigit0 = Len(Trim(.TextMatrix(i, 14))) - InStr(1, Trim(.TextMatrix(i, 14)), ".")
            If jmldigit0 > jmldigit Then jmldigit = jmldigit0
        Next i

        For i = 1 To .Rows - 1
            p0 = Trim(.TextMatrix(i, 14))
            If InStr(1, p0, ".") > 0 Then
                p1 = Len(p0) - InStr(1, p0, ".")
                For j = 1 To jmldigit - p1
                    p2 = p0 & " "
                    p0 = p2
                Next j
            End If
            .TextMatrix(i, 14) = p0
        Next i
    End With
End Sub

Sub Header()


ColReqNo = 1
ColCodeItem = 2
ColDesc = 3
ColUnit = 4
ColQtyBox = 5
ColLot = 6
ColReqQty = 7
ColOrder = 8
ColRemaining = 9
ColDelDate = 10
ColCurr = 11
ColPrice = 12
ColAmount = 13
ColReqSeqNo = 14
colTmpQty = 15
colSeqNo = 16

    With grid
        .clear
        .Rows = 1
        .ColS = 17
        
        .ColWidth(0) = 300
        .ColWidth(ColReqNo) = 1320
        .ColWidth(ColCodeItem) = 2000
        .ColWidth(ColDesc) = 2500
        '.ColHidden(4) = True
        .ColWidth(ColUnit) = 480
        .ColWidth(ColQtyBox) = 930
        .ColWidth(ColLot) = 750
        .ColWidth(ColReqQty) = 1170
        .ColWidth(ColRemaining) = 1170
        .ColWidth(ColDelDate) = 1500
        .ColWidth(ColPrice) = 1440
        .ColWidth(ColAmount) = 1500
        .ColHidden(colTmpQty) = True
        .ColHidden(ColReqSeqNo) = True
        .ColHidden(colSeqNo) = True
        
        
        
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, ColReqNo) = "Request No"
        .TextMatrix(0, ColCodeItem) = "Product Code"
        .TextMatrix(0, ColDesc) = "Description"
        .TextMatrix(0, ColUnit) = "Unit"
        .TextMatrix(0, ColQtyBox) = "Qty / Box"
        .TextMatrix(0, ColLot) = "Lot Qty"
        .TextMatrix(0, ColReqQty) = "Request Qty"
        .TextMatrix(0, ColOrder) = "Order"
        .TextMatrix(0, ColRemaining) = "Remaining"
        .TextMatrix(0, ColDelDate) = "Delivery Date"
        .TextMatrix(0, ColCurr) = "Curr"
        .TextMatrix(0, ColPrice) = "Price"
        .TextMatrix(0, ColAmount) = "Amount"
        
        
        .Cell(flexcpAlignment, 0, 0, 0, ColReqSeqNo) = flexAlignCenterCenter
        ''.ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(5) = flexAlignCenterCenter
        For i = 6 To 10
          .ColAlignment(i) = flexAlignRightCenter
        Next i
        .ColAlignment(11) = flexAlignCenterCenter
        .ColAlignment(ColDelDate) = flexAlignLeftCenter
        
        '.RowHeight(0) = 225
    
    sampun = False
    sql = "select * FROM curr_Cls"
    Dim RS As New ADODB.Recordset
    Set RS = Nothing
    RS.Open sql, Db, adOpenDynamic, adLockBatchOptimistic
    Dim s As String
    s = ""
    While Not RS.EOF
      If s = "" Then
        s = "#" & RS!curr_cls & ";" & RS!Description
      Else
        s = s & "|#" & RS!curr_cls & ";" & RS!Description
      End If
    
    RS.MoveNext
    Wend
    RS.Close
    .ColComboList(ColCurr) = s
    
    sql = "select * FROM Unit_Cls"
    Set RS = Nothing
    RS.Open sql, Db, adOpenDynamic, adLockBatchOptimistic
    s = ""
    While Not RS.EOF
      If s = "" Then
        s = "#" & RS!Unit_cls & ";" & RS!Description
      Else
        s = s & "|#" & RS!Unit_cls & ";" & RS!Description
      End If
    
    RS.MoveNext
    Wend
    .ColComboList(ColUnit) = s
    End With
End Sub

Private Sub HeaderAdd()
    
    bteColAddCode = 0
    bteColAddName = 1
    bteColAddQty = 2
    bteColAddPrice = 3
    bteColAddAmount = 4
    
    With GridAdd
        .clear
        .Rows = 1
        .ColS = 4
        
        .FormatString = "<Item Code|<Description|>Qty|>Price|>Amount"
        
        .ColWidth(0) = 2500
        .ColWidth(1) = 3000
        .ColWidth(2) = 1000
        .ColWidth(3) = 1300
        .ColWidth(4) = 1300
    End With
    
End Sub

Sub browseitem()
Dim sqlitem As String, RsItem As New ADODB.Recordset
Dim nextperiod As Date, endtgl As Integer, endperiod As Date, tglperiod As Date
    
    LblErrMsg = ""
    Call Header
    
    If ubah = False Then
        ' Add 20090112
        TxtSubAmount.Text = 0
        TxtDisc.Text = 0
        ' ---
        txtamount.Text = 0
        txtPPN.Text = 0
        txtGrandTotal.Text = 0
    End If
    activecurr = ""
    activecurrcd = ""
    ubahgrid = False
    i = 1
    
    If cborequestno.Text <> "" And cborequestno.MatchFound Then
    tglperiod = Left(cborequestno.Column(1), 4) & "-" & Right(cborequestno.Column(1), 2) & "-01"
    nextperiod = DateAdd("m", 1, tglperiod)
    endtgl = DateDiff("d", Format(tglperiod, "yyyy-mm-01"), Format(nextperiod, "yyyy-mm-01"))
    endperiod = Year(tglperiod) & "-" & Month(tglperiod) & "-" & Format(endtgl, "0#")
    End If
    
    'Detail PO NO with Different POREQUEST NO
    sqlitem = "select a.*, " & _
              "(select description from unit_cls uc where uc.unit_cls= a.unit_cls ) unit_desc , " & _
              "(select description from curr_cls where curr_cls.Curr_cls= a.Currency_Code) Curr_desc " & _
              "From ( " & _
              "select distinct '1' No, pod.item_code, pod.unit_cls, im.item_name, " & _
              "im.finishgoodpart_cls, im.number_entering, im.number_box, im.lot_qty, im.orderpoint_qty, " & _
              "isnull(prd.qty,0) RequestQty, " & _
              "isnull( (select sum(qty) qty from PurchaseOrder_Detail pod2 " & _
              "         inner join PurchaseOrder_Master pom on pom.po_no = pod.po_no " & _
              "         where pod2.item_code = pod.item_code and pod2.porequest_no = pod.PORequest_No and pod2.POReq_SeqNo = pod.POReq_SeqNo " & _
              "         and isnull(pom.others_cls,'0')='0') ,0) totalPOQty, " & _
              "isnull(prd.qty,0) - isnull( (select sum(qty) qty from PurchaseOrder_Detail pod2 " & _
              "                             inner join PurchaseOrder_Master pom on pom.po_no = pod.po_no " & _
              "                             where pod2.item_code = pod.item_code and pod2.porequest_no = pod.PORequest_No " & _
              "                             and pod2.POReq_SeqNo = pod.POReq_SeqNo and isnull(pom.others_cls,'0')='0') " & _
              ",0) RemainingQty, " & _
              "prd.ReqDelivery_Date1, pod.PORequest_No, pod.POReq_SeqNo,  prd.accountno, pod.Currency_Code, isnull(pod.Price,0) Price "
              
   sqlitem = sqlitem & " from PurchaseOrder_Detail pod " & _
              "inner join item_master im on im.item_code = pod.item_Code " & _
              "left outer join (select prd.*, isnull(prm.complete_cls,'0') Complete_Cls, " & _
              "                 cast(year(reqdelivery_date) as char(4)) + " & _
              "                 cast((case when month(reqdelivery_date) < 10 then '0' else '' end) + cast(month(reqdelivery_date) as char) as char(2)) + " & _
              "                 cast((case when day(reqdelivery_date) < 10 then '0' else '' end) + cast(day(reqdelivery_date) as char) as char(2)) as ReqDelivery_Date1 " & _
              "                 from PORequest_Detail prd " & _
              "                 inner join (select PORequest_No, Complete_Cls from PORequest_Master where isnull(others_cls,'0') = '0') prm " & _
              "                     on prm.porequest_no = prd.porequest_no) prd " & _
              "on prd.PORequest_No = pod.PORequest_No and prd.POReq_SeqNo = pod.POReq_SeqNo " & _
              "where pod.PO_No = '" & Trim(txtPoNo.Text) & "' and pod.PORequest_No <> (case prd.Complete_Cls when '1' then '' else '" & Trim(cborequestno.Text) & "' end) "
    
    'from PRICE MASTER and selected POREQUEST No
    sqlitem = sqlitem & _
              "UNION " & _
              "select distinct '2' No, pm.item_code, im.unit_cls, im.item_name, " & _
              "im.finishgoodpart_cls, im.number_entering, im.number_box, im.lot_qty, im.orderpoint_qty, " & _
              "isnull(prd.qty,0) RequestQty, " & _
              "isnull( (select sum(qty) qty from PurchaseOrder_Detail pod " & _
              "         inner join PurchaseOrder_Master pom on pom.po_no = pod.po_no " & _
              "         where pod.item_code = pm.item_code and pod.porequest_no = '" & Trim(cborequestno.Text) & "' and pod.POReq_Seqno = prd.POReq_SeqNo and isnull(pom.others_cls,'0')='0') " & _
              ",0) totalPOQty, " & _
              "isnull(prd.qty,0) - isnull( (select sum(qty) qty from PurchaseOrder_Detail pod " & _
              "                             inner join PurchaseOrder_Master pom on pom.po_no = pod.po_no " & _
              "                             where pod.item_code = pm.item_code and pod.porequest_no = '" & Trim(cborequestno.Text) & "' and pod.POReq_Seqno = prd.POReq_SeqNo and isnull(pom.others_cls,'0')='0') " & _
              ",0) RemainingQty, " & _
              "prd.ReqDelivery_Date1, prd.PORequest_No, prd.POReq_SeqNo,  prd.accountno, " & _
              "(select top 1 currency_code from Price_Master p " & _
              " where p.price_cls = '01' and p.item_code = pm.item_code and p.start_date <= prd.ReqDelivery_Date1 and p.end_date >= prd.ReqDelivery_Date1 " & _
              " and p.trade_code in ('" & Trim(cboCust.Text) & "','000000') order by p.trade_Code desc, p.priority_Cls desc) Currency_Code, " & _
              "(select top 1 price from Price_Master p " & _
              " where p.price_cls = '01' and p.item_code = pm.item_code and p.start_date <= prd.ReqDelivery_Date1 and p.end_date >= prd.ReqDelivery_Date1 " & _
              " and p.trade_code in ('" & Trim(cboCust.Text) & "','000000') order by p.trade_Code desc, p.priority_Cls desc) Price " & _
              " From Price_Master pm " & _
              "inner join Item_Master im on pm.item_code = im.item_code "
    sqlitem = sqlitem & _
              "inner join (select prd.*, isnull(prm.complete_cls,'0') Complete_Cls, " & _
              "            cast(year(reqdelivery_date) as char(4)) + " & _
              "            cast((case when month(reqdelivery_date) < 10 then '0' else '' end) + cast(month(reqdelivery_date) as char) as char(2)) + " & _
              "            cast((case when day(reqdelivery_date) < 10 then '0' else '' end) + cast(day(reqdelivery_date) as char) as char(2)) as ReqDelivery_Date1" & _
              "            from PORequest_Detail prd " & _
              "            inner join (select PORequest_No,Complete_cls from PORequest_Master where isnull(others_cls,'0') = '0') prm " & _
              "                 on prm.porequest_no = prd.porequest_no) prd " & _
              "on prd.item_code = pm.item_code " & _
              "where pm.price_cls = '01' and prd.Complete_Cls = '0' " & _
              "and pm.start_date <= prd.ReqDelivery_Date1 and pm.end_date >= prd.ReqDelivery_Date1 " & _
              "and pm.trade_code in ('" & Trim(cboCust.Text) & "','000000') and prd.porequest_no = '" & Trim(cborequestno.Text) & "' "

    'From ITEM MASTER and selected POREQUEST NO
    sqlitem = sqlitem & _
              "UNION " & _
              "select distinct '3' No, im.item_code, im.unit_cls, im.item_name, " & _
              "im.finishgoodpart_cls, im.number_entering, im.number_box, im.lot_qty, im.orderpoint_qty, isnull(prd.qty,0) RequestQty, " & _
              "isnull( (select sum(qty) qty from PurchaseOrder_Detail pod inner join PurchaseOrder_Master pom on pom.po_no = pod.po_no " & _
              "         where pod.item_code = im.item_code and pod.porequest_no = '" & Trim(cborequestno.Text) & "' and pod.POReq_Seqno = prd.POReq_SeqNo and isnull(pom.others_cls,'0')='0') ,0) totalPOQty, " & _
              "isnull(prd.qty,0) - isnull( (select sum(qty) qty from PurchaseOrder_Detail pod " & _
              "                             inner join PurchaseOrder_Master pom on pom.po_no = pod.po_no " & _
              "                             where pod.item_code = im.item_code and pod.porequest_no = '" & Trim(cborequestno.Text) & "' and pod.POReq_Seqno = prd.POReq_SeqNo and isnull(pom.others_cls,'0')='0') ,0) RemainingQty, " & _
              "prd.ReqDelivery_Date1, prd.PORequest_No, prd.POReq_SeqNo,  prd.accountno, '' Currency_Code, Null Price " & _
              " from item_master im " & _
              "inner join (select prd.*, isnull(prm.complete_cls,'0') Complete_Cls, " & _
              "            cast(year(reqdelivery_date) as char(4)) + " & _
              "            cast((case when month(reqdelivery_date) < 10 then '0' else '' end) + cast(month(reqdelivery_date) as char) as char(2)) + " & _
              "            cast((case when day(reqdelivery_date) < 10 then '0' else '' end) + cast(day(reqdelivery_date) as char) as char(2)) as ReqDelivery_Date1" & _
              "            from PORequest_Detail prd " & _
              "            inner join (select PORequest_No, Complete_Cls from PORequest_Master where isnull(others_cls,'0') = '0') prm " & _
              "                 on prm.porequest_no = prd.porequest_no) prd " & _
              "on prd.item_code = im.item_code " & _
              "where im.supplier_code = '" & Trim(cboCust.Text) & "' and prd.porequest_no = '" & Trim(cborequestno.Text) & "' and im.use_endday >= '" & Format(endperiod, "yyyymmdd") & "' " & _
              "and im.item_code not in " & _
              "    (select distinct pm2.item_Code From Price_Master pm2 " & _
              "     where pm2.price_cls = '01' and pm2.start_date <= prd.ReqDelivery_Date1 and pm2.end_date >= prd.ReqDelivery_Date1 " & _
              "     and pm2.trade_code in ('" & Trim(cboCust.Text) & "','000000') ) " & _
              "and prd.Complete_Cls = '0' "
    
    sqlitem = sqlitem & ") a "
    Set RsItem = Db.Execute(sqlitem)
    If Not (RsItem.BOF And RsItem.EOF) Then
        With grid
        Do While Not RsItem.EOF
    
          .Rows = .Rows + 1
            .Cell(flexcpBackColor, i, 0) = &HFFFFFF
            .Cell(flexcpChecked, i, 0) = flexUnchecked
            .TextMatrix(i, ColReqNo) = IIf(IsNull(RsItem("PORequest_No")), "", Trim(RsItem("PORequest_No")))
            .TextMatrix(i, ColCodeItem) = IIf(IsNull(RsItem("item_code")), "", Trim(RsItem("item_code")))
            .TextMatrix(i, ColDesc) = IIf(IsNull(RsItem("item_name")), "", Trim(RsItem("item_name")))
            .TextMatrix(i, ColUnit) = is_null(RsItem("Unit_cls"))
            
            
            .TextMatrix(i, ColLot) = IIf(IsNull(RsItem("lot_qty")), 0, Format(RsItem("lot_qty"), "#,##0"))
            
            '.TextMatrix(i, ColQtyBox) = IIf(IsNull(RsItem("TotalPOQty")), 0, Format(RsItem("TotalPOQty"), "#,##0.#0"))
            
            If RsItem("finishgoodpart_cls") = "01" Then
                .TextMatrix(i, ColQtyBox) = IIf(IsNull(RsItem("number_entering")), 0, Format(RsItem("number_entering"), "##,##0"))
            Else
                .TextMatrix(i, ColQtyBox) = IIf(IsNull(RsItem("number_box")), 0, Format(RsItem("number_box"), "##,##0"))
            End If
            
            .TextMatrix(i, ColReqQty) = IIf(IsNull(RsItem("RequestQTY")), 0, Format(RsItem("RequestQTY"), "#,##0.#0"))
            .Cell(flexcpBackColor, i, ColOrder) = &HFFFFFF
            .TextMatrix(i, ColDelDate) = RsItem!ReqDelivery_Date1 ' IIf(IsNull(RsItem("ReqDelivery_Date1")), "", Format(RsItem("ReqDelivery_Date1"), "MMM dd yyyy"))
            .TextMatrix(i, ColDelDate) = FormatDate(.TextMatrix(i, ColDelDate))
            .TextMatrix(i, ColRemaining) = IIf(IsNull(RsItem("RemainingQty")), 0, Format(RsItem("RemainingQty"), "#,##0.#0"))
            .TextMatrix(i, ColCurr) = is_null(RsItem("currency_code")) '
            .TextMatrix(i, ColPrice) = IIf(IsNull(RsItem("PRice")), 0, Format(RsItem("Price"), "#,##0.#0"))
             'simpan u/ perbandingan apakah telah disisi
             .TextMatrix(i, colTmpQty) = 0
             .TextMatrix(i, colSeqNo) = 0
            
            .Cell(flexcpBackColor, i, ColAmount) = &HFFFFFF
            .Cell(flexcpFontName, i, ColAmount) = "Courier New"
            .TextMatrix(i, ColAmount) = CDbl(.TextMatrix(i, ColPrice)) * CDbl(.TextMatrix(i, ColReqQty))   'Amount
            .TextMatrix(i, ColReqSeqNo) = is_null(RsItem("POReq_SeqNo"))
  
          
            
            RsItem.MoveNext
            
            i = i + 1
        Loop
        End With
    Else
        LblErrMsg = DisplayMsg(4006)
    End If
    Set RsItem = Nothing
End Sub
Function FormatDate(Data)
If Data <> "" Or Not IsNull(Data) Then
Data = Right(Data, 2) & "/" & Left(Right(Data, 4), 2) & "/" & Left(Data, 4)
FormatDate = Format(Data, "MMM-dd-yyyy")
Else
FormatDate = ""
End If
End Function
Sub Browse()
Dim a As Double

    LblErrMsg = ""
    
    sql = "select * from PurchaseOrder_Master where PO_No = '" & txtPoNo.Text & "' and others_cls = '0' and period is null"
    If RS.State <> adStateClosed Then RS.Close
    RS.Open sql, Db, adOpenKeyset, adLockOptimistic
    
    If Not (RS.BOF And RS.EOF) Then
        ada = True: ubah = True
        statusfix = IIf(IsNull(RS("fix_cls")), 0, RS("fix_cls"))
        Call browseitem
        Call BrowseGrid
        Call BrowseGridAdd
        Call formatprice
        
        'Count TOTAL AMOUNT
        For i = 1 To grid.Rows - 1
            If grid.Cell(flexcpChecked, i, 0) = flexChecked Then _
                a = a + grid.TextMatrix(i, ColAmount)
        Next i
        ' Add 20090112
        TxtSubAmount.Text = a
        If (TxtSubAmount.Text <> 0) Then TxtSubAmount.Text = Format(TxtSubAmount.Text, "##,##0.#0")
        
        txtamount.Text = CDbl(TxtSubAmount) - CDbl(TxtDisc)
        If (txtamount.Text <> 0) Then txtamount.Text = Format(txtamount.Text, "##,##0.#0")
        ' ---
        
        If ((cboCust.Column(3) = 1) Or (cboCust.Column(3) = 2) Or (cboCust.Column(3) = 3) Or (cboCust.Column(3) = 5)) Then
            txtPPN = 0
        Else
            txtPPN.Text = CDbl(isippn / 100) * CDbl(txtamount.Text)
        End If
        If (txtPPN.Text <> 0) Then txtPPN.Text = Format(txtPPN.Text, "##,##0.#0")
        txtGrandTotal = CDbl(txtPPN.Text) + CDbl(txtamount.Text)
        If (txtGrandTotal.Text <> 0) Then txtGrandTotal.Text = Format(txtGrandTotal.Text, "##,##0.#0")
        
        If statusfix = 1 Then Call kunci(True) Else Call kunci(False)
'        cboPacking.Text = is_null(rs!POPacking_Cls)
'        cboPaymentTerm.Text = is_null(rs!PaymentTerm_Cls)
'        cboprice.Text = is_null(rs!pricecondition_cls)
'        cboInsuranceCls = is_null(rs!Insurance_Cls)
'        cboTransport = is_null(rs!Transportation_Cls)
'        txtMarking(0) = is_null(rs!POMarking1)
'        txtMarking(1) = is_null(rs!POMarking2)
'        txtMarking(2) = is_null(rs!POMarking3)
'        txtMarking(3) = is_null(rs!POMarking4)
'        txtMarking(4) = is_null(rs!POMarking5)
'        txtMarking(5) = is_null(rs!POMarking6)
        
        
    Else
        ada = False
    End If
End Sub

Sub BrowseGrid()
Dim g As Integer
    
    sqlGrid = " select (select description from unit_cls uc where uc.unit_cls= PurchaseOrder_Detail.unit_cls ) unit_desc , " & _
              " (select description from curr_cls where curr_cls.Curr_cls= PurchaseOrder_Detail.Currency_Code) Curr_desc, " & _
              " * from PurchaseOrder_Detail where PO_No = '" & Trim(txtPoNo.Text) & "' order by item_code"
    If rsGrid.State <> adStateClosed Then rsGrid.Close
    rsGrid.Open sqlGrid, Db, adOpenKeyset, adLockOptimistic

    With grid
    Do While Not rsGrid.EOF
        For g = 1 To .Rows - 1
            If Trim(.TextMatrix(g, 2)) = Trim(rsGrid("Item_Code")) And Trim(.TextMatrix(g, 1)) = Trim(rsGrid("PORequest_No")) And Trim(.TextMatrix(g, ColReqSeqNo)) = CStr(rsGrid("POReq_SeqNo")) Then
                ubahgrid = True
                .Cell(flexcpChecked, g, 0) = flexChecked
                .TextMatrix(g, ColUnit) = Trim(rsGrid("Unit_cls"))
                .TextMatrix(g, ColOrder) = IIf(IsNull(rsGrid("qty")), 0, Format(Trim(rsGrid("qty")), "##,##0.#0"))
                .TextMatrix(g, ColDelDate) = IIf(IsNull(rsGrid("Delivery_Date")), "", Format(rsGrid("Delivery_Date"), "dd MMM yyyy"))
                .TextMatrix(g, ColCurr) = is_null(Trim(rsGrid("currency_code")))  '
                .TextMatrix(g, ColPrice) = IIf(IsNull(rsGrid("Price")), 120, Format(Trim(rsGrid("Price")), "##,##0.00###"))
                .TextMatrix(g, ColAmount) = IIf(IsNull(rsGrid("Amount")), 0, Format(rsGrid("Amount"), "##,##0.#0"))
                 .TextMatrix(g, colTmpQty) = IIf(IsNull(rsGrid("qty")), 0, Format(Trim(rsGrid("qty")), "##,##0.#0"))
                 .TextMatrix(g, colSeqNo) = is_null(rsGrid!Seq_no)
            End If
            
        Next g
        rsGrid.MoveNext
    Loop
    End With
End Sub

Private Sub BrowseGridAdd()
    
    Dim adoRs As New ADODB.Recordset
    
    HeaderAdd
    
    With GridAdd
        
        sql = "Select pd.Item_Code, im.Item_Name, pd.Qty, Isnull(pd.Price_Service, 0) Price_Service, Isnull(pd.Amount_Service, 0) Amount_Service " & _
            "From PurchaseOrder_Detail pd " & _
            "Inner Join Item_Master im On pd.Item_Code = im.Item_Code " & _
            "Where pd.PO_No = '" & txtPoNo & "'"
        
        adoRs.Open sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
        While Not adoRs.EOF
            .AddItem ""
            
            .TextMatrix(.Rows - 1, bteColAddCode) = Trim(adoRs.Fields("Item_Code") & "")
            .TextMatrix(.Rows - 1, bteColAddName) = Trim(adoRs.Fields("Item_Name") & "")
            .TextMatrix(.Rows - 1, bteColAddQty) = Format(Val(adoRs.Fields("Qty") & ""), gs_formatQty)
            .TextMatrix(.Rows - 1, bteColAddPrice) = Format(Val(adoRs.Fields("Price_Service") & ""), gs_formatPrice)
            .TextMatrix(.Rows - 1, bteColAddAmount) = Format(Val(adoRs.Fields("Amount_Service") & ""), gs_formatAmount)
            
'            .Cell(flexcpBackColor, .Rows - 1, bteColAddPrice) = vbWhite
            
            adoRs.MoveNext
        Wend
        adoRs.Close
        
    End With
    
End Sub

Sub BrowseAtas()
    sql = "select * from PurchaseOrder_Master where PO_No = '" & Trim(txtPoNo.Text) & "' and isnull(others_cls,'0') = '0' and period is null"
    If RS.State <> adStateClosed Then RS.Close
    RS.Open sql, Db, adOpenKeyset, adLockOptimistic

    If Not (RS.BOF And RS.EOF) Then
        PODate.Value = IIf(IsNull(RS("po_date")), "", Format(Trim(RS("po_date")), "dd MMM yyyy"))
        cboCust.Text = Trim(RS("Supplier_code"))
        txtRev.Text = is_null(RS!Revise_No)
        CboPacking.Text = is_null(RS!POPacking_Cls)
        cboPaymentTerm.Text = is_null(RS!PaymentTerm_Cls)
        cboPriceCondition.Text = is_null(RS!PriceCondition_Cls)
        cboInsuranceCls = is_null(RS!Insurance_Cls)
        cboTransport = is_null(RS!Transportation_Cls)
        txtMarking(0) = is_null(RS!POMarking1)
        txtMarking(1) = is_null(RS!POMarking2)
        txtMarking(2) = is_null(RS!POMarking3)
        txtMarking(3) = is_null(RS!POMarking4)
        txtMarking(4) = is_null(RS!POMarking5)
        txtMarking(5) = is_null(RS!POMarking6)
        txtremarks = is_null(RS!Remarks)
        TxtDisc = Format(RS!discount, gs_formatAmount) ' Add 20090112
        If Not IsNull(RS!delivery_Date) Then DeliveryDate.Value = (RS!delivery_Date)
    
        If Not IsNull(RS!po_date) Then PODate.Value = (RS!po_date)
         
        statusfix = IIf(IsNull(RS("fix_cls")), 0, RS("fix_cls"))
        If statusfix = 1 Then Call kunci(True) Else Call kunci(False)
    End If
End Sub

Function cekrecqty(ItemCode As String, PONO As String) As Double
Dim sqlcek As String, rsCek As New Recordset
    
    cekrecqty = 0
    sqlcek = "select item_code, sum(qty) recqty from Part_Receipt " & _
             "where PO_No = '" & Trim(PONO) & "' and item_code = '" & Trim(ItemCode) & "' " & _
             "group by item_code "
    If rsCek.State <> adStateClosed Then rsCek.Close
    rsCek.Open sqlcek, Db, adOpenKeyset, adLockOptimistic
    If Not (rsCek.BOF And rsCek.EOF) Then _
        cekrecqty = CDbl(rsCek("recqty"))
    Set rsCek = Nothing
End Function

Private Sub cboInsuranceCls_Change()
If cboInsuranceCls.MatchFound Then txtInsurance.Text = cboInsuranceCls.Column(1) Else txtInsurance.Text = ""
End Sub

Private Sub cboPacking_Change()
If CboPacking.MatchFound Then TxtPacking.Text = CboPacking.Column(1) Else TxtPacking.Text = ""
End Sub

Private Sub cboPaymentTerm_Change()
If cboPaymentTerm.MatchFound Then txtPaymentTerm.Text = cboPaymentTerm.Column(1) Else txtPaymentTerm.Text = ""
End Sub

Private Sub cborequestno_Change()
cborequestno_Click
End Sub

Private Sub cboTransport_Change()
If cboTransport.MatchFound Then TxtTransport.Text = cboTransport.Column(1) Else TxtTransport.Text = ""
End Sub

Private Sub cmdSearch_Click()
    Dim i As Double
    
    LblErrMsg = ""
    
    If txtSearch = "" Or grid.Rows = 2 Then txtSearch.SetFocus: Exit Sub
    If grid.Row = grid.Rows - 1 Then i = 2 Else i = grid.Row + 1
    
    Do
        Select Case cboSearch.ListIndex
        Case 0
            grid.Col = ColCodeItem
            If UCase(Mid(grid.TextMatrix(i, ColCodeItem), 1, Len(txtSearch))) = UCase(txtSearch) Then
                Exit Do
            End If
        Case 1
            grid.Col = ColDesc
            If InStr(UCase(grid.TextMatrix(i, ColDesc)), UCase(txtSearch)) <> 0 Then
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

Private Sub Form_Load()
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"

    Call adtocboCust(False)
    Call adtocombo
    Call Kosong
    combo1.ListIndex = 1
    Header
    HeaderAdd
End Sub

Private Sub requestdate1_Change()
Dim ketemu As Boolean

    LblErrMsg.Caption = ""
    If Format(requestdate1, "yyyy-mm-dd") > Format(requestdate2, "yyyy-mm-dd") Then
       LblErrMsg.Caption = DisplayMsg(4025) & " " & Format(requestdate2, "MMM yyyy") '"Start Date must be lower than "
       Exit Sub
    End If
    
    Call adtocborequestno
        
    If combo1.ListIndex = 1 Then    'UPDATE
        If cboCust.Text <> "" Then Call adtocbopono
        For i = 0 To CboPOnO.ListCount - 1
            If txtPoNo.Text = CboPOnO.List(i) Then
                ketemu = True
                CboPOnO.ListIndex = i
                Exit For
            End If
        Next i
        If ketemu = False Then txtPoNo.Text = ""
    End If
    
'    txtpono.Text = ""
'    cbopono.clear
    Call Header
    Call kosongBwh
End Sub

Private Sub requestdate2_Change()
Dim ketemu As Boolean

    LblErrMsg.Caption = ""
    If Format(requestdate2, "yyyy-mm-01") < Format(requestdate1, "yyyy-mm-01") Then
       LblErrMsg.Caption = DisplayMsg(4024) & " " & Format(requestdate1, "MMM yyyy") '"End Date must be higher than "
       Exit Sub
    End If

    Call adtocborequestno
    
    If combo1.ListIndex = 1 Then    'UPDATE
        If cboCust.Text <> "" Then Call adtocbopono
        For i = 0 To CboPOnO.ListCount - 1
            If txtPoNo.Text = CboPOnO.List(i) Then
                ketemu = True
                CboPOnO.ListIndex = i
                Exit For
            End If
        Next i
        If ketemu = False Then txtPoNo.Text = ""
    End If

'    txtpono.Text = ""
'    cbopono.clear
    Call Header
    Call kosongBwh
End Sub

Private Sub cborequestno_Click()
Dim ketemu As Boolean, tempcust As String

    LblErrMsg = ""
    If cborequestno.ListIndex <> -1 Then
        If Trim(cboCust.Text) <> "" Then tempcust = Trim(cboCust.Text)
        If cborequestno.Text <> "" Then
            Call adtocboCust(True)
            
            Call adtocbopono
        Else
            Call adtocboCust(False)
        End If
        cboCust.Text = tempcust: Call cboCust_Click
        
        If combo1.ListIndex = 1 Then    'UPDATE
            If cboCust.Text <> "" Then Call adtocbopono
            For i = 0 To CboPOnO.ListCount - 1
                If txtPoNo.Text = CboPOnO.List(i) Then
                    ketemu = True
                    CboPOnO.ListIndex = i
                    Exit For
                End If
            Next i
            If ketemu = False Then txtPoNo.Text = ""
            Call kosongBwh: Call Header
        End If
    Else
        CboPOnO.clear
        If Trim(cboCust.Text) <> "" Then tempcust = Trim(cboCust.Text)
        If cborequestno.Text <> "" Then Call adtocboCust(True) Else Call adtocboCust(False)
        cboCust.Text = tempcust: Call cboCust_Click
        
        If combo1.ListIndex = 1 Then 'UPDATE
            Call kosongBwh
            Call Header
            txtPoNo.Text = ""
        End If
        If cborequestno.Text <> "" Then LblErrMsg.Caption = DisplayMsg(4144) '"Record with this Request No not found"
    End If
End Sub

Private Sub cborequestno_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call cborequestno_Click
End Sub

Private Sub cborequestno_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub cboCust_Click()
Dim ketemu As Boolean

    LblErrMsg = ""
    ketemu = False
    Call kunci(False)

    If cboCust.ListIndex <> -1 Then
        lblcust(0).Text = cboCust.Column(1)
        lblcust(1).Text = cboCust.Column(2)
        countrycls = cboCust.Column(3)
            'If Split(cbocust.Column(5), ",")(0) = "" Then cbopocode.ListIndex = -1 Else cbopocode.ListIndex = Split(cbocust.Column(5), ",")(0) - 1
        'txtpoday.Text = Split(cbocust.Column(5), ",")(1)
        
        
        
        If combo1.ListIndex = 1 Then    'UPDATE
            Call ClearData
            Call adtocbopono
            For i = 0 To CboPOnO.ListCount - 1
                If txtPoNo.Text = CboPOnO.List(i) Then
                    ketemu = True
                    'cbopono.ListIndex = i
                    Exit For
                End If
            Next i
            If ketemu = False Then txtPoNo.Text = ""
            Call kosongBwh
            Call Header
        End If
    Else
        lblcust(0).Text = ""
        lblcust(1).Text = ""
        countrycls = 3
        
        
        
        
        'cbopono.clear
        If combo1.ListIndex = 1 Then 'UPDATE
            Call kosongBwh
            Call Header
            txtPoNo.Text = ""
        End If
        If cboCust.Text <> "" Then LblErrMsg.Caption = DisplayMsg(4050) '"Record with this Supplier Code not Exist"
        Exit Sub
    End If
        
    If (countrycls = 1 Or countrycls = 2) Then  'OVERSEAS
        isippn = 0
        txtPPN.Text = 0
        txtGrandTotal.Text = txtamount.Text
        If (txtGrandTotal.Text <> 0) Then txtGrandTotal.Text = Format(txtGrandTotal.Text, "##,##0.#0")
    Else 'DOMESTIC
        Call ppn(PODate.Value)
        txtPPN.Text = CDbl(isippn / 100) * CDbl(txtamount.Text)
        If (txtPPN.Text <> 0) Then txtPPN.Text = Format(txtPPN.Text, "##,##0.#0")
        txtGrandTotal = CDbl(txtPPN.Text) + CDbl(txtamount.Text)
        If (txtGrandTotal.Text <> 0) Then txtGrandTotal.Text = Format(txtGrandTotal.Text, "##,##0.#0")
    End If
End Sub

Private Sub cbocust_LostFocus()
    If sampun = False Then Call cboCust_Click   'sampun->false=tidak ada data di grid
End Sub

Private Sub cboCust_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call cboCust_Click
End Sub

Private Sub cbocust_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Combo1_Click()
Dim ketemu As Boolean

    LblErrMsg = ""
    ketemu = False
    Call kunci(False)
    Call kosongBwh
    Call Header

    If combo1.ListIndex = 0 Then    'CREATE
        Call ClearData
        Command1(2).Caption = "&Create"
        ubah = False
        CboPOnO.locked = True
        txtPoNo.Text = ""
        PODate.Value = Format(Now, "dd MMM yyyy")
        PODate.Enabled = False
        Call PONO(Right(Year(PODate), 2), Format(Month(PODate), "0#"))
        cboPriceCondition.ListIndex = -1
        
    Else    'UPDATE
        If cboCust.Text = "" Then   'Or cborequestno.Text = ""
            CboPOnO.clear
            txtPoNo.Text = ""
        Else
            Call adtocbopono
        End If
        ubah = True
        Command1(2).Caption = "&Update"
        CboPOnO.locked = False
        txtPoNo.locked = False
        PODate.Enabled = True

        For i = 0 To CboPOnO.ListCount - 1
            If txtPoNo.Text = CboPOnO.List(i) Then
                ketemu = True
                CboPOnO.ListIndex = i
                Exit For
            End If
        Next i
        If ketemu = False Then txtPoNo.Text = ""
    End If
End Sub

Private Sub combo1_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call Combo1_Click
End Sub

Private Sub cbopono_Click()
    LblErrMsg = ""
    txtPoNo.Text = CboPOnO.Text
    Call Header
    Call kosongBwh
    Call BrowseAtas

End Sub

Private Sub cbopono_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call cbopono_Click
End Sub

Private Sub TxtDisc_Change()
    If InStr(1, TxtDisc.Text, ",") = 1 Then TxtDisc.Text = Right(TxtDisc, Len(TxtDisc) - 1)
    If TxtDisc <> "" And TxtSubAmount <> "" And IsNumeric(TxtDisc) And IsNumeric(TxtSubAmount) Then txtamount.Text = Format(CDbl(TxtSubAmount.Text) - CDbl(TxtDisc.Text), "##,##0.#0")
End Sub

Private Sub TxtDisc_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then _
          KeyAscii = 0
    If InStr(1, TxtDisc, ".") > 0 Then If KeyAscii = Asc(".") Then KeyAscii = 0
End Sub

Private Sub TxtDisc_LostFocus()
txtamount = CDbl(TxtSubAmount) - CDbl(TxtDisc)
TxtDisc = Format(TxtDisc, gs_formatAmount)
End Sub

Private Sub txtpono_Change()
Dim ketemu As Boolean

    txtPONo2.Text = txtPoNo.Text
    If combo1.ListIndex = 1 Then
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

Private Sub txtpono_KeyPress(KeyAscii As Integer)
    LblErrMsg = ""
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = Asc("'") Then KeyAscii = 0
    If KeyAscii = 13 Then
        If combo1.ListIndex = 0 Then
            SendKeys vbTab
        Else
            Call Header
            Call kosongBwh
            Call BrowseAtas
        End If
    End If
End Sub

Private Sub PODate_Change()
    LblErrMsg = ""
    'CREATE
    If combo1.ListIndex = 0 Then _
        Call PONO(Right(Year(PODate), 2), Format(Month(PODate), "0#"))
    If (countrycls = 1 Or countrycls = 2) Then isippn = 0 Else Call ppn(PODate.Value)
End Sub

Private Sub deldate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys vbTab
End Sub

Private Sub deldate_LostFocus()
    DelDate.Visible = False
End Sub

Private Sub cbocurr_Click()
    If cbocurr.ListIndex <> -1 Then
        grid.TextMatrix(actrow, 12) = cbocurr.Column(0)
        grid.TextMatrix(actrow, 13) = cbocurr.Column(1)
    End If
End Sub

Private Sub cbocurr_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call cbocurr_Click
End Sub

Private Sub cbocurr_LostFocus()
    cbocurr.Visible = False
End Sub

Private Sub cboprice_Change()
    If InStr(1, cboprice.Text, ",") = 1 Then cboprice.Text = Right(cboprice, Len(cboprice) - 1)
End Sub

Private Sub cboprice_Click()
    If cboprice.ListIndex <> -1 Then
        grid.TextMatrix(actrow, 12) = cboprice.Column(2)
        For i = 0 To cbocurr.ListCount - 1
            If Trim(grid.TextMatrix(actrow, 12)) = Trim(cbocurr.List(i)) Then
                cbocurr.ListIndex = i
                Exit For
            End If
        Next i
        If Trim(cboprice.Column(2)) <> "" Then grid.TextMatrix(actrow, 13) = uf_GetCurrencyDescription(cboprice.Column(2))
        'Grid.TextMatrix(actrow, 4) = cboprice.Column(3)
        'Grid.TextMatrix(actrow, 5) = uf_GetUnitDescription(cboprice.Column(3))
    End If
End Sub

Private Sub cboprice_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call cboprice_Click
End Sub

Private Sub CboPrice_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then _
          KeyAscii = 0
    If InStr(1, cboprice.Text, ".") > 1 Then _
        If KeyAscii = Asc(".") Then KeyAscii = 0
End Sub

Private Sub cboPrice_LostFocus()
Dim ketemu As Boolean
Dim z
    
    If cboprice.Text = "" Then cboprice.Text = 0
    z = CDec(cboprice.Text)
    If z > 9999999999.99 Then cboprice.Text = Left(z, 10)
        
    grid.TextMatrix(actrow, 14) = Format(cboprice.Text, "#,##0.00###")
    Call Grid_AfterEdit(actrow, 14)
    
    cboprice.Text = Format(cboprice.Text, "#,##0.00###")
    For i = 0 To cboprice.ListCount - 1
        If Trim(cboprice.Text) = Trim(cboprice.List(i)) Then
            ketemu = True
            cboprice.ListIndex = i
            Exit For
        End If
    Next i
    cboprice.Visible = False
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim a As Double
Dim banding As Double
Dim intRow As Integer
    a = 0
    With grid
        If Col = ColOrder Then   'ORDER QTY
            If .TextMatrix(Row, ColOrder) = "" Then .TextMatrix(Row, ColOrder) = 0
            If IsNumeric(.TextMatrix(Row, ColOrder)) = False Then .TextMatrix(Row, ColOrder) = 0
            If CDbl(.TextMatrix(Row, ColOrder)) > 9999999.99 Then LblErrMsg = DisplayMsg(4045) & " 9,999,999.99": .SetFocus: Exit Sub   '"Quantity must be lower or equal than 9,999,999.99"
            .TextMatrix(Row, ColOrder) = Format(.TextMatrix(Row, ColOrder), "#,##0.#0")
        End If
        
        If Col = 0 Or Col = ColOrder Then
           .TextMatrix(Row, ColAmount) = 500 * IS_NOL((.TextMatrix(Row, ColOrder)))
            .TextMatrix(Row, ColAmount) = Format(.TextMatrix(Row, ColAmount), "##,##0.#0")
        End If
        'membandingkan apakah edit atau tambah baru
        If Col = ColOrder Then
          If .TextMatrix(Row, colTmpQty) = 0 Then
            banding = CDbl(IS_NOL(.TextMatrix(Row, ColRemaining)))
          Else
            banding = CDbl(IS_NOL(.TextMatrix(Row, ColRemaining))) + CDbl(IS_NOL(.TextMatrix(Row, colTmpQty)))
          End If
            If (banding) < IS_NOL(CDbl(.TextMatrix(Row, ColOrder))) Then
                .TextMatrix(Row, Col) = 0
                LblErrMsg = DisplayMsg(4045) & " " & banding
            
            Else
            LblErrMsg = ""
            
            End If
            
            ' Tambahan dari GridAdd
            If Col = bteColSelect Then
                If .Cell(flexcpChecked, Row, bteColSelect) = flexChecked Then
                    GridAdd.AddItem ""
                    GridAdd.TextMatrix(GridAdd.Rows - 1, bteColAddCode) = .TextMatrix(Row, bteColProdCode)
                    GridAdd.TextMatrix(GridAdd.Rows - 1, bteColAddName) = .TextMatrix(Row, bteColDesc)
                    GridAdd.TextMatrix(GridAdd.Rows - 1, bteColAddQty) = .TextMatrix(Row, bteColOrder)
                    'GridAdd.TextMatrix(GridAdd.Rows - 1, bteColAddPrice) = Format(GetServicePrice(.TextMatrix(Row, bteColProdCode), .TextMatrix(Row, bteColCurrCode)), gs_formatPrice)
                    GridAdd.TextMatrix(GridAdd.Rows - 1, bteColAddAmount) = Format(CDbl(GridAdd.TextMatrix(GridAdd.Rows - 1, bteColAddPrice)) * CDbl(.TextMatrix(Row, bteColOrder)), gs_formatAmount)
                    GridAdd.Cell(flexcpBackColor, GridAdd.Rows - 1, bteColAddPrice) = vbWhite
                Else
                    For intRow = 1 To GridAdd.Rows - 1
                        If Trim(GridAdd.TextMatrix(intRow, bteColAddCode)) = Trim(.TextMatrix(Row, bteColProdCode)) Then
                            GridAdd.RemoveItem (intRow)
                            Exit For
                        End If
                    Next
                End If
            End If
            
            If Col = bteColOrder Then
                For intRow = 1 To GridAdd.Rows - 1
                    If Trim(GridAdd.TextMatrix(intRow, bteColAddCode)) = Trim(.TextMatrix(Row, bteColProdCode)) Then
                        GridAdd.TextMatrix(intRow, bteColAddQty) = Format(CDbl(.TextMatrix(Row, bteColOrder)), gs_formatQty)
                        GridAdd.TextMatrix(intRow, bteColAddAmount) = Format(CDbl(GridAdd.TextMatrix(intRow, bteColAddPrice)) * CDbl(.TextMatrix(Row, bteColOrder)), gs_formatAmount)
                    End If
                Next
            End If
            
            
        End If
        
    End With
End Sub
Function IS_NOL(Data)
If IsNull(Data) Then
Data = 0
Exit Function
End If
If Data = "" Then
IS_NOL = 0
Else
IS_NOL = Data
End If

End Function
Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    actrow = Row
    If grid.Cell(flexcpChecked, Row, 0) <> flexChecked Then
        If Col <> 0 Then Cancel = True
    Else
        If Col <> 0 And Col <> ColOrder And Col <> colRemark Then
            Cancel = True
       End If
    End If
End Sub

Private Sub Grid_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = ColOrder Then _
        If InStr(1, grid.TextMatrix(Row, Col), ",") = 1 Then grid.TextMatrix(Row, Col) = Right(grid.TextMatrix(Row, Col), Len(grid.TextMatrix(Row, Col)) - 1)
End Sub

Private Sub grid_Click()
Dim reqdel As Date

    With grid
    If statusfix = 0 Then
        If .Row > 0 Then
            If .Cell(flexcpChecked, .Row, 0) = flexChecked Then
                If .Col = ColReqQty Then
                    .SelectionMode = flexSelectionFree
                Else
                    .SelectionMode = flexSelectionByRow
                End If
                
                
                If .Col = ColOrder Then .FocusRect = flexFocusInset Else .FocusRect = flexFocusNone
            Else
                .SelectionMode = flexSelectionByRow
            End If
        End If
    End If
    End With
End Sub

Private Sub Grid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    LblErrMsg = ""
    If Col = 9 Then
        If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then _
            KeyAscii = 0
    End If
End Sub

Private Sub grid_AfterSort(ByVal Col As Long, Order As Integer)
    DelDate.Visible = False
End Sub

Private Sub Grid_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call grid_Click
End Sub

Private Sub cbopricecondition_Click()
    If cboPriceCondition.ListIndex <> -1 Then
    '    txtPriceCondition.Text = cboPriceCondition.Column(1)
    Else
        txtPriceCondition.Text = ""
    End If
End Sub

Private Sub cbopricecondition_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call cbopricecondition_Click
End Sub

Private Sub cbopricecondition_Change()
    If cboPriceCondition.MatchFound Then txtPriceCondition.Text = cboPriceCondition.Column(1) Else txtPriceCondition.Text = ""
End Sub

Private Sub txtpoday_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyReturn Then _
        KeyAscii = 0
    If KeyAscii = Asc("'") Or KeyAscii = Asc(".") Then KeyAscii = 0
End Sub

Private Sub txtpoday_LostFocus()
   
End Sub

Private Sub hscrollbar_Change()
End Sub

Private Sub hscrollbar_Scroll()
    Call hscrollbar_Change
End Sub

Private Sub Command1_Click(Index As Integer)
'On Error GoTo Msg_Error
Dim sql4 As String
Dim rs4 As New Recordset

LblErrMsg = ""

Select Case Index
    Case 0: 'SUBMIT
            If hakUpdate(Me.Name) = 0 Then _
                LblErrMsg = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
            
            'HEADER VALIDATION
            If cborequestno.Text = "" Then
                cborequestno.SetFocus
               LblErrMsg = DisplayMsg(8121) '"Please Input Request No"
                Exit Sub
            ElseIf cborequestno.Text <> "" Then
                If cborequestno.MatchFound = False Then
                    LblErrMsg = DisplayMsg(4144)    'Record with This Request No Not found !
                    cborequestno.SetFocus
                    Exit Sub
                End If
            End If
            If cboCust.Text = "" Then
                cboCust.SetFocus
                LblErrMsg = DisplayMsg(1045) '"Please Select Supplier Code"
                Exit Sub
            ElseIf cboCust.Text <> "" Then
                If cboCust.MatchFound = False Then
                    LblErrMsg = DisplayMsg(4050)    'Record with This Supplier Code Not found !
                    cboCust.SetFocus
                    Exit Sub
                End If
            End If
            If txtPoNo.Text = "" Then
                txtPoNo.SetFocus
                LblErrMsg = DisplayMsg(1046) '"Please Input PO No"
                Exit Sub
            End If
                            
            'FOOTER VALIDATION
            If CekFooter(True) Then Exit Sub
            
            '-------------------------------------------------------------
            
            sql = "select * from PurchaseOrder_Master where PO_No = '" & txtPoNo.Text & "' and others_cls = '0' and period is null"
            If RS.State <> adStateClosed Then RS.Close
            RS.Open sql, Db, adOpenKeyset, adLockOptimistic
            If RS.BOF And RS.EOF Then
                LblErrMsg.Caption = DisplayMsg(4015)    'Record with This PO No not found
                txtPoNo.SetFocus
                Exit Sub
            End If
            
            If grid.Rows = 1 Then LblErrMsg = DisplayMsg(4047): Exit Sub  'There is no data to submit !
            
            If ubah = True Then
                If Not ValidDataSupplier(txtPoNo.Text) Then
                    LblErrMsg = "Can't change supplier! System found item[s] which no have price for this supplier. Please Input Price Master First!"
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                RS("po_date") = Format(PODate.Value, "yyyy-mm-dd")
                RS("discount") = CDbl(TxtDisc.Text) ' Add 20090112
                RS("amount") = CDbl(txtamount.Text)
                RS("ppn") = CDbl(txtPPN.Text)
                RS("total_amount") = CDbl(txtGrandTotal.Text)
                RS("remarks") = Trim(txtremarks.Text)
                RS("revise_no") = Right(Trim(txtRev.Text), 1)
                RS("pricecondition_cls") = is_null(cboPriceCondition.Text)
                RS("PaymentTerm_Cls") = is_null(cboPaymentTerm.Text)
                RS("POPacking_cls") = is_null(CboPacking.Text)
                RS("Insurance_Cls") = is_null(cboInsuranceCls.Text)
                RS("Transportation_cls") = is_null(cboTransport.Text)
                RS!POMarking1 = is_null(txtMarking(0))
                RS!POMarking2 = is_null(txtMarking(1))
                RS!POMarking3 = is_null(txtMarking(2))
                RS!POMarking4 = is_null(txtMarking(3))
                RS!POMarking5 = is_null(txtMarking(4))
                RS!POMarking6 = is_null(txtMarking(5))
                RS!po_date = PODate.Value
                RS!delivery_Date = DeliveryDate.Value
                
                RS.update

                Dim recqty As Double
                                
                'DETAIL VALIDATION
                With grid
                    For i = 1 To .Rows - 1
                        If .Cell(flexcpChecked, i, 0) = flexChecked Then
                            If .TextMatrix(i, ColOrder) = 0 Then
                               
                                .SelectionMode = flexSelectionFree
                                .Col = 9: .Row = i: actrow = i
                                .SetFocus
                                LblErrMsg = DisplayMsg(1012) '"Please Input Quantity"
                                Exit Sub
                            ElseIf CDbl(.TextMatrix(i, ColOrder)) > 9999999.99 Then
                               
                                .SelectionMode = flexSelectionFree
                                .Col = 9: .Row = i: actrow = i
                                .SetFocus
                                LblErrMsg = DisplayMsg(4045) & " 9,999,999.99" '"Quantity must be lower or equal than 9,999,999.99"
                                Exit Sub
                            ElseIf CDbl(.TextMatrix(i, ColRemaining)) < 0 Then
                               
                                .SelectionMode = flexSelectionFree
                                .Col = 9: .Row = i: actrow = i
                                .SetFocus
                                LblErrMsg = DisplayMsg(4045) & " Qty Remaining" '"Quantity must be lower or equal than Qty Remaining"
                                Exit Sub
                            End If
                            
                            recqty = cekrecqty(.TextMatrix(i, 2), txtPoNo.Text)
                            If CDbl(.TextMatrix(i, ColOrder)) < recqty Then
                                
                                .SelectionMode = flexSelectionFree
                                .Col = 9: .Row = i: actrow = i
                                .SetFocus
                                LblErrMsg = DisplayMsg(4036) & " " & recqty     '"Quantity must be higher or equal than "
                                Exit Sub
                            End If
                        Else
                            sql4 = "select * from Part_Receipt where po_no = '" & txtPoNo.Text & "' and item_code = '" & .TextMatrix(i, 2) & "'"
                            Set rs4 = Db.Execute(sql4)
                            If Not (rs4.BOF And rs4.EOF) Then
                                .Row = i: actrow = i
                                .SetFocus
                                LblErrMsg = DisplayMsg(1204)
                                Exit Sub
                            End If
                            Set rs4 = Nothing
                        End If
                    Next i
                                                                                                    
                    ' Cek currency harus sama
'                    Dim a As Integer, c As Integer
'                    If ubahgrid = False Or activecurr = "" Then  '*****
'                        For a = 1 To .Rows - 1
'                            If .Cell(flexcpChecked, a, 0) = flexChecked Then
'                                activecurr = .TextMatrix(a, 13): activecurrcd = .TextMatrix(a, 12): Exit For
'                            End If
'                        Next a
'                    End If
'                    For c = a + 1 To .Rows - 1
'                        If .Cell(flexcpChecked, c, 0) = flexChecked Then
'                            If .TextMatrix(c, 13) <> .TextMatrix(1, 13) Then
'                                hscrollbar.Value = 1
'                                .Col = 13: .Row = c: actrow = c
'                                .SetFocus
'                                Call Grid_Click
'                                LblErrMsg = DisplayMsg(4084)  'Cannot Select Different Currency !!
'                                Exit Sub
'                            End If
'                        End If
'                    Next c

                    Dim a As Integer, C As Integer
                    If ubahgrid = False Or activecurr = "" Then  '*****
                        For a = 1 To .Rows - 1
                            If .Cell(flexcpChecked, a, 0) = flexChecked Then
                                activecurr = .TextMatrix(a, 13): activecurrcd = .TextMatrix(a, 12): Exit For
                            End If
                        Next a
                    End If
                    
                    Dim barisan As Integer, indexawal As Integer, barisawal As Integer, jumlahBarisan As Integer
                    indexawal = 0
                    jumlahBarisan = 0
                    For barisan = 1 To .Rows - 1
                        If .Cell(flexcpChecked, barisan, 0) = flexChecked Then
                                If indexawal = 0 Then barisawal = barisan: indexawal = 1
                                jumlahBarisan = jumlahBarisan + 1
                        End If
                    Next barisan
                    
                    For barisan = 1 To .Rows - 1
                        If .Cell(flexcpChecked, barisan, 0) = flexChecked Then
                        Else
                            .TextMatrix(barisan, 13) = ""
                        End If
                    Next barisan
                    
            If jumlahBarisan > 1 Then
                    For C = a + 1 To .Rows - 1
                        If .Cell(flexcpChecked, C, 0) = flexChecked Then
                            If .TextMatrix(C, ColCurr) <> .TextMatrix(barisawal, ColCurr) Then
                                
                                .Col = 13: .Row = C: actrow = C
                                .SetFocus
                                Call grid_Click
                                LblErrMsg = DisplayMsg(4084)  'Cannot Select Different Currency !!
                                Exit Sub
                            End If
                        End If
                    Next C
            End If
                    
                    Dim rscekC As New Recordset
                    Dim RTmp As Double
                    'UPDATE DETAIL
                    For i = 1 To .Rows - 1
                        If .Cell(flexcpChecked, i, 0) = flexChecked Then
                            Set rsGrid = Nothing
                            sqlGrid = "select * From PurchaseOrder_Detail where Seq_No =" & .TextMatrix(i, colSeqNo) & ""
                            rsGrid.Open sqlGrid, Db, 1, 3
                            If rsGrid.EOF Then
                                rsGrid.AddNew
                                rsGrid("seq_No") = seqNo
                                RTmp = 0
                            Else
                            RTmp = rsGrid!Qty
                            End If
                            rsGrid("po_no") = Trim(txtPoNo.Text)
                            rsGrid("PORequest_No") = Trim(.TextMatrix(i, 1))
                            rsGrid("item_Code") = Trim(.TextMatrix(i, 2))
                            rsGrid("Delivery_Date") = Format(.TextMatrix(i, ColDelDate), "yyyy-mm-dd")
                            rsGrid("price") = CDbl(IS_NOL(.TextMatrix(i, ColPrice)))
                            rsGrid("currency_code") = .TextMatrix(i, ColCurr)
                            rsGrid("unit_cls") = .TextMatrix(i, ColUnit)
                            rsGrid("qty") = CDbl(.TextMatrix(i, ColOrder)) '+ CDbl(RTmp)
                            rsGrid("amount") = CDbl(.TextMatrix(i, ColAmount))
                            rsGrid!POReq_seqno = CDbl(.TextMatrix(i, ColReqSeqNo))
                            
                            
                            
                            rsGrid.update
                            rsGrid.Close
                            
                        Else
                            If Trim(.TextMatrix(i, colSeqNo)) <> "" Then
                                sql = "Delete from PurchaseOrder_Detail where seq_no = '" & .TextMatrix(i, colSeqNo) & "'"
                                Db.Execute sql
                            End If
                        End If
                        
                        'UPDATE COMPLETE_CLS (POREQUEST_MASTER)
                        sql = "select PORequest_No, avg(DComplete) Complete " & _
                              "from ( select prd.PORequest_No, prd.poreq_seqno, isnull(prd.Qty,0) Qty, isnull(sum(pod.Qty),0) POQty, " & _
                              "        (case when isnull(prd.Qty,0) = isnull(sum(pod.Qty),0) then 1 else 0 end) DComplete " & _
                              "        from PORequest_Detail prd " & _
                              "        left outer join PurchaseOrder_Detail pod on pod.PORequest_No = prd.PORequest_No and pod.POReq_SeqNo = prd.poreq_seqno" & _
                              "        group by prd.PORequest_No, prd.poreq_seqno, prd.Qty ) a " & _
                              "where PORequest_No = '" & Trim(.TextMatrix(i, 1)) & "' " & _
                              "group by PORequest_No "
                        If rscekC.State <> adStateClosed Then rscekC.Close
                        rscekC.Open sql, Db, adOpenKeyset, adLockOptimistic
                        If Not (rscekC.BOF And rscekC.EOF) Then
                            If rscekC("Complete") = "1" Then
                                sql = "Update PORequest_Master set Complete_Cls = '1' where PORequest_No = '" & Trim(rscekC("PORequest_No")) & "' "
                            ElseIf rscekC("Complete") = "0" Then
                                sql = "Update PORequest_Master set Complete_Cls = '0' where PORequest_No = '" & Trim(rscekC("PORequest_No")) & "' "
                            End If
                            Db.Execute sql
                        End If
                    Next i
                    Call updateMaster(True)
                    Call CekPONumber
                    Call Browse
                    LblErrMsg = DisplayMsg(1101)
                    ubahgrid = True
                End With
          End If

    Case 1: 'CLEAR
            Call Kosong
            combo1.ListIndex = 1
            Call Combo1_Click
            cborequestno.SetFocus
            
    Case 2: 'CREATE / UPDATE
            If combo1.ListIndex = 0 Then    'CREATE
                If hakUpdate(Me.Name) = 0 Then _
                    LblErrMsg = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
                
                'HEADER VALIDATION
                
                If cborequestno.Text = "" Then
                    cborequestno.SetFocus
                    LblErrMsg = DisplayMsg(8121) '"Please Input Request No"
                    Exit Sub
                ElseIf cborequestno.Text <> "" Then
                    If cborequestno.MatchFound = False Then
                        LblErrMsg = DisplayMsg(4144)    'Record with This Request No Not found !
                        cborequestno.SetFocus
                        Exit Sub
                    End If
                End If

                If cboCust.Text = "" Then
                    cboCust.SetFocus
                    LblErrMsg = DisplayMsg(1054) '"Please Select Supplier Code"
                    Exit Sub
                ElseIf cboCust.Text <> "" Then
                    If cboCust.MatchFound = False Then
                        LblErrMsg = DisplayMsg(4050)    'Record with This Supplier Code Not found !
                        cboCust.SetFocus
                        Exit Sub
                    End If
                End If
                
                If txtPoNo.Text = "" Then
                    txtPoNo.SetFocus
                    LblErrMsg = DisplayMsg(1046) '"Please Input PO No"
                    Exit Sub
                End If
                                
                'FOOTER VALIDATION
                If cboPriceCondition.Text <> "" Then
                    If cboPriceCondition.MatchFound = False Then
                        LblErrMsg = DisplayMsg(4147)    'Record with This Price Condition not found !
                        cboPriceCondition.SetFocus
                        Exit Sub
                    End If
                End If
                If CekFooter(False) Then Exit Sub
                '-------------------------------------------------------------
                
                If ubah = False Then
                    sql = "select * from PurchaseOrder_Master where PO_No = '" & txtPoNo.Text & "'"
                    If RS.State <> adStateClosed Then RS.Close
                    RS.Open sql, Db, adOpenKeyset, adLockOptimistic
                    If Not (RS.BOF And RS.EOF) Then
                        LblErrMsg.Caption = DisplayMsg(1023)
                        txtPoNo.SetFocus
                        Exit Sub
                    Else
                        RS.AddNew
                        RS("po_no") = txtPoNo.Text
                        RS("supplier_code") = cboCust.Text
                    End If
                End If
                    
                RS("po_date") = Format(PODate.Value, "yyyy-mm-dd")
                RS("discount") = CDbl(TxtDisc.Text) ' Add 20090112
                RS("amount") = CDbl(txtamount.Text)
                RS("ppn") = CDbl(txtPPN.Text)
                RS("total_amount") = CDbl(txtGrandTotal.Text)
                RS("remarks") = Trim(txtremarks.Text)
                
                RS("others_cls") = "0"
                
                
On Error Resume Next
                RS.update
errHandler:
                If InStr(1, err.Description, "Violation of PRIMARY KEY constraint") > 0 Then
                    Call PONO(Right(Year(PODate), 2), Format(Month(PODate), "0#"))
                    txtPONo2.Text = txtPoNo.Text
                    RS("po_No") = txtPoNo.Text
                    RS.update
                    If InStr(1, err.Description, "Violation of PRIMARY KEY constraint") > 0 Then
                        GoTo errHandler
                    Else
                        If Trim$(err.Description) <> "" Then
                            LblErrMsg = Trim$(err.number) + " : " + Trim$(err.Description)
                            Exit Sub
                        End If
                    End If
                Else
                    If Trim$(err.Description) <> "" Then
                        LblErrMsg = Trim$(err.number) + " : " + Trim$(err.Description)
                        Exit Sub
                    End If
                End If
    
                If CDate(PODate.Value) > CDate(requestdate1.Value) Then
                    If CDate(PODate.Value) > CDate(requestdate2.Value) Then _
                        requestdate2.Value = Format(PODate.Value, "dd MMM yyyy")
                Else
                    requestdate1.Value = Format(PODate.Value, "dd MMM yyyy")
                End If
                
                combo1.Text = "Update"
                If cboCust.Text <> "" And cborequestno.Text <> "" Then Call browseitem: Call formatprice
                LblErrMsg.Caption = DisplayMsg(1000)
                ubah = True
    
            Else    'UPDATE
            If sampun Then                      'sampun=true->ada data di grid; false->tidak ada data di grid
                Call updateMaster(False)
            Else
                Dim ketemu As Boolean
                
                If cborequestno.Text = "" Then
'                    cboRequestNo.SetFocus
'                    LblErrMsg = DisplayMsg(1067) '"Please Input Request No"
'                    Exit Sub
                ElseIf cborequestno.Text <> "" Then
                    If cborequestno.MatchFound = False Then
                        LblErrMsg = DisplayMsg(4144)    'Record with This Request No Not found !
                        cborequestno.SetFocus
                        Exit Sub
                    End If
                End If
                If cboCust.Text = "" Then
                    cboCust.SetFocus
                    LblErrMsg = DisplayMsg(1054) '"Please Select Supplier Code"
                    Exit Sub
                ElseIf cboCust.Text <> "" Then
                    If cboCust.MatchFound = False Then
                        LblErrMsg = DisplayMsg(4050)    'Record with This Supplier Code Not found !
                        cboCust.SetFocus
                        Exit Sub
                    End If
                End If
                
                If txtPoNo.Text = "" Then
                    txtPoNo.SetFocus
                    LblErrMsg = DisplayMsg(1046) '"Please Input PO No"
                    Exit Sub
                End If
                
                If CDate(PODate.Value) > CDate(requestdate1.Value) Then
                    If CDate(PODate.Value) > CDate(requestdate2.Value) Then _
                        requestdate2.Value = Format(PODate.Value, "dd MMM yyyy")
                Else
                    requestdate1.Value = Format(PODate.Value, "dd MMM yyyy")
                End If
    
                If cboCust.Text = "" Then   'Or cborequestno.Text = ""
                    CboPOnO.clear: txtPoNo.Text = ""
                Else
                    Call adtocbopono
                End If
                For i = 0 To CboPOnO.ListCount - 1
                    If txtPoNo.Text = CboPOnO.List(i) Then
                        ketemu = True
                        Exit For
                    End If
                Next
                If ketemu = False Then GoTo here
                If Not ValidDataSupplier(txtPoNo.Text) Then
                    LblErrMsg = "Can't change supplier! System found item[s] which no have price for this supplier. Please Input Price Master First!"
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                
                Call Browse
                sampun = True  ' true = Ada data di grid
                Call updateMaster(False)
            End If
                If ada = False Then
here:
                    Call kosongBwh
                    txtremarks.Text = ""
                    Call Header
                    LblErrMsg.Caption = DisplayMsg(4015)    'Record with This PO No not found
                    txtPoNo.SetFocus
                    Exit Sub
                End If
            End If
    
    Case 3: 'CANCEL
            If txtPoNo.Text <> "" And cboCust.Text <> "" Then   'And cboRequestNo.Text <> ""
                For i = 0 To CboPOnO.ListCount - 1
                    If txtPoNo.Text = CboPOnO.List(i) Then
                        ketemu = True
                        Exit For
                    End If
                Next i
                If ketemu = False Then
                    Call kosongBwh
                    txtremarks.Text = ""
                    Call Header
                    LblErrMsg.Caption = DisplayMsg(4015)    'Record with This PO No not found
                    txtPoNo.SetFocus
                    Exit Sub
                End If
                Call BrowseAtas
                Call Browse
            End If
End Select

Exit Sub
Msg_Error:
LblErrMsg = err.number & " " & err.Description

End Sub
Function is_null(Data)
If IsNull(Data) Then
is_null = ""
Exit Function
End If
If Data = "" Then
is_null = Null
Else
is_null = Trim(Data)
End If

End Function
Function CekFooter(nilai As Boolean)
CekFooter = False
If nilai = True Then
    If cboPaymentTerm.Text = "" Then
    LblErrMsg = DisplayMsg(8123)
    cboPaymentTerm.SetFocus
    CekFooter = True
    Exit Function
    End If
    If cboPriceCondition.Text = "" Then
        LblErrMsg = DisplayMsg(8129)    'Record with This Price Condition not found !
        cboPriceCondition.SetFocus
        CekFooter = True
        Exit Function
     End If
     If cboPriceCondition.Text = "" Then
        LblErrMsg = DisplayMsg(8129)    'Record with This Price Condition not found !
        cboPriceCondition.SetFocus
        CekFooter = True
        Exit Function
      End If
      
      If cboInsuranceCls.Text = "" Then
         LblErrMsg = "INSURANCE CLS" 'DisplayMsg(4147)    'Record with This Price Condition not found !
         cboInsuranceCls.SetFocus
         CekFooter = True
         Exit Function
       End If
       If CboPacking.Text = "" Then
            LblErrMsg = DisplayMsg(34)         'Record with This Price Condition not found !
            CboPacking.SetFocus
            CekFooter = True
            Exit Function
       End If
                
       If cboPriceCondition.Text = "" Then
         LblErrMsg = DisplayMsg(8129)
         cboPriceCondition.SetFocus
         CekFooter = True
         Exit Function
       End If
        If cboTransport.Text = "" Then
           LblErrMsg = DisplayMsg(8130)
           cboTransport.SetFocus
         CekFooter = True
         Exit Function
        End If
        
End If

If cboPaymentTerm.Text <> "" Then
    If cboPaymentTerm.MatchFound = False Then
    LblErrMsg = DisplayMsg(8050)
    cboPaymentTerm.SetFocus
    CekFooter = True
    Exit Function
    End If
    End If
If cboPriceCondition.Text <> "" Then
    If cboPriceCondition.MatchFound = False Then
        LblErrMsg = DisplayMsg(8051)    'Record with This Price Condition not found !
        cboPriceCondition.SetFocus
        CekFooter = True
    Exit Function
    End If
End If
      
      If cboInsuranceCls.Text <> "" Then
       If cboInsuranceCls.MatchFound = False Then
         LblErrMsg = "Record with This Insurance Clas not found " 'DisplayMsg(4147)    'Record with This Price Condition not found !
         cboInsuranceCls.SetFocus
         CekFooter = True
         Exit Function
       End If
      End If
       If CboPacking.Text <> "" Then
         If CboPacking.MatchFound = False Then
            LblErrMsg = DisplayMsg(4010)         'Record with This Price Condition not found !
            CboPacking.SetFocus
            CekFooter = True
            Exit Function
          End If
       End If
                
       
        If cboTransport.Text <> "" Then
         If cboTransport.MatchFound = False Then
           LblErrMsg = DisplayMsg(8059)
            cboTransport.SetFocus
           CekFooter = True
           Exit Function
        End If
        End If


            
End Function


Private Sub cmdReport_Click()
On Error GoTo ErrMsg
Dim application As New CRAXDDRT.application
Dim report As New CRAXDDRT.report
Dim rsRpt As New ADODB.Recordset
Dim Rpt As New FrmRpt3
Dim sqlcekdet As String, SqlRpt As String
Dim rscekdet As New Recordset
LblErrMsg = ""
    If combo1.ListIndex = 1 And txtPoNo.Text <> "" And cboCust.Text <> "" Then
        sqlcekdet = "select pom.PO_No from PurchaseOrder_Master pom " & _
                    "inner join PurchaseOrder_Detail pod on pod.PO_No = pom.PO_No " & _
                    "where pom.others_cls = '0' and pom.period is null " & _
                    "and pom.PO_No = '" & Trim(txtPoNo.Text) & "' and pom.supplier_Code = '" & Trim(cboCust.Text) & "'"
        Set rscekdet = Db.Execute(sqlcekdet)
        If rscekdet.EOF Then LblErrMsg.Caption = DisplayMsg(4006): Me.MousePointer = vbDefault: Exit Sub
        Set rscekdet = Nothing
        
        Me.MousePointer = vbHourglass

'        If cbocust.Column(4) = 1 Then   'PO CLS=YES
'            If cborequestno.Text = "" Then
''                cboRequestNo.SetFocus
''                LblErrMsg = DisplayMsg(1067) '"Please Input Request No"
''                Me.MousePointer = vbDefault
''                Exit Sub
'            ElseIf cborequestno.Text <> "" Then
'                If cborequestno.MatchFound = False Then
'                    lblErrMsg = DisplayMsg(4144)    'Record with This Request No Not found !
'                    cborequestno.SetFocus
'                    Me.MousePointer = vbDefault
'                    Exit Sub
'                End If
'            End If
'
''            Dim nextperiod As Date, endtgl As Integer, endperiod As Date, tglperiod As Date
''            'tglperiod = Left(cboRequestNo.Column(1), 4) & "-" & Right(cboRequestNo.Column(1), 2) & "-01"
''            'nextperiod = DateAdd("m", 1, tglperiod)
''            'endtgl = DateDiff("d", Format(tglperiod, "yyyy-mm-01"), Format(nextperiod, "yyyy-mm-01"))
''            'endperiod = year(tglperiod) & "-" & month(tglperiod) & "-" & Format(endtgl, "0#")
''
'            'PURCHASE ORDER DETAIL
'            SqlRpt = "select '1' Sort, rtrim(pom.po_no) po_no, pom.po_date, rtrim(pom.supplier_Code) Supplier_Code, " & _
'                    " trade_name = case when CHARINDEX ( ',' , rtrim(tm.trade_name)) <> 0 then " & _
'                    " ltrim(substring(rtrim(tm.trade_name), CHARINDEX ( ',' , rtrim(tm.trade_name)) + 1, 5)) + " & _
'                    " '.' + left(rtrim(tm.trade_name),CHARINDEX ( ',' , rtrim(tm.trade_name)) - 1) " & _
'                    " Else " & _
'                    " RTrim (tm.trade_name) " & _
'                    " End, "
'            SqlRpt = SqlRpt + " " & _
'                     "rtrim(tm.address1) taddress1, rtrim(tm.address2) taddress2, rtrim(tm.city) tcity, rtrim(tm.postal_code) tpostal_code, " & _
'                     "rtrim(tm.contact_person) contact_person, isnull(rtrim(tm.telephone),'') Supplierphone, isnull(rtrim(tm.Fax),'') SupplierFax, tm.POPayment_code, tm.POPayment_Day, pom.PaymentTerm_cls, " & _
'                     "rtrim(pod.PORequest_No) PORequest_No, pod.POReq_SeqNo, rtrim(pod.item_code) item_code, rtrim(im.item_name) item_name, " & _
'                     "pod.unit_cls, (select description from unit_cls uc where uc.unit_cls= pod.unit_cls ) unit_desc, isnull(pod.qty,0) Qty, pod.currency_code, (select description from curr_cls where curr_cls.Curr_cls= pod.Currency_Code) Curr_desc, isnull(pod.price,0) Price, isnull(pod.amount,0) Amount, " & _
'                     "pod.Delivery_Date, pom.PriceCondition_Cls, rtrim(pc.Description) PriceCondition, pom.Transportation_Cls, " & _
'                     "rtrim(pom.remarks) Remarks, isnull(pom.amount,0) as TAmount, isnull(pom.ppn,0) PPN, isnull(pom.total_amount,0) Total_Amount, " & _
'                     "rtrim(cp.company_name) company_name, rtrim(cp.address1) caddress1, rtrim(cp.address2) caddress2, " & _
'                     "rtrim(cp.Province) cprovince, rtrim(cp.City) ccity, rtrim(cp.postal_code) cpostal_code, rtrim(cp.phone1) cphone1, " & _
'                     "rtrim(cp.phone2) cphone2, rtrim(cp.fax) cfax, rtrim(cp.PO_position) po_position, rtrim(cp.PO_person) po_person, " '& _
'                     "rtrim(cp.POAcknowledge_Person) POAcknowledge_Person, rtrim(cp.POAcknowledge_Position) POAcknowledge_Position, " & _
'                     "rtrim(cp.POApproved_Person) POApproved_Person, rtrim(cp.POApproved_Position) POApproved_Position,
'
'            SqlRpt = SqlRpt + " " & _
'                     "tm.Trade_Cls, tm.Country_Cls, " & _
'                     "from PurchaseOrder_Master pom " & _
'                     "inner join PurchaseOrder_Detail pod on pod.PO_No = pom.PO_No " & _
'                     "left outer join Item_Master im on im.item_code = pod.Item_code " & _
'                     "left outer join Trade_Master tm on tm.trade_code = pom.supplier_code " & _
'                     "left outer join PriceCondition_Cls pc on pc.PriceCondition_Cls = pom.PriceCondition_Cls " & _
'                     "cross join Company_Profile cp " & _
'                     "where pom.po_no = '" & Trim(txtpono.Text) & "' and pom.others_cls = '0' and pom.period is null "
'
'            'from PRICE MASTER and selected POREQUEST No
'            SqlRpt = SqlRpt & _
'                     "UNION " & _
'                     "select '2' Sort, rtrim(pom.po_no) po_no, pom.po_date, rtrim(pom.supplier_Code) Supplier_Code, " & _
'                    " trade_name = case when CHARINDEX ( ',' , rtrim(tm.trade_name)) <> 0 then " & _
'                    " ltrim(substring(rtrim(tm.trade_name), CHARINDEX ( ',' , rtrim(tm.trade_name)) + 1, 5)) + " & _
'                    " '.' + left(rtrim(tm.trade_name),CHARINDEX ( ',' , rtrim(tm.trade_name)) - 1) " & _
'                    " Else " & _
'                    " RTrim (tm.trade_name) " & _
'                    " End, "
'            SqlRpt = SqlRpt + " " & _
'                     "rtrim(tm.address1) taddress1, rtrim(tm.address2) taddress2, rtrim(tm.city) tcity, rtrim(tm.postal_code) tpostal_code, " & _
'                     "rtrim(tm.contact_person) contact_person, isnull(rtrim(tm.telephone),'') Supplierphone, isnull(rtrim(tm.Fax),'') SupplierFax, tm.POPayment_code, tm.POPayment_Day, pom.PaymentTerm_cls, " & _
'                     "prd.PORequest_No, prd.Seq_No, rtrim(pm.item_code) item_code, rtrim(im.item_name) item_name, " & _
'                     "im.unit_cls, (select description from unit_cls uc where uc.unit_cls= im.unit_cls ) unit_desc , 0 Qty, pm.currency_code, (select description from curr_cls where curr_cls.Curr_cls= pm.Currency_Code) Curr_desc, isnull(pm.price,0) Price, 0 Amount, " & _
'                     "Null Delivery_Date, pom.PriceCondition_Cls, (select rtrim(pc.Description) from PriceCondition_Cls pc where pc.PriceCondition_cls=pom.PriceCondition_Cls) PriceCondition, pom.Transportation_Cls, " & _
'                     "rtrim(pom.remarks) Remarks,  isnull(pom.amount,0) as TAmount, isnull(pom.ppn,0) PPN, isnull(pom.total_amount,0) Total_Amount, " & _
'                     "rtrim(cp.company_name) company_name, rtrim(cp.address1) caddress1, rtrim(cp.address2) caddress2, " & _
'                     "rtrim(cp.Province) cprovince, rtrim(cp.City) ccity, rtrim(cp.postal_code) cpostal_code, rtrim(cp.phone1) cphone1, " & _
'                     "rtrim(cp.phone2) cphone2, rtrim(cp.fax) cfax, rtrim(cp.PO_position) po_position, rtrim(cp.PO_person) po_person, " '& _
'                     "rtrim(cp.POAcknowledge_Person) POAcknowledge_Person, rtrim(cp.POAcknowledge_Position) POAcknowledge_Position, " & _
'                     "rtrim(cp.POApproved_Person) POApproved_Person, rtrim(cp.POApproved_Position) POApproved_Position,
'            SqlRpt = SqlRpt + " " & _
'                     "tm.Trade_Cls, tm.Country_Cls " & _
'                     "from PurchaseOrder_Master pom, Price_Master pm, Trade_Master tm, Item_Master im, " & _
'                     "(select prd.*, " & _
'                     "        cast(year(reqdelivery_date) as char(4)) + " & _
'                     "        cast((case when month(reqdelivery_date) < 10 then '0' else '' end) + cast(month(reqdelivery_date) as char) as char(2)) + " & _
'                     "        cast((case when day(reqdelivery_date) < 10 then '0' else '' end) + cast(day(reqdelivery_date) as char) as char(2)) " & _
'                     "        as ReqDelivery_Date1, prm.Department_Cls, isnull(prm.Complete_Cls,'0') Complete_Cls " & _
'                     "        from PORequest_Detail prd " & _
'                     "        left outer join (select PORequest_No, Department_Cls, Complete_Cls from PORequest_Master where isnull(others_cls,'0') = '0') prm " & _
'                     "        on prm.porequest_no = prd.porequest_no) prd, " & _
'                     "Company_Profile cp "
'            SqlRpt = SqlRpt & _
'                     "Where pm.item_code = im.item_code And prd.item_code = pm.item_code " & _
'                     "And pom.Supplier_Code = tm.Trade_Code " & _
'                     "and pom.others_Cls = '0' and pom.period is null and pom.PO_No = '" & Trim(txtpono.Text) & "' And tm.PO_Cls = '1' " & _
'                     "and pm.trade_code in ('" & Trim(cbocust.Text) & "','000000') and pm.price_cls = '01' " & _
'                     "and start_date <= prd.ReqDelivery_Date1 and end_date >= prd.ReqDelivery_Date1 " & _
'                     "and prd.porequest_no = '" & Trim(cborequestno.Text) & "' and prd.Complete_Cls = '0' " & _
'                     "and prd.seq_no not in (select POReq_SeqNo from PurchaseOrder_Detail where PO_No = '" & Trim(txtpono.Text) & "') " & _
'                     "and pm.currency_code = (select top 1 currency_code from PurchaseOrder_Detail where PO_No = '" & Trim(txtpono.Text) & "' and Currency_Code is not null) "
'
'            'from ITEM MASTER and selected POREQUEST No
'            SqlRpt = SqlRpt & _
'                     "UNION " & _
'                     "select '3' Sort, rtrim(pom.po_no) po_no, pom.po_date, rtrim(pom.supplier_Code) Supplier_Code, " & _
'                    " trade_name = case when CHARINDEX ( ',' , rtrim(tm.trade_name)) <> 0 then " & _
'                    " ltrim(substring(rtrim(tm.trade_name), CHARINDEX ( ',' , rtrim(tm.trade_name)) + 1, 5)) + " & _
'                    " '.' + left(rtrim(tm.trade_name),CHARINDEX ( ',' , rtrim(tm.trade_name)) - 1) " & _
'                    " Else " & _
'                    " RTrim (tm.trade_name) " & _
'                    " End, "
'            SqlRpt = SqlRpt + " " & _
'                     "rtrim(tm.address1) taddress1, rtrim(tm.address2) taddress2, rtrim(tm.city) tcity, rtrim(tm.postal_code) tpostal_code, " & _
'                     "rtrim(tm.contact_person) contact_person, isnull(rtrim(tm.telephone),'') Supplierphone, isnull(rtrim(tm.Fax),'') SupplierFax, tm.POPayment_code, tm.POPayment_Day, pom.PaymentTerm_cls, " & _
'                     "prd.PORequest_No, prd.Seq_No, rtrim(im.item_code) item_code, rtrim(im.item_name) item_name, " & _
'                     "im.unit_cls, (select description from unit_cls uc where uc.unit_cls= im.unit_cls ) unit_desc, 0 Qty, " & _
'                     "(select top 1 currency_code from PurchaseOrder_Detail where PO_No = '" & Trim(txtpono.Text) & "' and Currency_Code is not null) currency_code, (select description from curr_cls where curr_cls.Curr_cls= (select top 1 currency_code from PurchaseOrder_Detail where PO_No = '" & Trim(txtpono.Text) & "' and Currency_Code is not null)) Curr_desc, 0 Price, 0 Amount, " & _
'                     "Null Delivery_Date, pom.PriceCondition_Cls, (select rtrim(pc.Description) from PriceCondition_Cls pc where pc.PriceCondition_cls=pom.PriceCondition_Cls) PriceCondition, pom.Transportation_Cls, " & _
'                     "rtrim(pom.remarks) Remarks, isnull(pom.amount,0) as TAmount, isnull(pom.ppn,0) PPN, isnull(pom.total_amount,0) Total_Amount, " & _
'                     "rtrim(cp.company_name) company_name, rtrim(cp.address1) caddress1, rtrim(cp.address2) caddress2, " & _
'                     "rtrim(cp.Province) cprovince, rtrim(cp.City) ccity, rtrim(cp.postal_code) cpostal_code, rtrim(cp.phone1) cphone1, " & _
'                     "rtrim(cp.phone2) cphone2, rtrim(cp.fax) cfax, rtrim(cp.PO_position) po_position, rtrim(cp.PO_person) po_person, " '& _
'                     "rtrim(cp.POAcknowledge_Person) POAcknowledge_Person, rtrim(cp.POAcknowledge_Position) POAcknowledge_Position, " & _
'                     "rtrim(cp.POApproved_Person) POApproved_Person, rtrim(cp.POApproved_Position) POApproved_Position,
'
'            SqlRpt = SqlRpt + " " & _
'                     "tm.Trade_Cls, tm.Country_Cls  " & _
'                     "from PurchaseOrder_Master pom, Trade_Master tm, Item_Master im, " & _
'                     "(select prd.*, " & _
'                     "        cast(year(reqdelivery_date) as char(4)) + " & _
'                     "        cast((case when month(reqdelivery_date) < 10 then '0' else '' end) + cast(month(reqdelivery_date) as char) as char(2)) + " & _
'                     "        cast((case when day(reqdelivery_date) < 10 then '0' else '' end) + cast(day(reqdelivery_date) as char) as char(2)) " & _
'                     "        as ReqDelivery_Date1, prm.Department_cls, isnull(prm.Complete_Cls,'0') Complete_Cls " & _
'                     "        from PORequest_Detail prd " & _
'                     "        inner join (select PORequest_No, Department_Cls, Complete_Cls from PORequest_Master where isnull(others_cls,'0') = '0') prm " & _
'                     "        on prm.porequest_no = prd.porequest_no) prd, " & _
'                     "Company_Profile cp "
'                     'dicabut setelah request no and im.use_endday >= '" & Format(endperiod, "yyyymmdd") & "'
'            SqlRpt = SqlRpt & _
'                     "Where prd.item_code = im.item_code And pom.Supplier_Code = tm.Trade_Code " & _
'                     "and pom.others_Cls = '0' and pom.period is null and pom.PO_No = '" & Trim(txtpono.Text) & "' And tm.PO_Cls = '1' " & _
'                     "and im.supplier_code = '" & Trim(cbocust.Text) & "' and prd.porequest_no = '" & Trim(cborequestno.Text) & "' " & _
'                     "and prd.seq_no not in (select POReq_SeqNo from PurchaseOrder_Detail where PO_No = '" & Trim(txtpono.Text) & "') and prd.Complete_Cls = '0' " & _
'                     "and im.item_code not in " & _
'                     "    (select distinct pm2.item_Code From Price_Master pm2 " & _
'                     "     where pm2.price_cls = '01' and pm2.start_date <= prd.ReqDelivery_Date1 and pm2.end_date >= prd.ReqDelivery_Date1 " & _
'                     "     and pm2.trade_code in ('" & Trim(cbocust.Text) & "','000000') ) " & _
'                     "Order by Sort "
'
'        Else    'PO CLS=NO rtrim(tm.trade_name) trade_name,
'            SqlRpt = "select '1' Sort, rtrim(pom.po_no) po_no, pom.po_date, rtrim(pom.supplier_Code) Supplier_Code, " & _
'                    " trade_name = case when CHARINDEX ( ',' , rtrim(tm.trade_name)) <> 0 then " & _
'                    " ltrim(substring(rtrim(tm.trade_name), CHARINDEX ( ',' , rtrim(tm.trade_name)) + 1, 5)) + " & _
'                    " '.' + left(rtrim(tm.trade_name),CHARINDEX ( ',' , rtrim(tm.trade_name)) - 1) " & _
'                    " Else " & _
'                    " RTrim (tm.trade_name) " & _
'                    " End, "
'            SqlRpt = SqlRpt + " " & _
'                     "rtrim(tm.address1) taddress1, rtrim(tm.address2) taddress2, rtrim(tm.city) tcity, rtrim(tm.postal_code) tpostal_code, " & _
'                     "rtrim(tm.contact_person) contact_person, isnull(rtrim(tm.telephone),'') Supplierphone, isnull(rtrim(tm.Fax),'') SupplierFax, tm.POPayment_code, tm.POPayment_Day, pom.PaymentTerm_cls, " & _
'                     "rtrim(pod.PORequest_No) PORequest_No, pod.POReq_SeqNo, rtrim(pod.item_code) item_code, rtrim(im.item_name) item_name, " & _
'                     "pod.unit_cls, (select description from unit_cls uc where uc.unit_cls= pod.unit_cls ) unit_desc, isnull(pod.qty,0) Qty, pod.currency_code, (select description from curr_cls where curr_cls.Curr_cls= pod.Currency_Code) Curr_desc ,isnull(pod.price,0) Price, isnull(pod.amount,0) Amount, " & _
'                     "pod.Delivery_Date, pom.PriceCondition_Cls, rtrim(pc.Description) PriceCondition, pom.Transportation_Cls, " & _
'                     "rtrim(pom.remarks) Remarks, isnull(pom.amount,0) as TAmount, isnull(pom.ppn,0) PPN, isnull(pom.total_amount,0) Total_Amount, " & _
'                     "rtrim(cp.company_name) company_name, rtrim(cp.address1) caddress1, rtrim(cp.address2) caddress2, " & _
'                     "rtrim(cp.Province) cprovince, rtrim(cp.City) ccity, rtrim(cp.postal_code) cpostal_code, rtrim(cp.phone1) cphone1, " & _
'                     "rtrim(cp.phone2) cphone2, rtrim(cp.fax) cfax, rtrim(cp.PO_position) po_position, rtrim(cp.PO_person) po_person, " '& _
'                     "rtrim(cp.POAcknowledge_Person) POAcknowledge_Person, rtrim(cp.POAcknowledge_Position) POAcknowledge_Position, " & _
'                     "rtrim(cp.POApproved_Person) POApproved_Person, rtrim(cp.POApproved_Position) POApproved_Position,
'
'            SqlRpt = SqlRpt + " " & _
'                     "tm.Trade_Cls, tm.Country_Cls " & _
'                     "from PurchaseOrder_Master pom " & _
'                     "inner join PurchaseOrder_Detail pod on pod.PO_No = pom.PO_No " & _
'                     "left outer join Item_Master im on im.item_code = pod.Item_code " & _
'                     "left outer join Trade_Master tm on tm.trade_code = pom.supplier_code " & _
'                     "left outer join PriceCondition_Cls pc on pc.PriceCondition_Cls = pom.PriceCondition_Cls " & _
'                     "cross join Company_Profile cp " & _
'                     "where pom.po_no = '" & Trim(txtpono.Text) & "' and pom.others_cls = '0' and pom.period is null " & _
'                     "order by pod.PORequest_No, pod.Item_Code, pod.POReq_SeqNo "
'        End If
        
'-----------------------
 ' Untuk Format PO Scheduled Musashi
 ' ----------------------
        
SqlRpt = " Select POM.Po_No, POM.Po_Date,POM.delivery_Date,PRD.PoRequest_No,PRM.PersonInCharge_Cls,PIC.Description, " & _
            vbLf & " POM.Supplier_Code,TM.Trade_Name,TM.Contact_Person,TM.Address1,TM.Address2,TM.City,TM.Country, " & _
            vbLf & " TM.Telephone,Tm.Fax,POM.PaymentTerm_Cls, " & _
            vbLf & " POD.Item_code,POD.Price,POD.Qty,POD.Amount,IM.Item_Name, " & _
            vbLf & " POD.Unit_Cls,U.Description Unit,POD.Currency_Code,C.Description Currency," & _
            vbLf & "' ' Ref,' ' ShipVia,POM.Remarks comments, " & _
            vbLf & " isnull((Select ChildRequirement_Qty from requirement_Master Where ChildRequirement_Month=month(POD.delivery_Date)+1 and ChildRequirement_Year=year(POD.delivery_Date) and ChildItem_Code=POD.Item_code),0) F1, " & _
            vbLf & " isnull((Select ChildRequirement_Qty from requirement_Master Where ChildRequirement_Month=month(POD.delivery_Date)+2 and ChildRequirement_Year=year(POD.delivery_Date) and ChildItem_Code=POD.Item_code),0) F2 " & _
            vbLf & " From PurchaseOrder_Master POM inner join PurchaseOrder_Detail POD " & _
            vbLf & " On POM.Po_No=POd.Po_no " & _
            vbLf & " Inner Join Trade_Master TM on POM.Supplier_Code=TM.Trade_Code " & _
            vbLf & " inner Join Item_Master IM on POD.Item_Code=IM.Item_Code " & _
            vbLf & " inner Join Unit_Cls U on POD.Unit_Cls=U.Unit_Cls " & _
            vbLf & " inner Join PORequest_Detail PRD on POD.PORequest_No=PRD.PoRequest_No and POD.PoReq_SeqNo=PRD.PoReq_SeqNo " & _
            vbLf & " inner Join PoREquest_Master PRM on POD.PORequest_No=PRM.PoRequest_No " & _
            vbLf & " inner join PersonInCharge_Cls PIC on PRM.PersonInCharge_Cls=PIC.PersonInCharge_Cls " & _
            vbLf & " inner join curr_cls C on POD.Currency_Code=C.Curr_Cls " & _
            vbLf & " where pom.po_no = '" & Trim(txtPoNo.Text) & "' and pom.others_cls = '0' and pom.period is null " & _
            vbLf & " order by pod.PORequest_No, pod.Item_Code, pod.POReq_SeqNo "
' --------
        
        If rsRpt.State <> adStateClosed Then rsRpt.Close
        rsRpt.Open SqlRpt, Db, adOpenDynamic, adLockOptimistic
        
        sqlprint = SqlRpt
        reportcode = "poparts"
        Fbulan = txtPoNo.Text
        printorient = 1
        
        If rsRpt.EOF Then LblErrMsg.Caption = DisplayMsg(4006): Me.MousePointer = vbDefault: Exit Sub
        Set report = application.OpenReport(App.path & "\Reports\rptPONew.rpt")
        report.Database.Tables(1).SetDataSource rsRpt
        
        Rpt.CRViewer1.ReportSource = report
        Rpt.CRViewer1.ViewReport
        Rpt.CRViewer1.Zoom 1
        Rpt.WindowState = 2
        Rpt.Show 1
        
        Set rsRpt = Nothing
        Me.MousePointer = vbDefault
        End If
    Exit Sub
ErrMsg:
    LblErrMsg = err.number & "-" & err.Description
    MousePointer = vbDefault
End Sub

Private Sub CmdSubMenu_Click()
    ClearData
    Unload Me
    frmMainMenu.Show
End Sub

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
    Set RS = Nothing
    Set rsGrid = Nothing
End Sub

Private Sub updateMaster(Flag As Boolean)
    Dim sQl_Master As String
    Dim rs_Master As New ADODB.Recordset
        sQl_Master = "select * from PurchaseOrder_Master where PO_No = '" & Trim$(txtPoNo.Text) & "' and others_cls = '0' and period is null"
        If rs_Master.State <> adStateClosed Then rs_Master.Close
        rs_Master.Open sQl_Master, Db, adOpenKeyset, adLockOptimistic
        If rs_Master.BOF And rs_Master.EOF Then
            LblErrMsg.Caption = DisplayMsg(4015)    'Record with This PO No not found
            txtPoNo.SetFocus
            rs_Master.Close
            Set rs_Master = Nothing
            Exit Sub
        End If
        rs_Master("po_date") = Format(PODate.Value, "yyyy-mm-dd")
        rs_Master("revise_No") = Trim(txtRev.Text)
        rs_Master("supplier_code") = Trim(cboCust.Text)
        If Flag = True Then
            rs_Master("discount") = CDbl(TxtDisc.Text)
            rs_Master("amount") = CDbl(txtamount.Text)
            rs_Master("ppn") = CDbl(txtPPN.Text)
            rs_Master("total_amount") = CDbl(txtGrandTotal.Text)
            rs_Master("remarks") = Trim(txtremarks.Text)
            
            
        End If
        rs_Master.update
        rs_Master.Close
        Set rs_Master = Nothing
End Sub


Function ValidDataSupplier(pPONO As String) As Boolean
Dim ls_sql  As String
Dim rsCek As New Recordset, rsCek2 As New Recordset, lint_recordcount As Integer

ValidDataSupplier = True
ls_sql = "Select distinct item_code, poreq_seqno  from purchaseOrder_detail where po_no ='" & pPONO & "' order by item_code"
If rsCek.State <> adStateClosed Then rsCek.Close
rsCek.CursorLocation = adUseClient
rsCek.Open ls_sql, Db, adOpenKeyset, adLockOptimistic
If rsCek.EOF = False Then


'    lint_recordcount = rsCek.RecordCount
'    ls_sql = " select * from (select distinct item_Code from  price_master where price_Cls = '01' and trade_code = '" & CboCust.Text & "' " & _
'             "   and exists " & _
'             "       ( " & _
'             "       (select distinct a.item_code, reqdelivery_date " & _
'             "       from purchaseOrder_detail a, PORequest_Detail b " & _
'             "       Where a.porequest_no = b.porequest_no " & _
'             "       and a.poreq_seqno = b.seq_no and a.item_code = price_master.item_Code " & _
'             "       and (convert(char(8), Reqdelivery_date,112) between start_date and end_date)  " & _
'             "       and a.PO_no ='" & pPONO & "'))" & _
'             "  union " & _
'             "   select b.item_Code from item_master a,porequest_detail b " & _
'             "   where supplier_code = '" & CboCust.Text & "' and a.item_code = b.item_code " & _
'             "   and POrequest_no = '" & cboRequestNo & "') a order by item_code "
    Do While Not rsCek.EOF
        ls_sql = "select distinct item_code from price_master where trade_code in ('" & cboCust.Text & "','000000') and item_code= '" & rsCek!Item_Code & "' " & _
                    " and convert(char(8), (select reqdelivery_date from porequest_detail where PoReq_seqno='" & rsCek!POReq_seqno & "' ),112) between start_date and end_date " & _
                    " Union " & _
                    " select item_code from item_master where supplier_code = '" & cboCust.Text & "' and item_code ='" & rsCek!Item_Code & "' "
    
        If rsCek2.State <> adStateClosed Then rsCek2.Close
        rsCek2.CursorLocation = adUseClient
        rsCek2.Open ls_sql, Db, adOpenKeyset, adLockOptimistic
        If rsCek2.EOF Then
            ValidDataSupplier = False
            Set rsCek2 = Nothing
            Set rsCek = Nothing
            Exit Function
        End If
    rsCek.MoveNext
    Loop
    Set rsCek2 = Nothing
End If
rsCek.Close
Set rsCek = Nothing
End Function
