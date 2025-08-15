VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDOCreate 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Delivery Note Create"
   ClientHeight    =   10785
   ClientLeft      =   120
   ClientTop       =   345
   ClientWidth     =   17265
   Icon            =   "frmDOCreate.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10785
   ScaleWidth      =   17265
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtRegisterNo 
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
      Left            =   14880
      MaxLength       =   25
      TabIndex        =   65
      Tag             =   "TTFF*/"
      Top             =   9090
      Width           =   1725
   End
   Begin MSComCtl2.DTPicker DTPbcdate 
      Height          =   315
      Left            =   4380
      TabIndex        =   12
      Top             =   9090
      Width           =   1545
      _ExtentX        =   2725
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
      Format          =   67174401
      CurrentDate     =   41092
   End
   Begin VB.TextBox txtbcno 
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
      Left            =   2340
      MaxLength       =   30
      TabIndex        =   11
      Top             =   9090
      Width           =   1965
   End
   Begin VB.TextBox TxtRemarks 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#,###"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
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
      Left            =   9960
      MaxLength       =   35
      TabIndex        =   14
      Top             =   9090
      Width           =   3045
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FDDFE3&
      Height          =   645
      Left            =   480
      TabIndex        =   46
      Top             =   1830
      Width           =   15195
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
         ItemData        =   "frmDOCreate.frx":0E42
         Left            =   150
         List            =   "frmDOCreate.frx":0E4C
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   195
         Width           =   1335
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H0080FFFF&
         Caption         =   "Update"
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
         Left            =   12720
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   150
         Width           =   1140
      End
      Begin MSComCtl2.DTPicker dtDO 
         Height          =   330
         Left            =   5100
         TabIndex        =   7
         Top             =   180
         Width           =   1530
         _ExtentX        =   2699
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
         Format          =   67174403
         CurrentDate     =   37799
      End
      Begin MSForms.ComboBox cboDeliveryCls 
         Height          =   330
         Left            =   8160
         TabIndex        =   64
         Top             =   180
         Width           =   900
         VariousPropertyBits=   612386843
         MaxLength       =   15
         DisplayStyle    =   3
         Size            =   "1587;582"
         ListRows        =   15
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delivery Cls"
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
         Left            =   6960
         TabIndex        =   63
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DN Date"
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
         Left            =   4230
         TabIndex        =   49
         Top             =   255
         Width           =   720
      End
      Begin MSForms.ComboBox cbo 
         Height          =   330
         Index           =   1
         Left            =   2400
         TabIndex        =   6
         Top             =   180
         Width           =   1695
         VariousPropertyBits=   746604571
         MaxLength       =   25
         DisplayStyle    =   3
         Size            =   "2990;582"
         ColumnCount     =   2
         ListRows        =   7
         ShowDropButtonWhen=   2
         Value           =   "YY9999/YYYY"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DN No"
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
         Left            =   1710
         TabIndex        =   48
         Top             =   255
         Width           =   540
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SI/PO. No"
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
         Left            =   9480
         TabIndex        =   47
         Top             =   240
         Width           =   1035
      End
      Begin MSForms.ComboBox cbo 
         Height          =   330
         Index           =   2
         Left            =   10560
         TabIndex        =   8
         Top             =   180
         Width           =   2085
         VariousPropertyBits=   746604571
         MaxLength       =   25
         DisplayStyle    =   3
         Size            =   "3678;582"
         ColumnCount     =   2
         ListRows        =   7
         ShowDropButtonWhen=   2
         Value           =   "001/AAAAAA/07/2003"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.CommandButton cmdReport 
      BackColor       =   &H0080FFFF&
      Caption         =   "D. Instruction"
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
      Height          =   375
      Index           =   1
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   10260
      Width           =   1380
   End
   Begin VB.TextBox txtDO 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#,###"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
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
      Left            =   13050
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   9090
      Width           =   1695
   End
   Begin VB.TextBox txtDoNO 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#,###"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
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
      Left            =   5970
      Locked          =   -1  'True
      TabIndex        =   36
      Text            =   "YY9999/YYYY"
      Top             =   9090
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
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
      Index           =   1
      Left            =   14400
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   10260
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
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
      Left            =   15600
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   10260
      Width           =   1140
   End
   Begin VB.CommandButton cmdPage 
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
      Index           =   0
      Left            =   2685
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   10260
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton cmdPage 
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
      Index           =   1
      Left            =   3945
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   10260
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton cmdPage 
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
      Index           =   2
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   10260
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton cmdPage 
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
      Index           =   3
      Left            =   6405
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   10260
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton cmdReport 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Print"
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
      Height          =   375
      Index           =   0
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   10260
      Width           =   1380
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   450
      TabIndex        =   28
      Top             =   9600
      Width           =   16365
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
         TabIndex        =   29
         Top             =   210
         Width           =   16125
      End
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
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   10260
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1095
      Left            =   480
      TabIndex        =   26
      Top             =   720
      Width           =   16365
      Begin VB.TextBox TxtForwarder 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#,###"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   9660
         MaxLength       =   10
         TabIndex        =   3
         Top             =   300
         Width           =   1035
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#,###"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
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
         Left            =   11040
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   240
         Width           =   3075
      End
      Begin MSComCtl2.DTPicker dtAwal 
         Height          =   330
         Left            =   1470
         TabIndex        =   1
         Top             =   645
         Width           =   1500
         _ExtentX        =   2646
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
         Format          =   67174403
         CurrentDate     =   37799
      End
      Begin MSComCtl2.DTPicker dtAkhir 
         Height          =   330
         Left            =   3360
         TabIndex        =   2
         Top             =   645
         Width           =   1500
         _ExtentX        =   2646
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
         Format          =   67174403
         CurrentDate     =   37860
      End
      Begin VB.Line Line3 
         X1              =   11040
         X2              =   14145
         Y1              =   555
         Y2              =   555
      End
      Begin VB.Label Label 
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
         Index           =   8
         Left            =   8580
         TabIndex        =   56
         Top             =   315
         Width           =   870
      End
      Begin MSForms.ComboBox cbo 
         Height          =   330
         Index           =   3
         Left            =   9600
         TabIndex        =   25
         Top             =   240
         Width           =   1380
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2434;582"
         ColumnCount     =   2
         ListRows        =   15
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         Value           =   "AAAAAAAAAA"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboWH 
         Height          =   330
         Left            =   9600
         TabIndex        =   4
         Top             =   630
         Width           =   1380
         VariousPropertyBits=   612386843
         MaxLength       =   15
         DisplayStyle    =   3
         Size            =   "2434;582"
         ListRows        =   15
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
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
         Left            =   8580
         TabIndex        =   45
         Top             =   705
         Width           =   960
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
         Left            =   11055
         TabIndex        =   44
         Top             =   705
         Width           =   3105
      End
      Begin VB.Line Line2 
         X1              =   11040
         X2              =   14160
         Y1              =   945
         Y2              =   945
      End
      Begin VB.Label Label 
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
         Left            =   3075
         TabIndex        =   40
         Top             =   720
         Width           =   165
      End
      Begin VB.Label Label 
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
         Left            =   120
         TabIndex        =   39
         Top             =   720
         Width           =   1185
      End
      Begin VB.Label lblNm 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
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
         Left            =   3075
         TabIndex        =   35
         Top             =   285
         Width           =   5355
      End
      Begin MSForms.ComboBox cbo 
         Height          =   330
         Index           =   0
         Left            =   1470
         TabIndex        =   0
         Top             =   225
         Width           =   1500
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2646;582"
         ColumnCount     =   2
         ListRows        =   15
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         Value           =   "AAAAAAAAAA"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label 
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
         Left            =   120
         TabIndex        =   34
         Top             =   285
         Width           =   1170
      End
      Begin VB.Line Line1 
         X1              =   3075
         X2              =   8400
         Y1              =   525
         Y2              =   525
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid gridAtas 
      Height          =   2475
      Left            =   480
      TabIndex        =   41
      Top             =   2910
      Width           =   16365
      _cx             =   28866
      _cy             =   4366
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
      BackColorBkg    =   16637923
      BackColorAlternate=   -2147483624
      GridColor       =   12582912
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
   Begin VSFlex8Ctl.VSFlexGrid gridBawah 
      Height          =   2745
      Left            =   480
      TabIndex        =   42
      Top             =   5820
      Width           =   16365
      _cx             =   28866
      _cy             =   4842
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
      BackColorBkg    =   16637923
      BackColorAlternate=   -2147483624
      GridColor       =   12582912
      GridColorFixed  =   12582912
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   1
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
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   420
      Left            =   15000
      TabIndex        =   43
      Top             =   60
      Width           =   1860
      _extentx        =   3281
      _extenty        =   741
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No. Register"
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
      Left            =   14880
      TabIndex        =   66
      Top             =   8760
      Width           =   1050
   End
   Begin MSForms.ComboBox cbobctype 
      Height          =   315
      Left            =   540
      TabIndex        =   10
      Top             =   9090
      Width           =   1725
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "3043;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BC Date"
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
      Left            =   4470
      TabIndex        =   62
      Top             =   8760
      Width           =   720
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BC No."
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
      Left            =   2370
      TabIndex        =   61
      Top             =   8760
      Width           =   600
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BC Type"
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
      Left            =   570
      TabIndex        =   60
      Top             =   8760
      Width           =   735
   End
   Begin VB.Line Line4 
      X1              =   8550
      X2              =   9900
      Y1              =   9330
      Y2              =   9330
   End
   Begin VB.Label LblRemarks 
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
      Left            =   8610
      TabIndex        =   59
      Top             =   9090
      Width           =   1305
   End
   Begin MSForms.ComboBox CboRemarks 
      Height          =   315
      Left            =   7740
      TabIndex        =   13
      Top             =   9090
      Width           =   765
      VariousPropertyBits=   612386843
      MaxLength       =   15
      DisplayStyle    =   3
      Size            =   "1349;556"
      ListRows        =   15
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Additional Remarks"
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
      Left            =   9990
      TabIndex        =   58
      Top             =   8760
      Width           =   1665
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks Cls"
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
      Left            =   7710
      TabIndex        =   57
      Top             =   8760
      Width           =   1095
   End
   Begin VB.Label l6 
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   0
      TabIndex        =   55
      Top             =   0
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label l5 
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   0
      TabIndex        =   54
      Top             =   0
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label l4 
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   -30
      TabIndex        =   53
      Top             =   240
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label l3 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3480
      TabIndex        =   52
      Top             =   60
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label l2 
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   1500
      TabIndex        =   51
      Top             =   390
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Label l1 
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   0
      TabIndex        =   50
      Top             =   0
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label lblfix 
      Alignment       =   2  'Center
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
      Height          =   255
      Left            =   15600
      TabIndex        =   38
      Top             =   2040
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DN No"
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
      Left            =   5970
      TabIndex        =   33
      Top             =   8760
      Width           =   540
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00A6D2FF&
      Height          =   525
      Index           =   0
      Left            =   480
      Top             =   9000
      Width           =   16365
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total DN Amount"
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
      Left            =   13080
      TabIndex        =   32
      Top             =   8760
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FDDFE3&
      BackStyle       =   0  'Transparent
      Caption         =   "Order Entry Data"
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
      Height          =   195
      Index           =   1
      Left            =   480
      TabIndex        =   30
      Top             =   2595
      Width           =   16365
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery Note Create"
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
      Height          =   390
      Index           =   0
      Left            =   510
      TabIndex        =   27
      Top             =   210
      Width           =   16365
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   0
      Left            =   480
      Top             =   2580
      Width           =   16365
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FDDFE3&
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery Note"
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
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   31
      Top             =   5505
      Width           =   16365
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   1
      Left            =   480
      Top             =   5490
      Width           =   16365
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00A6D2FF&
      BackStyle       =   1  'Opaque
      Height          =   315
      Index           =   0
      Left            =   480
      Top             =   8700
      Width           =   16365
   End
End
Attribute VB_Name = "frmDOCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim i As Integer, sql As String
Dim nilKosong As Boolean
Dim gantiDealer As Boolean, gantiDO As Boolean
Dim gantiPO As Boolean, gantiDtAwal As Boolean, gantiDtAkhir As Boolean
Dim prosesSimpan As Boolean, ubahAtas As Boolean
Dim FixCls As String
Dim dblTempValue As Double
Dim defaultCurr As String, defaultLocation As String
Dim li_Row As Integer

Dim bteColSelect As Byte
Dim bteColProdCode As Byte
Dim bteColMakerItem As Byte
Dim bteColDesc As Byte
Dim bteColLotNo As Byte
Dim bteColPONo As Byte
Dim bteColQty As Byte
Dim bteColUnitCls As Byte
Dim bteColUnit As Byte
Dim bteColCtn As Byte
Dim bteColNet As Byte
Dim bteColGross As Byte
Dim bteColDate As Byte
Dim bteColTime As Byte
Dim bteColCurrCls As Byte
Dim bteColCurr As Byte
Dim bteColPrice As Byte
Dim bteColAmount As Byte
Dim bteColQtySisaAwal As Byte
Dim bteColFixOrder As Byte
Dim bteColUpdate As Byte
Dim bteColSeqNo As Byte
Dim bteColDOSeqNo As Byte
Dim bteColQtySisa As Byte
Dim bteColNetWeight As Byte
Dim bteColGrossWeight As Byte
Dim bteColQtyPerCtn As Byte
Dim bteColService As Byte
Dim bteColSerialNoFrom As Byte
Dim bteColSerialNoTo As Byte
Dim bteTempSerial As Byte
Dim bteCek As Byte
Dim bteColBctype As Byte
Dim bteColBcNo As Byte
Dim bteColBCDate As Byte

Dim bteHakPrice As Byte

Private Sub headerGrid(nmGrid, Optional gridAtas As Byte)
    Dim i As Integer
    
    bteColSelect = 0
    bteColProdCode = 1
    bteColMakerItem = 2
    bteColDesc = 3
    bteColLotNo = 4
    bteColPONo = 5
    bteColQty = 6
    bteColUnitCls = 7
    bteColUnit = 8
    bteColSerialNoFrom = 9
    bteColSerialNoTo = 10
    bteColDate = 11
    bteColTime = 12
    bteColCurrCls = 13
    bteColCurr = 14
    bteColPrice = 15
    bteColService = 16
    bteColAmount = 17
    bteColCtn = 18
    bteColNet = 19
    bteColGross = 20
    bteColQtySisaAwal = 21
    bteColFixOrder = 22
    bteColUpdate = 23
    bteColSeqNo = 24
    bteColDOSeqNo = 25
    bteColQtySisa = 26
    bteColNetWeight = 27
    bteColGrossWeight = 28
    bteColQtyPerCtn = 29
    bteTempSerial = 30
    bteCek = 31
    bteColBctype = 32
    bteColBcNo = 33
    bteColBCDate = 34
    
    With nmGrid
        .clear
        .ColS = 35
        .Rows = 1
        
        .TextMatrix(0, bteColSelect) = "S"
        .TextMatrix(0, bteColProdCode) = "Product Code"
        .TextMatrix(0, bteColMakerItem) = "Part Number"
        .TextMatrix(0, bteColDesc) = "Description"
        .TextMatrix(0, bteColLotNo) = "Lot No"
        .TextMatrix(0, bteColPONo) = "SI/PO No."
        .TextMatrix(0, bteColQty) = "Qty"
        .TextMatrix(0, bteColUnitCls) = "UnitCls"
        .TextMatrix(0, bteColUnit) = "Unit"
        .TextMatrix(0, bteColSerialNoFrom) = "Serial No From"
        .TextMatrix(0, bteColSerialNoTo) = "Serial No To"
        .TextMatrix(0, bteColCtn) = "Ctn. Qty"
        .TextMatrix(0, bteColNet) = "Net W."
        .TextMatrix(0, bteColGross) = "Gross W."
        .TextMatrix(0, bteColDate) = "Delivery Date"
        .TextMatrix(0, bteColTime) = "Time"
        .TextMatrix(0, bteColCurrCls) = "CurrCls"
        .TextMatrix(0, bteColCurr) = "Curr"
        .TextMatrix(0, bteColPrice) = "Price"
        .TextMatrix(0, bteColService) = "Service"
        .TextMatrix(0, bteColAmount) = "Amount"
        .TextMatrix(0, bteColQtySisaAwal) = "Grid Atas = Tampung Qty Sisa, Grid Bawah = Tampung Qty Awal"
        .TextMatrix(0, bteColFixOrder) = "Fix Order Entry"
        .TextMatrix(0, bteColUpdate) = "Update"
        .TextMatrix(0, bteColSeqNo) = "SeqNo"
        .TextMatrix(0, bteColDOSeqNo) = "DOSeqNo"
        .TextMatrix(0, bteColQtySisa) = "SisaQty"
        .TextMatrix(0, bteColNetWeight) = "Std Net W."
        .TextMatrix(0, bteColGrossWeight) = "Std Gross W."
        .TextMatrix(0, bteColQtyPerCtn) = "Qty Per Ctn"
        .TextMatrix(0, bteColBctype) = "BC Type"
        .TextMatrix(0, bteColBcNo) = "BC No"
        .TextMatrix(0, bteColBCDate) = "BC Date"
        
        '.ColHidden(BteColSelect) = True
        .ColHidden(bteColCtn) = True
        .ColHidden(bteColNet) = True
        .ColHidden(bteColGross) = True
        .ColHidden(bteCek) = True
                       
        .ColWidth(bteColSelect) = 300
        .ColWidth(bteColProdCode) = 1250
        .ColWidth(bteColMakerItem) = 1250
        
        If gridAtas = 1 Then .ColWidth(bteColDesc) = 2700 Else .ColWidth(bteColDesc) = 2100
        
        .ColWidth(bteColLotNo) = 650
        .ColWidth(bteColPONo) = 1450
        .ColWidth(bteColQty) = 940
        .ColWidth(bteColUnit) = 500
        .ColWidth(bteColSerialNoFrom) = 1500
        .ColWidth(bteColSerialNoTo) = 1500
        .ColWidth(bteColCtn) = 800
        .ColWidth(bteColNet) = 950
        .ColWidth(bteColGross) = 950
        .ColWidth(bteColDate) = 1300
        .ColWidth(bteColTime) = 600
        .ColWidth(bteColCurr) = 500
        .ColWidth(bteColPrice) = 1350
        .ColWidth(bteColService) = 1350
        .ColWidth(bteColAmount) = 1500
        
        .ColAlignment(bteColSelect) = flexAlignCenterCenter
        .ColAlignment(bteColProdCode) = flexAlignLeftCenter
        .ColAlignment(bteColMakerItem) = flexAlignLeftCenter
        .ColAlignment(bteColDesc) = flexAlignLeftCenter
        .ColAlignment(bteColLotNo) = flexAlignLeftCenter
        .ColAlignment(bteColPONo) = flexAlignLeftCenter
        .ColAlignment(bteColQty) = flexAlignRightCenter
        .ColAlignment(bteColCtn) = flexAlignRightCenter
        .ColAlignment(bteColNet) = flexAlignRightCenter
        .ColAlignment(bteColGross) = flexAlignRightCenter
        .ColAlignment(bteColUnit) = flexAlignCenterCenter
        .ColAlignment(bteColTime) = flexAlignCenterCenter
        .ColAlignment(bteColCurr) = flexAlignCenterCenter
        .ColAlignment(bteColPrice) = flexAlignRightCenter
        .ColAlignment(bteColService) = flexAlignRightCenter
        .ColAlignment(bteColAmount) = flexAlignRightCenter
        
        If gridAtas = 1 Then
        .ColHidden(bteColLotNo) = True
'        .ColHidden(BteColSelect) = True
        End If
      
        .ColHidden(bteColUnitCls) = True
        .ColHidden(bteColCurrCls) = True
        .ColHidden(bteColQtySisaAwal) = True
        .ColHidden(bteColFixOrder) = True
        .ColHidden(bteColUpdate) = True
        .ColHidden(bteColSeqNo) = True
        .ColHidden(bteColDOSeqNo) = True
        .ColHidden(bteColQtySisa) = True
        .ColHidden(bteColNetWeight) = True
        .ColHidden(bteColGrossWeight) = True
        .ColHidden(bteColQtyPerCtn) = True
        .ColHidden(bteTempSerial) = True
        .ColHidden(bteColBctype) = True
        .ColHidden(bteColBcNo) = True
        .ColHidden(bteColBCDate) = True
        
        .ColHidden(bteColCurr) = (bteHakPrice = 0)
        .ColHidden(bteColPrice) = (bteHakPrice = 0)
        .ColHidden(bteColService) = (bteHakPrice = 0)
        .ColHidden(bteColAmount) = (bteHakPrice = 0)
        
        .EditMaxLength = 1
    End With
End Sub
Private Sub comboBCtype()
Dim ls_sql As String
Dim rs_combo As New ADODB.Recordset
Dim i As Integer


cboBCType.columnCount = 1
cboBCType.clear

ls_sql = "select bc_type from BC_master"
rs_combo.Open ls_sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
i = 0

Do While Not rs_combo.EOF
cboBCType.AddItem rs_combo("Bc_type")
rs_combo.MoveNext
Loop

cboBCType.ColumnWidths = "90"
cboBCType.ListWidth = 90
cboBCType.ListRows = 7


End Sub
Sub kosongAtas(Optional semua As Integer)
nilKosong = True
    If semua = 0 Then cbo(0) = "": lblNm(0) = "": cbo(2) = ""
    Call headerGrid(gridAtas, 1) 'Kosong Grid Atas
nilKosong = False
End Sub

Sub kosongBwh(Optional semua As Integer)
nilKosong = True
    If semua = 0 Then cbo(1) = "": cbo(1).Enabled = True
    txtDO = 0
    txtDoNO = ""
    txtRegisterNo.Text = ""
    Call headerGrid(gridBawah) 'Kosong Grid Atas
nilKosong = False
End Sub

'******** Combo **********
Sub isiCboCust() 'Isi Combo Dealer CD dr Customer Master
Dim rscbo As New ADODB.Recordset
    
With cbo(0)
    .clear
    .columnCount = 2
    .TextColumn = 1

'    Sql = "select Trade_Code,Trade_Name from Trade_Master " & _
        "where (Trade_Cls = 2  or Trade_Cls = 3) and country_cls ='0' order by Trade_Code"
        sql = "select Trade_Code,Trade_Name from Trade_Master " & _
        "Where Trade_Cls IN ('2','3')  order by Trade_Code"
    Set rscbo = Db.Execute(sql)
            
    i = 0
    Do While Not (rscbo.EOF)
        .AddItem ""
        .List(i, 0) = Trim(rscbo(0))
        .List(i, 1) = Trim(rscbo(1))
        i = i + 1
        rscbo.MoveNext
    Loop
    
    .Text = ""
    .ListWidth = 350
    .ColumnWidths = "50 pt;300 pt"
    Set rscbo = Nothing
End With

With cbo(3)
    .clear
    .columnCount = 2
    .TextColumn = 1

'    Sql = "select Trade_Code,Trade_Name from Trade_Master " & _
        "where (Trade_Cls = 2  or Trade_Cls = 3) and country_cls ='0' order by Trade_Code"
        sql = "select Trade_Code,Trade_Name from Trade_Master " & _
        "where (Trade_Cls = '5')  order by Trade_Code"
    Set rscbo = Db.Execute(sql)
            
    i = 0
    Do While Not (rscbo.EOF)
        .AddItem ""
        .List(i, 0) = Trim(rscbo(0))
        .List(i, 1) = Trim(rscbo(1))
        i = i + 1
        rscbo.MoveNext
    Loop
    
    .Text = ""
    .ListWidth = 250
    .ColumnWidths = "50 pt;200 pt"
    Set rscbo = Nothing
End With

End Sub

Sub isiCboDO() 'Isi Combo DO dr Giro Master
Dim rscbo As New ADODB.Recordset
    
With cbo(1)
    .clear
    .columnCount = 1
    .TextColumn = 1
    
    '**** Utk DO Master
    sql = "select DO_NO from Do_master " & _
        "where cust_code = '" & cbo(0) & _
        "' and Do_Date >='" & Format(dtAwal, "yyyy-MM-dd") & _
        "' and Do_Date <='" & Format(dtAkhir, "yyyy-MM-dd") & _
        "' order by right(rtrim(do_no),4) + subString(rtrim(do_no),12,2) + left(rtrim(do_no),3) desc"
    Set rscbo = Db.Execute(sql)

    i = 0
    Do While Not (rscbo.EOF)
        .AddItem ""
        .List(i, 0) = Trim(rscbo("Do_No"))
        i = i + 1
        rscbo.MoveNext
    Loop
    
    .Text = ""
    .ListWidth = 150
    .ColumnWidths = "150 pt"
    Set rscbo = Nothing
End With
End Sub

Sub isiCboPO() 'Isi Combo DO dr Giro Master
Dim rscbo As New ADODB.Recordset

With cbo(2)
    .clear
    
    .columnCount = 1
    .TextColumn = 1
    
    '.AddItem strAll
    '**** Utk Order Entry Master '*** Ambil NoPO  ** untuk Kawai 1 Dn untuk 1 PO
    sql = "select distinct a.PO_No from OrderEntry_Master a, OrderEntry_Detail b where " & _
        "a.Cust_Code = b.Cust_Code and a.PO_NO = b.PO_NO " & _
        "and a.cust_code = '" & cbo(0) & _
        "' and delivery_Date >='" & Format(dtAwal, "yyyy-MM-dd") & _
        "' and delivery_Date <= '" & Format(dtAkhir, "yyyy-MM-dd") & _
        "' and (Fix_Cls = 0 or Fix_Cls is null) "
    If cboStatus = "Create" Then sql = sql & " And a.Po_no Not In (Select List_Po From DO_Master) "

    sql = sql & " order by a.PO_NO"
    Set rscbo = Db.Execute(sql)
        
    i = 0
    Do While Not (rscbo.EOF)
        .AddItem ""
        .List(i, 0) = Trim(rscbo("Po_No"))
        i = i + 1
        rscbo.MoveNext
    Loop
    
    .Text = ""
    .ListWidth = 150
    .ColumnWidths = "150 pt"
    Set rscbo = Nothing
End With
End Sub

Private Sub cbobctype_Change()
If cboBCType.MatchFound = False Then
    LblErrMsg.Caption = "Please Input Valid BC Type !"
    cboBCType.SetFocus
End If
End Sub

Private Sub cbobctype_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = Asc("'") Then
    KeyAscii = 0
End If
End Sub


Private Sub CboRemarks_Change()
    If CboRemarks.MatchFound Then
        LblRemarks.Caption = CboRemarks.Column(1)
    Else
        LblRemarks.Caption = ""
    End If

End Sub

Private Sub cboWH_Change()
    If cboWH.MatchFound Then
        lblWH.Caption = cboWH.Column(1)
    Else
        lblWH.Caption = ""
    End If
End Sub

'******************
Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
    
    bteHakPrice = hakPrice(Me.Name)
    dtAwal = Date
    dtAkhir = DateAdd("d", 1, Date)
    dtDO.Value = Date
    Call IsiComboWH
    Call isiCboCust
    Call IsiCboRemarks
    
    Call kosongAtas
    Call kosongBwh
    Call comboBCtype
    cboStatus.ListIndex = 0
    
    up_FillCboDelivery
End Sub

'************** Tampilkan Data ***************
Private Sub cbo_Change(Index As Integer)
Dim rsDataDo As New ADODB.Recordset
    
    If nilKosong Then Exit Sub
    LblErrMsg = ""
    If Index = 0 Then
        lblNm(0) = "": gantiDealer = True
        cbo(0) = cbo(0)
        If cbo(0).MatchFound Then
            Call cbo_Click(0)
        Else
            Call headerGrid(gridAtas, 1)
            Call headerGrid(gridBawah)
             txtRegisterNo.Text = ""
            Call cbo(1).clear
            Call cbo(2).clear
        End If
    ElseIf Index = 1 Then
        txtDoNO = Trim(cbo(1)): gantiDO = True
        cbo(1) = cbo(1)
        If cbo(Index).MatchFound Then
            sql = "select Fix_Cls, Amount, Do_Date, WHCode, Forwarder_Code, List_PO, Remarks_Cls, Remarks, BC_Type, BC40_No, BC40_date, ISNULL(Delivery_Cls,'')Delivery_Cls," & _
                "(select max(delivery_date) from delivery_order where do_no = DO_Master.do_no) delivery_date, " & _
                "(select max(delivery_time) from delivery_order where do_no = DO_Master.do_no) delivery_time , ISNULL(No_Register,'')No_Register " & _
                "from DO_Master where DO_NO = '" & cbo(1) & "'"
            Set rsDataDo = Db.Execute(sql)
            If Not (rsDataDo.EOF) Then
                FixCls = IIf(IsNull(rsDataDo("Fix_Cls")), 0, rsDataDo("Fix_Cls"))
                txtDO = Format(rsDataDo("Amount"), gs_formatAmountIDR)
                dtDO.Value = Format(rsDataDo("DO_Date"), "dd MMM yyyy")
                cboWH = Trim(rsDataDo("WHCode"))
                cbo(3) = Trim(rsDataDo("Forwarder_Code") & "")
                cbo(2) = Trim(rsDataDo("List_PO"))
                CboRemarks = IIf(IsNull(Trim(rsDataDo("Remarks_Cls"))), "", Trim(rsDataDo("Remarks_Cls")))
                txtremarks = IIf(IsNull(Trim(rsDataDo("Remarks"))), "", Trim(rsDataDo("Remarks")))
                cboBCType = IIf(IsNull(Trim(rsDataDo("BC_Type"))), "", Trim(rsDataDo("BC_Type")))
                txtBCNo = IIf(IsNull(Trim(rsDataDo("BC40_No"))), "", Trim(rsDataDo("BC40_No")))
                txtRegisterNo.Text = Trim(rsDataDo("No_Register") & "")
                dtpBCDate.Value = IIf(IsNull(rsDataDo("BC40_date")), Format(Now, "dd MMM yyyy"), Format(rsDataDo("BC40_date"), "dd MMM yyyy"))
                cboDeliveryCls.Text = IIf(IsNull(Trim(rsDataDo("Delivery_Cls"))), "", Trim(rsDataDo("Delivery_Cls")))
            End If
            Set rsDataDo = Nothing
        Else
            FixCls = 0
            txtDO = 0
            txtRegisterNo.Text = ""
        End If
        
        Call headerGrid(gridAtas)
        Call headerGrid(gridBawah)
        Call cmdUpdate_Click
        
        Set rsDataDo = Nothing
        If FixCls = 1 Then lblFix.Visible = True Else lblFix.Visible = False
    ElseIf Index = 2 Then
        gantiPO = True
    ElseIf Index = 3 Then ' Add Combo Forwarder
        Text1 = ""
        TxtForwarder = cbo(3)
        cbo(3) = cbo(3)
        If cbo(3).MatchFound Then
            Call cbo_Click(3)
        End If
            
    End If
End Sub

Private Sub cbo_KeyDown(Index As Integer, KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 And Index = 0 Then Call cbo_Click(Index)
End Sub

Public Sub cbo_Click(Index As Integer)
    If nilKosong Then Exit Sub
    Me.MousePointer = vbHourglass
    
    If Index = 0 Then 'Jika Combo Dealer
        If gantiDealer Then Call tampilCust: gantiDealer = False 'tampil Order Entry (Grid Atas)
    ElseIf Index = 3 Then 'Jika Combo Forwarder
        If cbo(3).MatchFound = True Then
            Text1 = cbo(3).Column(1)
            TxtForwarder = cbo(3)
        Else
            Text1 = ""
            TxtForwarder = ""
        End If
    End If
    
    Me.MousePointer = vbDefault
End Sub

Sub tampilCust(Optional drCboStatus As Integer)
Dim tampung As String

nilKosong = True
Me.MousePointer = vbHourglass

If cbo(0) <> "" Then
    cbo(0) = cbo(0)
    
    If cbo(0).MatchFound = False Then 'Jika Customer Tidak ketemu
        cbo(1).clear
        cbo(2).clear
        cmdreport(0).Enabled = False
        cmdreport(1).Enabled = False
        If drCboStatus = 0 Then
            lblNm(0) = ""
            Call headerGrid(gridAtas, 1)
        End If
        Call kosongBwh
        LblErrMsg = DisplayMsg(4011)
    
    Else 'Jika ketemu
        lblNm(0) = cbo(0).Column(1)
                
        Call kosongAtas(1)
        Call kosongBwh(1)
        If drCboStatus = 0 Then
            Call isiCboPO
            cmdreport(0).Enabled = True
            cmdreport(1).Enabled = True
        End If
        
        If cboStatus = "Create" Then Call buatDoBaru Else Call isiCboDO
        LblErrMsg = ""
    End If
End If
Me.MousePointer = vbDefault
nilKosong = False
End Sub

Sub buatDoBaru()
    Dim rsBuatNo As New ADODB.Recordset
    sql = "Select Isnull(Max(Substring(DO_No, 3, 4)), 0)  + 1 Nomor " & _
        "From DO_Master Where Year(DO_Date) = " & dtDO.Year
        
    Set rsBuatNo = Db.Execute(sql)
    cbo(1) = Format(dtDO, "YY") & Format(rsBuatNo(0), "0000") & "/" & Format(dtDO, "YYYY")
    cbo(1).Enabled = False
    txtDoNO = Trim(cbo(1))
    Set rsBuatNo = Nothing
    
    uf_GetNoRegister
End Sub

Private Sub cboStatus_Click()
    If nilKosong = True Then Exit Sub
    If cboStatus = "Create" Then HapusDOMaster
    Call tampilCust(1)
    Call isiCboPO
    If cboStatus = "Create" Then
        cbo(1).Enabled = False
        cmdUpdate.Caption = "Create"
        TxtForwarder = ""
        cboWH = ""
        cbo(2) = ""
        
        uf_GetNoRegister
    Else
        cbo(1).Enabled = True
        cmdUpdate.Caption = "Update"
     End If
End Sub

Private Sub dtDO_Change()
    If cboStatus = "Create" Then
        Call tampilCust(1)
        uf_GetNoRegister
    End If
End Sub

Private Sub dtAwal_Change()
    LblErrMsg = ""
    gantiDtAwal = True
    If dtAwal > dtAkhir Then LblErrMsg = DisplayMsg("4076") & " " & Format(dtAkhir, "dd MMM yyyy"): cbo(1).clear: cbo(2).clear: Exit Sub
    Call filterCboPO
    If cboStatus = "Update" Then Call filterCboDO
End Sub

Private Sub dtAkhir_Change()
    LblErrMsg = ""
    gantiDtAkhir = True
    If dtAkhir < dtAwal Then LblErrMsg = DisplayMsg("4077") & " " & Format(dtAwal, "dd MMM yyyy"): cbo(1).clear: cbo(2).clear: Exit Sub
    Call filterCboPO
    If cboStatus = "Update" Then Call filterCboDO
End Sub

Sub filterCboPO()
Dim tampungPO As String
    
    nilKosong = True
    tampungPO = cbo(2)
    Call isiCboPO
    cbo(2) = tampungPO
    cbo(2) = cbo(2)
    If Not (cbo(2).MatchFound) Then cbo(2) = ""
    nilKosong = False
End Sub

Sub filterCboDO()
Dim tampungDO As String
    
    nilKosong = True
    tampungDO = cbo(1)
    Call isiCboDO
    cbo(1) = tampungDO
    cbo(1) = cbo(1)
    If Not (cbo(1).MatchFound) Then cbo(1) = ""
    nilKosong = False
End Sub

Private Sub dtAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys vbTab
End Sub

Private Sub dtAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys vbTab
End Sub

Sub isiGridAtas()
Dim RsOe As New ADODB.Recordset

Dim dblQtyCtn As Double
Dim dblNet As Double
Dim dblGross As Double

Call headerGrid(gridAtas, 1)
With gridAtas
    '******** Ambil dr Order Entry utk Grid Atas
    sql = "select b.item_code, d.makeritem_code, item_name,a.po_no,b.serialNoFrom,b.SerialNoTo," & _
            "delivery_Date,Seq_No,Delivery_Time," & _
            "Currency_Code,qty,Fix_Cls,b.MakerItem_Code," & _
            "(select isnull(sum(c.qty),0) from Delivery_Order c " & _
                "where c.PO_No = a.PO_NO " & _
                "And c.Seq_No = b.Seq_No) as qtyDO, price, b.unit_Cls, " & _
            "isnull(e.NetWeight, 0) NetWeight, isnull(e.GrossWeight, 0) GrossWeight, isnull(e.Length, 0) Length, " & _
            "isnull(e.Width, 0) Width, isnull(e.Thickness, 0) Thickness, isnull(e.Number_Entering, 0) Number_Entering, "
    
    sql = sql & _
            "qty - " & _
            "(select isnull(sum(c.qty),0) from Delivery_Order c " & _
                "where c.PO_No = a.PO_NO " & _
                "and c.Seq_No = b.Seq_No) as sisa "
    sql = sql & ",Service " 'ISNULL((SELECT Price FROM Price_master WHERE Price_Cls='05' AND Item_Code=b.ITem_Code AND trade_Code=a.cust_code),0)As Service"
    sql = sql & _
            " from orderentry_master a " & vbCrLf & _
            "   inner join orderentry_Detail b on a.cust_code = b.cust_code And a.po_no = b.po_no " & _
            "   Inner join  item_master d on b.Item_Code = d.Item_Code " & vbCrLf & _
            "   Left join (SElect * From packingitem_master Where PackingStyle_Cls = '01') e  on d.item_code = e.item_code " & _
            "Where Delivery_Date >= '" & Format(dtAwal, "yyyy-MM-dd") & "' " & _
            "And Delivery_Date <= '" & Format(dtAkhir, "yyyy-MM-dd") & "' " & _
            "And (Fix_Cls = 0 or Fix_Cls is null) " & _
            " " & _
            "And a.Cust_Code ='" & Trim(cbo(0)) & "' "

    'If cbo(2).ListIndex <> -1 Then Sql = Sql & "and a.PO_NO ='" & Trim(cbo(2)) & "' "
    
    sql = sql & "and a.PO_NO ='" & Trim(cbo(2)) & "' "
    
    sql = sql & "order by a.Po_no, b.Delivery_Date, Seq_No"
    Set RsOe = Db.Execute(sql)
    
    i = 1
    Do While Not RsOe.EOF
        
        If RsOe("Number_Entering") = 0 Then
            dblQtyCtn = 0
            dblNet = 0
            dblGross = 0
        Else
            dblQtyCtn = uf_Ceiling(RsOe("sisa") / RsOe("Number_Entering"))
            dblNet = dblQtyCtn * RsOe("NetWeight")
            dblGross = dblQtyCtn * RsOe("GrossWeight")
        End If
    
        .Rows = .Rows + 1
        .Cell(flexcpBackColor, i, bteColSelect) = vbWhite
        .Cell(flexcpChecked, i, bteColSelect) = flexUnchecked
        
        .TextMatrix(i, bteColProdCode) = Trim(RsOe("Item_Code"))
        .TextMatrix(i, bteColMakerItem) = Trim(RsOe("MakerItem_Code"))
        .TextMatrix(i, bteColDesc) = Trim(RsOe("Item_Name"))
        .TextMatrix(i, bteColPONo) = Trim(RsOe("PO_NO"))
        .TextMatrix(i, bteColQty) = Format(RsOe("sisa"), gs_formatQty)
        .TextMatrix(i, bteColCtn) = Format(dblQtyCtn, gs_formatBox)
        .TextMatrix(i, bteColNet) = Format(dblNet, gs_formatWeight)
        .TextMatrix(i, bteColGross) = Format(dblGross, gs_formatWeight)
        .TextMatrix(i, bteColUnitCls) = Trim(RsOe("unit_Cls"))
        .TextMatrix(i, bteColUnit) = uf_GetUnitDescription(Trim(RsOe("unit_Cls")))
        ' Add For Kawai
        .TextMatrix(i, bteColSerialNoFrom) = IIf(IsNull(Trim(RsOe("SerialNoFrom"))), "", Trim(RsOe("SerialNoFrom")))
        .TextMatrix(i, bteColSerialNoTo) = IIf(IsNull(Trim(RsOe("SerialNoTo"))), "", Trim(RsOe("SerialNoTo")))
        ' ---
        .TextMatrix(i, bteColDate) = Format(RsOe("Delivery_Date"), "dd MMM yyyy")
        .TextMatrix(i, bteColTime) = Format(RsOe("Delivery_Time"), "HH:MM")
        .TextMatrix(i, bteColCurrCls) = IIf(IsNull(RsOe("Currency_Code")), "", Trim(RsOe("Currency_Code"))) 'Curr
        If IsNull(RsOe("Currency_Code")) Or Trim(RsOe("Currency_Code")) = "" Then
            .TextMatrix(i, bteColCurr) = ""
        Else
            .TextMatrix(i, bteColCurr) = uf_GetCurrencyDescription(Trim(RsOe("Currency_Code")))
        End If
        If Trim(RsOe("Currency_Code")) = "03" Then
            .TextMatrix(i, bteColPrice) = Format(RsOe("Price"), gs_formatPriceIDR)  'Harga
            .TextMatrix(i, bteColService) = Format(RsOe("Service"), gs_formatPriceIDR)   'Harga
        Else
            .TextMatrix(i, bteColPrice) = Format(RsOe("Price"), gs_formatPrice) 'Harga
            .TextMatrix(i, bteColService) = Format(RsOe("Service"), gs_formatPrice)  'Harga
        End If
        .TextMatrix(i, bteColAmount) = (Format(RsOe("Price"), gs_formatAmount) + RsOe("Service")) * RsOe("sisa")
        'Format(uf_Trunc((Format(RsOe("Price"), gs_formatAmountIDR) + RsOe("Service")) * RsOe("sisa"), gi_decimalDigitAmountIDR), gs_formatAmountIDR) 'Harga
        .TextMatrix(i, bteColQtySisaAwal) = CDbl(.TextMatrix(i, bteColQty))  'Tampung Qty Sisa
        .TextMatrix(i, bteColFixOrder) = IIf(IsNull(Trim(RsOe("Fix_Cls"))), 0, RsOe("Fix_Cls")) 'Fix/ Belon
        .TextMatrix(i, bteColSeqNo) = RsOe("Seq_No")
        .TextMatrix(i, bteColNetWeight) = RsOe("NetWeight")
        .TextMatrix(i, bteColGrossWeight) = RsOe("GrossWeight")
        .TextMatrix(i, bteColQtyPerCtn) = RsOe("Number_Entering")
        Call warnaGrid(CLng(i), CDbl(.TextMatrix(i, bteColQty)))
        i = i + 1
        RsOe.MoveNext
    Loop
    Set RsOe = Nothing
End With
End Sub

Sub isiGridBawah()
    
    Dim RsDo As New ADODB.Recordset
    Dim RsDoBawah As New ADODB.Recordset
    Call headerGrid(gridBawah)
    Dim X As Integer
    
    With gridBawah
        '******** Ambil dr DO utk Grid Bawah
        sql = " SELECT A.DO_No,A.Item_Code,A.MakerItem_Code,A.Delivery_Date " & vbCrLf & _
            "     ,A.Delivery_Time,A.PO_No,A.Seq_No,A.DOSeq_No, a.SerialNoFrom, a.SerialNoTo " & vbCrLf & _
            "      ,A.Qty,A.Unit_Cls,A.CtnQty,A.NetWeight " & vbCrLf & _
            "      ,A.GrossWeight,A.Currency_Code,A.Price " & vbCrLf & _
            "      ,iSNULL(A.Service,0)AS Service,A.Amount,A.Revised_Cls " & vbCrLf & _
            "      ,A.Lot_No,A.Last_Update,A.Last_User,A.Register_Date, " & vbCrLf & _
            "isnull((Select Fix_Cls from orderentry_master c " & vbCrLf & _
                "where c.Cust_Code = b.Cust_Code and  c.PO_NO = a.PO_No),0) as fix,d.Item_Name, " & vbCrLf & _
            "isnull(e.GrossWeight, 0) StdGross, isnull(e.NetWeight, 0) StdNet, isnull(e.Number_Entering, 0) StdQtyCtn " & vbCrLf & _
            "From delivery_order a " & vbCrLf & _
            "   Inner join do_master b on A.do_no = B.do_no " & vbCrLf & _
            "   Inner join Item_Master d on a.Item_Code = d.Item_Code " & vbCrLf & _
            "   Left join (Select * From PackingItem_Master Where PackingStyle_Cls = '01' ) e  on d.item_code = e.item_code " & vbCrLf & _
            "Where 'A' = 'A' " & vbCrLf & _
            "--e.PackingStyle_Cls = '01' " & vbCrLf & _
            "And a.DO_NO ='" & cbo(1) & "' " & vbCrLf & _
            "And a.PO_NO ='" & Trim(cbo(2)) & "'" & vbCrLf & _
            " order by a.Po_no, a.Delivery_Date, A.Seq_No"
            
            '"Order by d.Group_Cls, a.MakerItem_Code,Delivery_Date,a.PO_NO,Seq_No, DOSeq_No,Lot_No"
        Set RsDo = Db.Execute(sql)
        
        If Not (RsDo.EOF) Then
            i = 1
            Do While Not RsDo.EOF
                
                .Rows = .Rows + 1
                .TextMatrix(i, bteColSelect) = ""
                .Cell(flexcpBackColor, i, bteColSelect) = vbWhite
                .TextMatrix(i, bteColProdCode) = Trim(RsDo("Item_Code"))
                .TextMatrix(i, bteColMakerItem) = Trim(RsDo("MakerItem_Code"))
                .TextMatrix(i, bteColDesc) = Trim(RsDo("Item_Name"))
                
                .Cell(flexcpBackColor, i, bteColLotNo) = vbWhite
                .TextMatrix(i, bteColLotNo) = IIf(IsNull(RsDo("Lot_No")), "", Trim(RsDo("Lot_No")))
                
                .TextMatrix(i, bteColPONo) = Trim(RsDo("PO_NO"))
                
                .Cell(flexcpBackColor, i, bteColQty) = vbWhite
                .Cell(flexcpBackColor, i, bteColCtn) = vbWhite
                .Cell(flexcpBackColor, i, bteColNet) = vbWhite
                .Cell(flexcpBackColor, i, bteColGross) = vbWhite
                
                .TextMatrix(i, bteColQty) = Format(RsDo("qty"), gs_formatQty)
                .TextMatrix(i, bteColUnitCls) = Trim(RsDo("unit_Cls"))
                .TextMatrix(i, bteColUnit) = uf_GetUnitDescription(Trim(RsDo("unit_Cls")))
                .TextMatrix(i, bteColSerialNoFrom) = IIf(IsNull(Trim(RsDo("SerialNoFrom"))), "", RsDo("SerialNoFrom"))
                .TextMatrix(i, bteColSerialNoTo) = IIf(IsNull(Trim(RsDo("SerialNoTo"))), "", RsDo("SerialNoTo"))
                .TextMatrix(i, bteColCtn) = Format(RsDo("CtnQty"), gs_formatBox)
                .TextMatrix(i, bteColNet) = Format(RsDo("NetWeight"), gs_formatWeight)
                .TextMatrix(i, bteColGross) = Format(RsDo("GrossWeight"), gs_formatWeight)
                .TextMatrix(i, bteColDate) = Format(RsDo("Delivery_Date"), "dd MMM yyyy")
                .TextMatrix(i, bteColTime) = Format(RsDo("Delivery_Time"), "HH:MM")
                .TextMatrix(i, bteColCurrCls) = IIf(IsNull(RsDo("Currency_Code")), "", Trim(RsDo("Currency_Code")))
                
                If IsNull(RsDo("Currency_Code")) Or Trim(RsDo("Currency_Code")) = "" Then
                    .TextMatrix(i, bteColCurr) = ""
                Else
                    .TextMatrix(i, bteColCurr) = uf_GetCurrencyDescription(RsDo("Currency_Code"))
                End If
                
                .Cell(flexcpBackColor, i, bteColPrice) = vbWhite
                .Cell(flexcpBackColor, i, bteColService) = vbWhite
                If RsDo("Currency_Code") = "03" Then
                    .TextMatrix(i, bteColPrice) = Format(RsDo("Price"), gs_formatPriceIDR)  'Harga
                    .TextMatrix(i, bteColService) = Format(RsDo("Service"), gs_formatPriceIDR)   'Harga
                Else
                    .TextMatrix(i, bteColPrice) = Format(RsDo("Price"), gs_formatPrice) 'Harga
                    .TextMatrix(i, bteColService) = Format(RsDo("Service"), gs_formatPrice)  'Harga
                End If
                .TextMatrix(i, bteColAmount) = (Format(RsDo("Price"), gs_formatAmount) + RsDo("Service")) * RsDo("qty")
                'Format(uf_Trunc((Format(RsDo("Price"), gs_formatAmountIDR) + RsDo("Service")) * RsDo("qty"), gi_decimalDigitAmountIDR), gs_formatAmountIDR) 'Harga
                
                '***Hide
                .TextMatrix(i, bteColQtySisaAwal) = CDbl(.TextMatrix(i, bteColQty)) 'tuk menampung nilai qtyAwal
                .TextMatrix(i, bteColFixOrder) = Trim(RsDo("Fix")) 'tuk menentukan Fix Order Entry
                .TextMatrix(i, bteColUpdate) = ""
                .TextMatrix(i, bteColSeqNo) = RsDo("Seq_No")
                
                .TextMatrix(i, bteColDOSeqNo) = RsDo("DOSeq_No")
                .TextMatrix(i, bteColQtySisa) = nilQty(cbo(0), RsDo!Item_Code, RsDo!po_no, Format(RsDo!delivery_Date, "yyyy-MM-dd"), RsDo!Seq_no)
                .TextMatrix(i, bteColNetWeight) = RsDo("StdNet")
                .TextMatrix(i, bteColGrossWeight) = RsDo("StdGross")
                .TextMatrix(i, bteColQtyPerCtn) = RsDo("StdQtyCtn")
                .TextMatrix(i, bteTempSerial) = IIf(IsNull(RsDo("SerialNoFrom")), String(7, " "), Trim(RsDo("SerialNoFrom"))) & IIf(IsNull(RsDo("SerialNoTo")), String(7, " "), Trim(RsDo("SerialNoTo"))) ' Save Temporary SerialNo
    
                defaultCurr = Trim(.TextMatrix(i, bteColCurr))
                
                i = i + 1
                RsDo.MoveNext
            Loop
            If RsDo.State <> adStateClosed Then RsDo.Close
        Else
        'Jika DO Tidak ketemu
        sql = " Select * From ( SELECT A.Do_No,A.Item_Code,A.MakerItem_Code, " & vbCrLf & _
                "     d.Item_Name,A.PO_No,A.Delivery_Date, " & vbCrLf & _
                "     A.Delivery_Time,A.SerialNoFrom,A.SerialNoTo, " & vbCrLf & _
                "     A.Seq_No,A.Currency_Code,A.Qty, " & vbCrLf & _
                "     A.Unit_Cls,A.CtnQty,A.NetWeight, " & vbCrLf & _
                "     A.GrossWeight,A.Price, " & vbCrLf & _
                "     Isnull(A.Service,0) As [Service], " & vbCrLf & _
                "     Isnull((Select Fix_Cls From Orderentry_Master C  " & vbCrLf & _
                "         Where C.Cust_Code = B.Cust_Code And  C.Po_No = a.Po_No),0) As Fix_Cls, " & vbCrLf & _
                "     '' MarkerItem, A.Amount,A.Revised_Cls, " & vbCrLf & _
                "     A.Lot_No,A.Last_Update,A.Last_User, "

        sql = sql + "     A.Register_Date, Isnull(E.GrossWeight, 0) StdGross, " & vbCrLf & _
                "     Isnull(E.NetWeight, 0) StdNet, " & vbCrLf & _
                "     Isnull(E.Number_Entering, 0) StdQtyCtn, " & vbCrLf & _
                "     A.DOSeq_No,0 Length,0 Width,0 Thickness, " & vbCrLf & _
                "     0 Number_Entering,0 Sisa  " & vbCrLf & _
                " From Delivery_Order A inner join Do_Master B on A.Do_No = B.Do_No " & vbCrLf & _
                "   Inner Join Item_Master D on  A.Item_Code = D.Item_Code " & vbCrLf & _
                "   Left Join PackingItem_Master E on D.Item_Code = E.Item_Code " & vbCrLf & _
                " Where E.PackingStyle_Cls = '01'  " & vbCrLf & _
                " And A.Do_No ='" & cbo(1) & "' And A.Po_No ='" & Trim(cbo(2)) & "' " & vbCrLf & _
                "  " & vbCrLf & _
                " Union All "

        sql = sql + "  " & vbCrLf & _
                " Select '' Do_No,B.Item_Code,D.MakerItem_Code, " & vbCrLf & _
                "     Item_Name,A.Po_No,Delivery_Date,Delivery_Time, " & vbCrLf & _
                "     B.SerialNoFrom,B.SerialNoTo,Seq_No,Currency_Code, " & vbCrLf & _
                "     Qty,B.Unit_Cls,(Select Isnull(Sum(C.Qty),0)  " & vbCrLf & _
                "         From Delivery_Order C  " & vbCrLf & _
                "         where C.PO_No = A.Po_No  " & vbCrLf & _
                "         And C.Seq_No = B.Seq_No) As QtyDO, " & vbCrLf & _
                "     Isnull(E.NetWeight, 0) NetWeight, " & vbCrLf & _
                "     Isnull(E.GrossWeight, 0) GrossWeight, " & vbCrLf & _
                "     Price,[Service],isnull(Fix_Cls,0) Fix_Cls,B.MakerItem_Code, "

        sql = sql + "     0 Amount,0 Revised_Cls,Month(Delivery_Date) Lot_No, " & vbCrLf & _
                "     Getdate() Last_Update,'' Last_User,Getdate() Register_Date, " & vbCrLf & _
                "     0 StdGross,0 StdNet,0 StdQtyCtn,0 DoSeq_No, " & vbCrLf & _
                "     Isnull(E.Length, 0) Length, isnull(E.Width, 0) Width, " & vbCrLf & _
                "     Isnull(E.Thickness, 0) Thickness, " & vbCrLf & _
                "     Isnull(E.Number_Entering, 0) Number_Entering, " & vbCrLf & _
                "     Qty - (Select Isnull(Sum(C.Qty),0)  " & vbCrLf & _
                "         From Delivery_Order C  " & vbCrLf & _
                "         Where C.Po_No = A.Po_No And C.Seq_No = B.Seq_No) As Sisa " & vbCrLf & _
                " From Orderentry_Master A " & vbCrLf & _
                "   Inner join Orderentry_Detail B On A.Cust_Code = B.Cust_Code And A.Po_No = B.Po_No " & vbCrLf & _
                "   Inner Join Item_Master D On B.Item_Code = D.Item_Code " & vbCrLf & _
                "   Left Join PackingItem_Master E On D.Item_Code = E.Item_Code  "

        sql = sql + " Where Delivery_Date >= '" & Format(dtAwal, "YYYY-MM-DD") & "' And Delivery_Date <= '" & Format(dtAkhir, "YYYY-MM-DD") & "'  " & vbCrLf & _
                " And (Fix_Cls = 0 or Fix_Cls is null) And E.PackingStyle_Cls = '01' " & vbCrLf & _
                " And A.Cust_Code ='" & Trim(cbo(0)) & "' And a.Po_No ='" & Trim(cbo(2)) & "' " & vbCrLf & _
                " ) DN Order By PO_No, Delivery_Date, Seq_No "

        Set RsDoBawah = Db.Execute(sql)
            If Not (RsDoBawah.EOF) Then
                X = 1
                Do While Not RsDoBawah.EOF

                .Rows = .Rows + 1
                .TextMatrix(X, bteColSelect) = ""
                .Cell(flexcpBackColor, X, bteColSelect) = vbWhite
                .TextMatrix(X, bteColProdCode) = Trim(RsDoBawah("Item_Code"))
                .TextMatrix(X, bteColMakerItem) = Trim(RsDoBawah("MakerItem_Code"))
                .TextMatrix(X, bteColDesc) = Trim(RsDoBawah("Item_Name"))

                .Cell(flexcpBackColor, X, bteColLotNo) = vbWhite
                .TextMatrix(X, bteColLotNo) = IIf(IsNull(RsDoBawah("Lot_No")), "", Trim(RsDoBawah("Lot_No")))

                .TextMatrix(X, bteColPONo) = Trim(RsDoBawah("PO_NO"))

                .Cell(flexcpBackColor, X, bteColQty) = vbWhite
                .Cell(flexcpBackColor, X, bteColCtn) = vbWhite
                .Cell(flexcpBackColor, X, bteColNet) = vbWhite
                .Cell(flexcpBackColor, X, bteColGross) = vbWhite

                .TextMatrix(X, bteColQty) = Format(RsDoBawah("qty"), gs_formatQty)
                .TextMatrix(X, bteColUnitCls) = Trim(RsDoBawah("unit_Cls"))
                .TextMatrix(X, bteColUnit) = uf_GetUnitDescription(Trim(RsDoBawah("unit_Cls")))
                .TextMatrix(X, bteColSerialNoFrom) = IIf(IsNull(Trim(RsDoBawah("SerialNoFrom"))), "", RsDoBawah("SerialNoFrom"))
                .TextMatrix(X, bteColSerialNoTo) = IIf(IsNull(Trim(RsDoBawah("SerialNoTo"))), "", RsDoBawah("SerialNoTo"))
                .TextMatrix(X, bteColCtn) = Format(RsDoBawah("CtnQty"), gs_formatBox)
                .TextMatrix(X, bteColNet) = Format(RsDoBawah("NetWeight"), gs_formatWeight)
                .TextMatrix(X, bteColGross) = Format(RsDoBawah("GrossWeight"), gs_formatWeight)
                .TextMatrix(X, bteColDate) = Format(RsDoBawah("Delivery_Date"), "dd MMM yyyy")
                .TextMatrix(X, bteColTime) = Format(RsDoBawah("Delivery_Time"), "HH:MM")
                .TextMatrix(X, bteColCurrCls) = IIf(IsNull(RsDoBawah("Currency_Code")), "", Trim(RsDoBawah("Currency_Code")))

                If IsNull(RsDoBawah("Currency_Code")) Or Trim(RsDoBawah("Currency_Code")) = "" Then
                    .TextMatrix(X, bteColCurr) = ""
                Else
                    .TextMatrix(X, bteColCurr) = uf_GetCurrencyDescription(RsDoBawah("Currency_Code"))
                End If

                .Cell(flexcpBackColor, X, bteColPrice) = vbWhite
                .Cell(flexcpBackColor, X, bteColService) = vbWhite
                If RsDoBawah("Currency_Code") = "03" Then
                    .TextMatrix(X, bteColPrice) = Format(RsDoBawah("Price"), gs_formatPriceIDR)  'Harga
                    .TextMatrix(X, bteColService) = Format(RsDoBawah("Service"), gs_formatPriceIDR)   'Harga
                Else
                    .TextMatrix(X, bteColPrice) = Format(RsDoBawah("Price"), gs_formatPrice) 'Harga
                    .TextMatrix(X, bteColService) = Format(RsDoBawah("Service"), gs_formatPrice)  'Harga
                End If
                .TextMatrix(X, bteColAmount) = (Format(RsDoBawah("Price"), gs_formatAmount) + RsDoBawah("Service")) * RsDoBawah("qty") 'Harga 'Harga
                'Format(uf_Trunc((Format(RsDoBawah("Price"), gs_formatAmountIDR) + RsDoBawah("Service")) * RsDoBawah("qty"), gi_decimalDigitAmountIDR), gs_formatAmountIDR) 'Harga
'----------------
                .TextMatrix(X, bteColSeqNo) = RsDoBawah("Seq_No")
                .TextMatrix(X, bteColUpdate) = "u"
                .TextMatrix(X, bteColNetWeight) = RsDoBawah("StdNet")
                .TextMatrix(X, bteColGrossWeight) = RsDoBawah("StdGross")
                .TextMatrix(X, bteColQtyPerCtn) = RsDoBawah("StdQtyCtn")
                .TextMatrix(X, bteColDOSeqNo) = X
'-----------
'               ***Hide
                '.TextMatrix(x, bteColQtySisaAwal) = CDbl(.TextMatrix(i, bteColQty)) 'tuk menampung nilai qtyAwal
'                .TextMatrix(x, bteColFixOrder) = Trim(RsDoBawah("Fix")) 'tuk menentukan Fix Order Entry
'                .TextMatrix(x, bteColUpdate) = ""
'                .TextMatrix(x, bteColSeqNo) = RsDoBawah("Seq_No")
'
'                .TextMatrix(x, bteColDOSeqNo) = RsDoBawah("DOSeq_No")
'                .TextMatrix(x, bteColQtySisa) = nilQty(cbo(0), RsDoBawah!Item_Code, RsDo!po_no, Format(RsDoBawah!delivery_Date, "yyyy-MM-dd"), RsDoBawah!Seq_No)
'                .TextMatrix(x, bteColNetWeight) = RsDoBawah("StdNet")
'                .TextMatrix(x, bteColGrossWeight) = RsDoBawah("StdGross")
'                .TextMatrix(x, bteColQtyPerCtn) = RsDoBawah("StdQtyCtn")
'                .TextMatrix(x, bteTempSerial) = IIf(IsNull(RsDoBawah("SerialNoFrom")), String(7, " "), Trim(RsDoBawah("SerialNoFrom"))) & IIf(IsNull(RsDoBawah("SerialNoTo")), String(7, " "), Trim(RsDoBawah("SerialNoTo"))) ' Save Temporary SerialNo
'
                defaultCurr = Trim(.TextMatrix(X, bteColCurr))

                X = X + 1
                RsDoBawah.MoveNext
                Loop
              If RsDoBawah.State <> adStateClosed Then RsDoBawah.Close
            End If
        defaultCurr = ""
        End If
       
        Set RsDoBawah = Nothing
        Set RsDo = Nothing
    End With

End Sub

Private Sub gridAtas_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If nilKosong Then Exit Sub
If FixCls = 1 Then Cancel = True: Exit Sub
With gridAtas
    If Col = bteColSelect Then
        If cbo(1) = "" Then 'Belon isi DO
            Cancel = True
            LblErrMsg = DisplayMsg(1035)
        ElseIf cmdUpdate.Caption = "Create" Then 'Blm klik Create
            Cancel = True
            cmdUpdate.SetFocus
            LblErrMsg = DisplayMsg(1038)
        ElseIf .TextMatrix(Row, bteColFixOrder) = "1" Then 'Fix
            Cancel = True
            LblErrMsg = DisplayMsg(1104)
        ElseIf (Trim(.TextMatrix(Row, bteColCurr)) <> defaultCurr) And defaultCurr <> "" Then
            LblErrMsg = DisplayMsg("0040") & " " & defaultCurr
            Cancel = True
            prosesSimpan = False
'        ElseIf Not CheckPONo(Row) Then
'            Cancel = True
        Else
            Cancel = False
        End If
    Else
        Cancel = True
    End If
End With
End Sub

Private Sub gridBawah_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If FixCls = 1 Then Cancel = True
With gridBawah
    Select Case Col
    Case bteColSelect, bteColLotNo, bteColQty, bteColCtn, bteColNet, bteColGross, bteColPrice, bteColService, bteColSerialNoFrom, bteColSerialNoTo
        If .TextMatrix(Row, bteColFixOrder) = "1" And prosesSimpan = False Then
            LblErrMsg = DisplayMsg(1104)
            Cancel = True
            prosesSimpan = False
        Else
            Select Case Col
            Case bteColSelect: .EditMaxLength = 1
            Case bteColLotNo: .EditMaxLength = 7
            Case bteColQty
                .EditMaxLength = 7
                dblTempValue = .TextMatrix(Row, Col)
            Case bteColCtn
                .EditMaxLength = 7
                dblTempValue = .TextMatrix(Row, Col)
            Case bteColNet
                .EditMaxLength = 10
                dblTempValue = .TextMatrix(Row, Col)
            Case bteColGross
                .EditMaxLength = 10
                dblTempValue = .TextMatrix(Row, Col)
            Case bteColPrice
                .EditMaxLength = 18
                dblTempValue = .TextMatrix(Row, Col)
             Case bteColService
                .EditMaxLength = 18
                dblTempValue = .TextMatrix(Row, Col)
            Case bteColSerialNoFrom: .EditMaxLength = 7
            Case bteColSerialNoTo: .EditMaxLength = 7
            End Select
        End If
    Case Else: Cancel = True
    End Select
End With
End Sub

Private Sub gridAtas_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim sama As Boolean, nilRow As Integer
Dim itemCD As String, noPO As String, DeliveryDate As String, seqNo As Double
Dim Qty As String, Price As String, tampungQty As String
Dim service As String
Dim TempSF As String, TempST As String

LblErrMsg = ""
With gridAtas

If Row <> 0 Then
    If Col = bteColSelect Then
        itemCD = .TextMatrix(Row, bteColProdCode)
        noPO = .TextMatrix(Row, bteColPONo)
        Qty = .TextMatrix(Row, bteColQty)
        DeliveryDate = .TextMatrix(Row, bteColDate)
        seqNo = .TextMatrix(Row, bteColSeqNo)
        Price = .TextMatrix(Row, bteColPrice)
        service = .TextMatrix(Row, bteColService)
        tampungQty = .TextMatrix(Row, bteColQtySisaAwal)
        TempSF = GetMinSerial(noPO, itemCD, seqNo, Trim(txtDO))
        TempST = .TextMatrix(Row, bteColSerialNoTo)
        
        If .Cell(flexcpChecked, Row, bteColSelect) = 1 Then  'jika cek
            'Jika DP udah abis
            If Qty = 0 Then LblErrMsg = DisplayMsg("0041"): Exit Sub
            If defaultCurr = "" Then defaultCurr = .TextMatrix(Row, bteColCurr)
            
            For i = 1 To gridBawah.Rows - 1 'Utk mengecek apakah udah ada data di DO
                If itemCD & noPO & DeliveryDate & seqNo = _
                    gridBawah.TextMatrix(i, bteColProdCode) & gridBawah.TextMatrix(i, bteColPONo) & _
                    gridBawah.TextMatrix(i, bteColDate) & gridBawah.TextMatrix(i, bteColSeqNo) Then
                    sama = True 'Jika ada
                    Exit For
                Else
                    sama = False 'Jika tdk
                End If
            Next i

            If sama = True Then 'Jika telah ada data
                nilRow = i
            Else 'Jika tidak tambah baris
                gridBawah.Rows = gridBawah.Rows + 1
                nilRow = gridBawah.Rows - 1
            End If
            
            '********** Isi data yg dipilih dr Grid Atas ke Grid  Bawah ***********
            gridBawah.Cell(flexcpBackColor, nilRow, bteColSelect) = vbWhite
            gridBawah.TextMatrix(nilRow, bteColSelect) = ""
            gridBawah.TextMatrix(nilRow, bteColProdCode) = .TextMatrix(Row, bteColProdCode) 'Item Code
            gridBawah.TextMatrix(nilRow, bteColMakerItem) = .TextMatrix(Row, bteColMakerItem)   'Maker Item Code
            gridBawah.TextMatrix(nilRow, bteColDesc) = .TextMatrix(Row, bteColDesc) 'Desc
            gridBawah.TextMatrix(nilRow, bteColPONo) = .TextMatrix(Row, bteColPONo) 'PO
            
            gridBawah.Cell(flexcpBackColor, nilRow, bteColLotNo) = vbWhite
            gridBawah.Cell(flexcpBackColor, nilRow, bteColCtn) = vbWhite
            gridBawah.Cell(flexcpBackColor, nilRow, bteColNet) = vbWhite
            gridBawah.Cell(flexcpBackColor, nilRow, bteColGross) = vbWhite
            
            gridBawah.TextMatrix(nilRow, bteColLotNo) = Month(Format(.TextMatrix(Row, bteColDate), "dd-MMM-YYYY")) 'Unit Desc
            gridBawah.TextMatrix(nilRow, bteColUnitCls) = .TextMatrix(Row, bteColUnitCls) 'Unit Cls
            gridBawah.TextMatrix(nilRow, bteColUnit) = .TextMatrix(Row, bteColUnit) 'Unit Desc
            gridBawah.TextMatrix(nilRow, bteColCtn) = .TextMatrix(Row, bteColCtn)  'Qty Ctn
            gridBawah.TextMatrix(nilRow, bteColNet) = .TextMatrix(Row, bteColNet)  'Net Weight
            gridBawah.TextMatrix(nilRow, bteColGross) = .TextMatrix(Row, bteColGross)  'Gross Weight
            gridBawah.TextMatrix(nilRow, bteColDate) = .TextMatrix(Row, bteColDate) 'Delivery Date
            gridBawah.TextMatrix(nilRow, bteColTime) = .TextMatrix(Row, bteColTime) 'Delivery Time
            
            gridBawah.TextMatrix(nilRow, bteColCurrCls) = .TextMatrix(Row, bteColCurrCls) 'Curr Code
            gridBawah.TextMatrix(nilRow, bteColCurr) = .TextMatrix(Row, bteColCurr) 'Curr Desc
            
            gridBawah.Cell(flexcpBackColor, nilRow, bteColPrice) = vbWhite 'Price
            gridBawah.TextMatrix(nilRow, bteColPrice) = .TextMatrix(Row, bteColPrice) 'Price
            
            gridBawah.Cell(flexcpBackColor, nilRow, bteColService) = vbWhite  'Price
            gridBawah.TextMatrix(nilRow, bteColService) = .TextMatrix(Row, bteColService)  'Price
            
            'Jika kol ke bteColQtySisaAwal di Grid kosong berarti  ngga ada di DB
            gridBawah.Cell(flexcpBackColor, nilRow, bteColQty) = vbWhite 'Qty
            If gridBawah.TextMatrix(nilRow, bteColQtySisaAwal) = "" Then
                gridBawah.TextMatrix(nilRow, bteColQty) = Qty
            Else 'jika dr DB sehingga nilai di Bawah harus ditambah dengan Grid atas
                gridBawah.TextMatrix(nilRow, bteColQty) = CDbl(gridBawah.TextMatrix(nilRow, bteColQty)) + CDbl(Qty)
            End If
            
            gridBawah.TextMatrix(nilRow, bteColAmount) = itungAmount(gridBawah.TextMatrix(nilRow, bteColQty), gridBawah.TextMatrix(nilRow, bteColPrice), gridBawah.TextMatrix(nilRow, bteColService))
            .TextMatrix(Row, bteColQty) = 0 'Qty Atas Auto 0
            '*************************
            gridBawah.TextMatrix(nilRow, bteColSerialNoFrom) = .TextMatrix(Row, bteColSerialNoFrom)
            gridBawah.TextMatrix(nilRow, bteColSerialNoTo) = .TextMatrix(Row, bteColSerialNoTo)
            
'            gridBawah.TextMatrix(nilRow, bteColSerialNoFrom) = TempSF
'            gridBawah.TextMatrix(nilRow, bteColSerialNoTo) = TempST

            gridBawah.TextMatrix(nilRow, bteColFixOrder) = .TextMatrix(Row, bteColFixOrder) 'Fix
            gridBawah.TextMatrix(nilRow, bteColUpdate) = "u" 'Ubah
            gridBawah.TextMatrix(nilRow, bteColSeqNo) = .TextMatrix(Row, bteColSeqNo) 'Seq No
            gridBawah.TextMatrix(nilRow, bteColNetWeight) = .TextMatrix(Row, bteColNetWeight)   'Net Weight
            gridBawah.TextMatrix(nilRow, bteColGrossWeight) = .TextMatrix(Row, bteColGrossWeight)    'Net Weight
            gridBawah.TextMatrix(nilRow, bteColQtyPerCtn) = .TextMatrix(Row, bteColQtyPerCtn)    'Net Weight
            
            '**** langsung set ke data tersebut *****
            gridBawah.Row = nilRow: gridBawah.Col = bteColQty
            gridBawah.TopRow = nilRow
            gridBawah.FocusRect = flexFocusInset
            gridBawah.SetFocus
            '*********
            
            Call isiBatasDO(itemCD, noPO, DeliveryDate, seqNo, .TextMatrix(Row, bteColQty))
            
        Else 'Jika uncek
            .TextMatrix(Row, bteColQty) = CDbl(tampungQty)
            For i = 1 To gridBawah.Rows - 1 'Utk mengecek apakah udah ada data di DO
                If i > gridBawah.Rows - 1 Then Exit For
                If itemCD & noPO & DeliveryDate & seqNo = _
                    gridBawah.TextMatrix(i, bteColProdCode) & gridBawah.TextMatrix(i, bteColPONo) & _
                    gridBawah.TextMatrix(i, bteColDate) & gridBawah.TextMatrix(i, bteColSeqNo) Then
                    If gridBawah.TextMatrix(i, bteColQtySisaAwal) = "" Then 'Jika bukan dr DB
                        gridBawah.TextMatrix(i, bteColQty) = 0
                        gridBawah.TextMatrix(i, bteColAmount) = 0
                        gridBawah.TextMatrix(i, bteColUpdate) = "u" 'Ubah
                        gridBawah.TextMatrix(i, bteColQtySisa) = .TextMatrix(Row, bteColQty)
                    Else 'Jika dr DB
                        gridBawah.TextMatrix(i, bteColQty) = Format(gridBawah.TextMatrix(i, bteColQtySisaAwal), gs_formatQty)
                        gridBawah.TextMatrix(i, bteColAmount) = itungAmount(gridBawah.TextMatrix(i, bteColQty), gridBawah.TextMatrix(i, bteColPrice), gridBawah.TextMatrix(i, bteColService))
                        gridBawah.TextMatrix(i, bteColUpdate) = "u" 'Ubah
                        gridBawah.TextMatrix(i, bteColQtySisa) = .TextMatrix(Row, bteColQty)
                        gridBawah.TextMatrix(i, bteColSerialNoFrom) = Left(gridBawah.TextMatrix(i, bteTempSerial), 7)
                        gridBawah.TextMatrix(i, bteColSerialNoTo) = Right(gridBawah.TextMatrix(i, bteTempSerial), 7)
                     End If
                End If
            Next i
            If gridBawah.Rows = 1 Then defaultCurr = ""
        End If
        
        Call warnaGrid(Row, .TextMatrix(Row, bteColQty))
        .TextMatrix(Row, bteColAmount) = itungAmount(.TextMatrix(Row, bteColQty), Price, service)
        Call itungAmountText  'Itung Amount di TxtDo
        ubahAtas = True
    End If
End If
End With
End Sub

Sub isiBatasDO(itemCD As String, noPO As String, DeliveryDate As String, seqNo As Double, _
    batasDO As Double)

With gridBawah
    For i = 1 To .Rows - 1
        'Cari Semua Item en Po No yg sama dengan yg baru diinput supaya bisa cek Sisa Qty nya
        If itemCD & noPO & DeliveryDate & seqNo = _
                gridBawah.TextMatrix(i, bteColProdCode) & gridBawah.TextMatrix(i, bteColPONo) & _
                gridBawah.TextMatrix(i, bteColDate) & gridBawah.TextMatrix(i, bteColSeqNo) Then
            .TextMatrix(i, bteColQtySisa) = batasDO
        End If
    Next i
End With
End Sub

Sub gridBawah_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim batasDO As Double, posDP As Double
Dim sama As Boolean, nilRow As Long
Dim itemCD As String, noPO As String, DeliveryDate As String, seqNo As Double
Dim Qty As String, Price As String, tampungQty As String
Dim service As String
Dim lngTempRow As Long

Dim dblQtyCtn As Double
Dim dblNet As Double
Dim dblGross As Double

LblErrMsg = ""
With gridBawah
If Row <> 0 And Col <> bteColSelect Then
    If IsNumeric(.TextMatrix(Row, Col)) = False And Col <> bteColLotNo And Col <> bteColSerialNoFrom _
            And Col <> bteColSerialNoTo Then .TextMatrix(Row, Col) = 0
    
    itemCD = .TextMatrix(Row, bteColProdCode)
    noPO = .TextMatrix(Row, bteColPONo)
    Qty = .TextMatrix(Row, bteColQty)
    DeliveryDate = .TextMatrix(Row, bteColDate)
    Price = .TextMatrix(Row, bteColPrice)
    service = .TextMatrix(Row, bteColService)
    tampungQty = .TextMatrix(Row, bteColQtySisaAwal)
    seqNo = .TextMatrix(Row, bteColSeqNo)
        
    If .Col = bteColQty Then 'Cek Nilai Quantity
        sama = False
        For i = 1 To gridAtas.Rows - 1
            'Cari yg sama dengan yg diatas
            If itemCD & noPO & DeliveryDate & seqNo = _
                    gridAtas.TextMatrix(i, bteColProdCode) & gridAtas.TextMatrix(i, bteColPONo) & _
                    gridAtas.TextMatrix(i, bteColDate) & gridAtas.TextMatrix(i, bteColSeqNo) Then
                sama = True
                nilRow = i
                Exit For
            End If
        Next i
    
        If .TextMatrix(Row, bteColQtySisa) = "" Then
            batasDO = nilQty(cbo(0), itemCD, noPO, Format(DeliveryDate, "yyyy-MM-dd"), seqNo)
        Else
            batasDO = CDbl(.TextMatrix(Row, bteColQtySisa)) + dblTempValue
        End If
                
        'Jika ada di Invoice cek Batas DO dan Invoice
        If CDbl(Qty) > batasDO Then
            'Jika tdk ada di Invoice cek Batas DO saja
            .TextMatrix(Row, bteColQty) = Format(batasDO, gs_formatQty)
            LblErrMsg = DisplayMsg(4045) & " " & Format(batasDO, gs_formatQty)
        End If
        
        batasDO = batasDO - .TextMatrix(Row, bteColQty)
        
        If sama = True Then 'Jika ada Grid Atas
            gridAtas.TextMatrix(nilRow, bteColQty) = Format(batasDO, gs_formatQty)
            gridAtas.TextMatrix(nilRow, bteColAmount) = itungAmount(gridAtas.TextMatrix(nilRow, bteColQty), gridAtas.TextMatrix(nilRow, bteColPrice), gridAtas.TextMatrix(nilRow, bteColService))
            Call warnaGrid(nilRow, CDbl(gridAtas.TextMatrix(nilRow, bteColQty)))
        End If
        
        Call isiBatasDO(itemCD, noPO, DeliveryDate, seqNo, batasDO)
    End If
    
    Select Case Col
    Case bteColQty
        If CDbl(.TextMatrix(Row, bteColQty)) > gd_MaxQty Then
            LblErrMsg = DisplayMsg(4045) & " " & gd_MaxQty & " !"
            .TextMatrix(Row, Col) = dblTempValue
            If gridAtas.TextMatrix(.TextMatrix(Row, bteColSeqNo), bteColSerialNoFrom) = "" Then
                .TextMatrix(.TextMatrix(Row, bteColSeqNo), bteColSerialNoFrom) = ""
                .TextMatrix(.TextMatrix(Row, bteColSeqNo), bteColSerialNoTo) = ""
            Else
                .TextMatrix(Row, bteColSerialNoFrom) = Left(Trim(.TextMatrix(Row - 1, bteColSerialNoTo)), 1) & Format(Val(Right(Trim(.TextMatrix(Row - 1, bteColSerialNoTo)), 6)) + 1, "000000")
                .TextMatrix(Row, bteColSerialNoTo) = Left(Trim(.TextMatrix(Row, bteColSerialNoFrom)), 1) & Format(Val(Right(Trim(.TextMatrix(Row, bteColSerialNoFrom)), 6)) + CDbl(.TextMatrix(Row, bteColQty)) - 1, "000000")
            End If
            .SetFocus
        End If
        .TextMatrix(Row, Col) = Format(.TextMatrix(Row, Col), gs_formatQty)
        If .TextMatrix(Row, bteColSerialNoFrom) <> "" Then
            If CDbl(.TextMatrix(Row, bteColQty)) <> 0 Then
                .TextMatrix(Row, bteColSerialNoTo) = Left(Trim(.TextMatrix(Row, bteColSerialNoFrom)), 1) & Format(Val(Right(Trim(.TextMatrix(Row, bteColSerialNoFrom)), 6)) + CDbl(.TextMatrix(Row, bteColQty)) - 1, "000000")
            Else
                .TextMatrix(Row, bteColSerialNoFrom) = ""
                .TextMatrix(Row, bteColSerialNoTo) = ""
            End If
        Else
            If CDbl(.TextMatrix(Row, bteColQty)) <> 0 Then
                If CDbl(gridAtas.TextMatrix(.TextMatrix(Row, bteColSeqNo), bteColQty)) >= CDbl(.TextMatrix(Row, bteColQty)) Then
                    If gridAtas.TextMatrix(.TextMatrix(Row, bteColSeqNo), bteColSerialNoFrom) = "" Then
                        .TextMatrix(.TextMatrix(Row, bteColSeqNo), bteColSerialNoFrom) = ""
                        .TextMatrix(.TextMatrix(Row, bteColSeqNo), bteColSerialNoTo) = ""
                    Else
                        .TextMatrix(Row, bteColSerialNoFrom) = Left(Trim(.TextMatrix(Row - 1, bteColSerialNoTo)), 1) & Format(Val(Right(Trim(.TextMatrix(Row - 1, bteColSerialNoTo)), 6)) + 1, "000000")
                        .TextMatrix(Row, bteColSerialNoTo) = Left(Trim(.TextMatrix(Row, bteColSerialNoFrom)), 1) & Format(Val(Right(Trim(.TextMatrix(Row, bteColSerialNoFrom)), 6)) + CDbl(.TextMatrix(Row, bteColQty)) - 1, "000000")
                    End If
                Else
                    .TextMatrix(Row, bteColSerialNoFrom) = ""
                    .TextMatrix(Row, bteColSerialNoTo) = ""
                End If
            Else
                .TextMatrix(Row, bteColSerialNoFrom) = ""
                .TextMatrix(Row, bteColSerialNoTo) = ""
            End If
        End If
        
    Case bteColCtn
        If CDbl(.TextMatrix(Row, bteColCtn)) > gd_MaxBox Then
            LblErrMsg = DisplayMsg(4037) & " " & gd_MaxBox & " !"
            .TextMatrix(Row, Col) = dblTempValue
            .SetFocus
        End If
        .TextMatrix(Row, Col) = Format(.TextMatrix(Row, Col), gs_formatBox)
    Case bteColNet
        If CDbl(.TextMatrix(Row, bteColNet)) > gd_MaxWeight Then
            LblErrMsg = DisplayMsg(8030) & " " & gd_MaxWeight & " !"
            .TextMatrix(Row, Col) = dblTempValue
            .SetFocus
        End If
        .TextMatrix(Row, Col) = Format(.TextMatrix(Row, Col), gs_formatWeight)
    Case bteColGross
        If CDbl(.TextMatrix(Row, bteColGross)) > gd_MaxWeight Then
            LblErrMsg = DisplayMsg(8030) & " " & gd_MaxWeight & " !"
            .TextMatrix(Row, Col) = dblTempValue
            .SetFocus
        End If
        .TextMatrix(Row, Col) = Format(.TextMatrix(Row, Col), gs_formatWeight)
    Case bteColPrice
        If CDbl(.TextMatrix(Row, bteColPrice)) > gd_MaxPrice Then
            LblErrMsg = DisplayMsg(4048) & " " & gd_MaxPrice & " !"
            .TextMatrix(Row, Col) = dblTempValue
            Price = dblTempValue
            .SetFocus
        End If
        If .TextMatrix(Row, bteColCurr) = "IDR" Then
        .TextMatrix(Row, Col) = Format(.TextMatrix(Row, Col), gs_formatPriceIDR)
        Else
        .TextMatrix(Row, Col) = Format(.TextMatrix(Row, Col), gs_formatPrice)
        End If
    
    Case bteColService
        If CDbl(.TextMatrix(Row, bteColService)) > gd_MaxPrice Then
            LblErrMsg = DisplayMsg(4048) & " " & gd_MaxPrice & " !"
            .TextMatrix(Row, Col) = dblTempValue
            service = dblTempValue
            .SetFocus
        End If
        If .TextMatrix(Row, bteColCurr) = "IDR" Then
        .TextMatrix(Row, Col) = Format(.TextMatrix(Row, Col), gs_formatPriceIDR)
        Else
        .TextMatrix(Row, Col) = Format(.TextMatrix(Row, Col), gs_formatPrice)
        End If
    
    Case bteColSerialNoFrom
        If Trim(.TextMatrix(Row, Col)) <> "" Then .TextMatrix(Row, bteColSerialNoTo) = GetSerialTo(Trim(.TextMatrix(Row, Col)), CDbl(.TextMatrix(Row, bteColQty)))
        
    End Select
    
    .TextMatrix(Row, bteColAmount) = itungAmount(.TextMatrix(Row, bteColQty), Price, service)
    
    If CDbl(.TextMatrix(Row, bteColQtyPerCtn)) = 0 Then
        dblQtyCtn = 0
        dblNet = 0
        dblGross = 0
    Else
        If .Col = bteColCtn Then
            dblQtyCtn = CDbl(.TextMatrix(Row, Col))
        Else
            If IsNumeric(.TextMatrix(Row, Col)) Then
                dblQtyCtn = CDbl(.TextMatrix(Row, Col)) / CDbl(.TextMatrix(Row, bteColQtyPerCtn))
            End If
        End If
        
        dblNet = dblQtyCtn * CDbl(.TextMatrix(Row, bteColNetWeight))
        dblGross = dblQtyCtn * CDbl(.TextMatrix(Row, bteColGrossWeight))
    End If
    
    .TextMatrix(Row, bteColCtn) = Format(dblQtyCtn, gs_formatBox)
    .TextMatrix(Row, bteColNet) = Format(dblNet, gs_formatWeight)
    .TextMatrix(Row, bteColGross) = Format(dblGross, gs_formatWeight)
    
ElseIf .Row <> 0 And .Col = bteColSelect And .TextMatrix(Row, bteColSelect) = "C" Then 'Utk Buat Lot NO again
    If cbo(2) = "" Then
        cbo(2) = "ALL"
    End If
    Call isiGridAtas
    
    .Rows = .Rows + 1
    .Cell(flexcpBackColor, .Rows - 1, bteColSelect) = vbWhite 'Lot No
    .TextMatrix(.Rows - 1, bteColProdCode) = .TextMatrix(Row, bteColProdCode) 'Product Code
    .TextMatrix(.Rows - 1, bteColMakerItem) = .TextMatrix(Row, bteColMakerItem)   'Part Number
    .TextMatrix(.Rows - 1, bteColDesc) = .TextMatrix(Row, bteColDesc) 'Desc
    
    .Cell(flexcpBackColor, .Rows - 1, bteColLotNo) = vbWhite 'Lot No
    .TextMatrix(.Rows - 1, bteColLotNo) = "" 'Lot No
    
    .TextMatrix(.Rows - 1, bteColPONo) = .TextMatrix(Row, bteColPONo) 'PO No
    
    .Cell(flexcpBackColor, .Rows - 1, bteColQty) = vbWhite  'Qty
    .TextMatrix(.Rows - 1, bteColQty) = 0 'Qty
    .TextMatrix(.Rows - 1, bteColUnitCls) = .TextMatrix(Row, bteColUnitCls) 'Unit Cls
    .TextMatrix(.Rows - 1, bteColUnit) = .TextMatrix(Row, bteColUnit) 'Desc
    
    .TextMatrix(.Rows - 1, bteColDate) = .TextMatrix(Row, bteColDate) 'Delivery Plan
    .TextMatrix(.Rows - 1, bteColTime) = .TextMatrix(Row, bteColTime) 'Time
    
    .Cell(flexcpBackColor, .Rows - 1, bteColCtn) = vbWhite 'Ctn Qty
    .TextMatrix(.Rows - 1, bteColCtn) = .TextMatrix(Row, bteColCtn) 'Ctn Qty
    
    .Cell(flexcpBackColor, .Rows - 1, bteColNet) = vbWhite 'Net Weight
    .TextMatrix(.Rows - 1, bteColNet) = .TextMatrix(Row, bteColNet) 'Net Weight
    
    .Cell(flexcpBackColor, .Rows - 1, bteColGross) = vbWhite 'Gross Weight
    .TextMatrix(.Rows - 1, bteColGross) = .TextMatrix(Row, bteColGross) 'Gross Weight
    
    .TextMatrix(.Rows - 1, bteColCurrCls) = .TextMatrix(Row, bteColCurrCls) 'Curr Cls
    .TextMatrix(.Rows - 1, bteColCurr) = .TextMatrix(Row, bteColCurr) 'Curr Desc
    .Cell(flexcpBackColor, .Rows - 1, bteColPrice) = vbWhite 'Price
    .TextMatrix(.Rows - 1, bteColPrice) = .TextMatrix(Row, bteColPrice) 'Price
    
    .Cell(flexcpBackColor, .Rows - 1, bteColService) = vbWhite  'Price
    .TextMatrix(.Rows - 1, bteColService) = .TextMatrix(Row, bteColService)   'Price
    
    
    .TextMatrix(.Rows - 1, bteColAmount) = 0 'Amount
    .TextMatrix(.Rows - 1, bteColFixOrder) = .TextMatrix(Row, bteColFixOrder) 'Fix Order Entry
    .TextMatrix(.Rows - 1, bteColUpdate) = .TextMatrix(Row, bteColUpdate) 'Ubah
    .TextMatrix(.Rows - 1, bteColSeqNo) = .TextMatrix(Row, bteColSeqNo) 'Seq No
    .TextMatrix(.Rows - 1, bteColQtySisa) = .TextMatrix(Row, bteColQtySisa) 'Tampung Sisa Qty
    .TextMatrix(.Rows - 1, bteColQtyPerCtn) = .TextMatrix(Row, bteColQtyPerCtn) '
    .TextMatrix(.Rows - 1, bteColNetWeight) = .TextMatrix(Row, bteColNetWeight) '
    .TextMatrix(.Rows - 1, bteColGrossWeight) = .TextMatrix(Row, bteColGrossWeight) '
        
    .Col = bteColSeqNo
    .Sort = flexSortStringAscending
    .Col = bteColProdCode
    .Sort = flexSortStringAscending
'    .Col = bteColdo
'    .Sort = flexSortStringAscending
    .Col = bteColSelect
End If

If .Row <> 0 Then .TextMatrix(Row, bteColUpdate) = "u": Call itungAmountText  'Itung Amount di TxtDo
End With
End Sub

'*********** Validate Grid ******
Sub warnaGrid(Baris As Long, qtyAtas As Double)
With gridAtas
    If qtyAtas = 0 Then 'Jika nilainya 0
        .Cell(flexcpBackColor, Baris, 1, Baris, bteColAmount) = vbHighlight
        .Cell(flexcpForeColor, Baris, 1, Baris, bteColAmount) = vbWhite
    Else
        .Cell(flexcpBackColor, Baris, 1, Baris, bteColAmount) = &H80000018
        .Cell(flexcpForeColor, Baris, 1, Baris, bteColAmount) = vbBlack
    End If
End With
End Sub

Function itungAmount(jmlBrg As String, Harga As String, service As String)
    If jmlBrg = "" Or IsNumeric(jmlBrg) = False Then jmlBrg = 0
    If Harga = "" Or IsNumeric(Harga) = False Then Harga = 0
    If service = "" Or IsNumeric(service) = False Then service = 0
    itungAmount = CDbl(jmlBrg) * (Format(CDbl(Harga), gs_formatAmount) + CDbl(service))
    'Format(uf_Trunc(CDbl(jmlBrg) * (Format(CDbl(Harga), gs_formatAmountIDR) + CDbl(service)), gi_decimalDigitAmountIDR), gs_formatAmountIDR)
End Function

Sub itungAmountText()
Dim jmlAmount As Double
    jmlAmount = 0
    For i = 1 To gridBawah.Rows - 1
        jmlAmount = jmlAmount + CDbl(gridBawah.TextMatrix(i, bteColAmount))
    Next i
    txtDO = Format(jmlAmount, gs_formatAmount)
End Sub

'**** Mengitung Sisa dr Planning dan DO
Function nilQty(custCD As String, ItemCode As String, noPO As String, Tgl As Date, seqNo As Double) As Double
Dim rsQtyDO As New ADODB.Recordset
    sql = "select qty - " & _
        "(select isnull(sum(qty),0) from Delivery_Order " & _
        "where PO_No = a.PO_NO " & _
        "And Seq_No = a.Seq_No) as sisa " & _
        "from orderentry_Detail a Where " & _
        "a.Cust_Code ='" & Trim(custCD) & _
        "' and a.PO_NO ='" & noPO & _
        "' and a.Seq_No =" & seqNo
    Set rsQtyDO = Db.Execute(sql)
    
    nilQty = IIf(rsQtyDO.EOF, 0, rsQtyDO("sisa"))
    Set rsQtyDO = Nothing
End Function
'***********************

Private Sub gridBawah_Click()
With gridBawah
If .Row <> 0 Then
    If .Col = bteColSelect Or .Col = bteColQty Or .Col = bteColPrice Or .Col = bteColService Then .FocusRect = flexFocusInset Else .FocusRect = flexFocusNone
End If
End With
End Sub

Private Sub gridBawah_DblClick()
If lblFix.Visible = False Then
    If gridBawah.Rows > 1 Then
        l1.Caption = gridAtas.TextMatrix(gridBawah.TextMatrix(gridBawah.RowSel, bteColSeqNo), bteColSerialNoFrom)
        l2.Caption = gridAtas.TextMatrix(gridBawah.TextMatrix(gridBawah.RowSel, bteColSeqNo), bteColSerialNoTo)
        l3.Caption = gridBawah.TextMatrix(gridBawah.RowSel, bteColDOSeqNo)
        l4.Caption = gridBawah.TextMatrix(gridBawah.RowSel, bteColSeqNo)
        l5.Caption = gridBawah.TextMatrix(gridBawah.RowSel, bteColSerialNoFrom)
        l6.Caption = gridBawah.TextMatrix(gridBawah.RowSel, bteColSerialNoTo)

        Me.Enabled = False
        frmDODetail.Show
    End If
Else
    LblErrMsg = DisplayMsg(4046)
End If
End Sub

Private Sub gridBawah_KeyDown(KeyCode As Integer, Shift As Integer)
With gridBawah
    If KeyCode = vbKeyRight Or KeyCode = vbKeyTab Then
        If .Col = bteColSelect Then
            .Col = bteColDesc
        ElseIf .Col = bteColLotNo Then
            .Col = bteColPONo
        ElseIf .Col = bteColQty Then
            .Col = bteColUnit
        ElseIf .Col = bteColCtn Then
            .Col = bteColCtn
        ElseIf .Col = bteColNet Then
            .Col = bteColNet
        ElseIf .Col = bteColGross Then
            .Col = bteColCurr
        ElseIf .Col = bteColService Then
            .Col = bteColService
        Else
            .Col = -1
        End If
        .SetFocus
    ElseIf KeyCode = vbKeyLeft Then
        If .Col = bteColSelect Then
            .Col = bteColPrice + 1
        ElseIf .Col = bteColPrice Then
            .Col = bteColGross + 1
        ElseIf .Col = bteColGross Then
            .Col = bteColNet + 1
        ElseIf .Col = bteColNet Then
            .Col = bteColCtn + 1
        ElseIf .Col = bteColCtn Then
            .Col = bteColQty + 1
        ElseIf .Col = bteColQty Then
            .Col = bteColLotNo + 1
        ElseIf .Col = bteColLotNo Then
            .Col = bteColSelect + 1
        ElseIf .Col = bteColService Then
            .Col = bteColService + 1
        
        End If
        .SetFocus
    End If
End With
End Sub

Private Sub gridBawah_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    With gridBawah
        Select Case Col
        Case bteColSelect
            If InStr(1, "CD", UCase(Chr(KeyAscii))) = 0 Then KeyAscii = 0
        Case bteColQty, bteColCtn, bteColNet, bteColGross, bteColPrice, bteColService
            If InStr(1, "0123456789.", Chr(KeyAscii)) = 0 Then KeyAscii = 0
            If CDbl(.TextMatrix(.Row, bteColQty)) > gd_MaxQty Then KeyAscii = 0
            If CDbl(.TextMatrix(.Row, bteColPrice)) > gd_MaxPrice Then KeyAscii = 0
            If CDbl(.TextMatrix(.Row, bteColGross)) > gd_MaxWeight Then KeyAscii = 0
            If CDbl(.TextMatrix(.Row, bteColNet)) > gd_MaxWeight Then KeyAscii = 0
            If CDbl(.TextMatrix(.Row, bteColCtn)) > gd_MaxBox Then KeyAscii = 0
            If CDbl(.TextMatrix(.Row, bteColService)) > gd_MaxPrice Then KeyAscii = 0
            
        End Select
    End With
End Sub

Function dataError(Optional drUpdate As Integer) As Boolean
Dim pesanError As String
    
    
    If hakUpdate(Me.Name) = 0 Then _
        LblErrMsg = DisplayMsg(3008): dataError = True: Exit Function
    If cbo(2) = "" Then
        'cbo(2).ListIndex = 0
        LblErrMsg = DisplayMsg(1048)
        cbo(2).SetFocus
        dataError = True
        Exit Function
    End If
    If cbo(0) = "" Then
        LblErrMsg = DisplayMsg(1033)
        cbo(0).SetFocus
        dataError = True
    Else
        cbo(0) = cbo(0)
        If cbo(0).MatchFound = False Then
            LblErrMsg = DisplayMsg(4011)
            cbo(0).SetFocus
            dataError = True
        ElseIf cbo(1) = "" Then
            LblErrMsg = DisplayMsg(1035)
            If cbo(1).Enabled Then cbo(1).SetFocus
            dataError = True
        ElseIf cbo(1).Enabled Then
            cbo(1) = cbo(1)
            If cbo(1).MatchFound = False Then
                LblErrMsg = DisplayMsg(8129)
                cbo(1).SetFocus
                dataError = True
            ElseIf FixCls = 1 Then
                LblErrMsg = DisplayMsg(4046)
                 Command1(0).Enabled = False
                cbo(1).SetFocus
                dataError = True
                gantiDO = True
                If gantiDO = True Then Call isiGridBawah: gantiDO = False
                If gantiPO Or gantiDtAwal Or gantiDtAkhir Or ubahAtas Or gridAtas.Rows = 1 Then Call isiGridAtas: gantiPO = False: gantiDtAwal = False: gantiDtAkhir = False: ubahAtas = False
            Else
                cbo(2) = cbo(2)
                If cbo(2) <> "" And cbo(2).MatchFound = False Then
                    LblErrMsg = DisplayMsg(4015)
                    cbo(2).SetFocus
                    dataError = True
                ElseIf cmdUpdate.Caption = "Create" Then 'Blm klik Create
                    LblErrMsg = DisplayMsg(1038)
                    cmdUpdate.SetFocus
                    dataError = True
                End If
            End If
        End If
    End If
    'mengecek keberadaan warehouse. update by dudi januari 08
    If cboWH = "" Then
       LblErrMsg = DisplayMsg(31)
       cboWH.SetFocus
       dataError = True
    ElseIf cboWH.MatchFound = False Then
        LblErrMsg = DisplayMsg(4018)
        cboWH.SetFocus
        dataError = True
    End If
    
    ' Cek Forwarder
    If Text1.locked = False Then
        If Text1 = "" Then
            LblErrMsg = "Please Type Forwarder Name ! "
            Text1.SetFocus
            dataError = True
        End If
    End If
        
    If TxtForwarder = "" Then
       LblErrMsg = "Please Select Forwarder ! "
       cbo(3).SetFocus
       dataError = True
'    ElseIf cbo(3).MatchFound = False Then
'        lblErrMsg = DisplayMsg(4018)
'        cbo(3).SetFocus
'        dataError = True
    End If
    
'    If cboDeliveryCls.Text = "" Then
'        LblErrMsg = "Please Select Delivery Cls ! "
'        cboDeliveryCls.SetFocus
'        dataError = True
'    End If

End Function

Private Sub cmdUpdate_Click()
Dim strS As String
Dim RsFor As New ADODB.Recordset

'On Error Resume Next
Me.MousePointer = vbHourglass
    LblErrMsg = ""
'    If cbo(2) = "" Then
'        'cbo(2).ListIndex = 0
'        LblErrMsg = DisplayMsg(9001)
'        'Exit Sub
'    End If
    If dataError(1) Then Me.MousePointer = vbDefault: Exit Sub
    Command1(0).Enabled = True 'mengaktivkan Submit
    If CekInvoice Then
        Call isiGridAtas
        Call isiGridBawah
        LblErrMsg = DisplayMsg(4110)
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    
    ' Add Forwarder Trade for KAWAI need
    strS = "Select * From Trade_Master Where Trade_Code='" & Trim(TxtForwarder) & "' "
    
    If RsFor.State <> adStateClosed Then RsFor.Close
    Set RsFor = Db.Execute(strS)
    If RsFor.EOF Then
        strS = "Insert into Trade_Master(Trade_Code,Trade_Cls,Trade_Name) Values " & _
            "   ('" & Trim(TxtForwarder) & "','5','" & Trim(Text1) & "')"
         
         Db.Execute strS
    End If
    ' -----------------------------
    If cboStatus = "Create" Then
        
        Call simpanMaster
        LblErrMsg = DisplayMsg(1000)
    Else
        Call simpanMaster
        LblErrMsg = DisplayMsg(1101)
    End If
    Call SetDeliveryRange
    Call isiGridAtas
    Call isiGridBawah
    Text1.locked = True
    
Me.MousePointer = vbDefault
End Sub

Private Sub SetDeliveryRange()
Dim RsRange As New ADODB.Recordset
Dim strSQL As String, TempDate As String

strSQL = "Select Delivery_Date From Delivery_Order Where DO_No='" & cbo(1) & "' order by delivery_Date"

RsRange.Open strSQL, Db, adOpenDynamic, adLockReadOnly
If Not RsRange.EOF Then
    TempDate = Format(dtAwal, "yyyy-MM-dd")
    If RsRange.Fields("Delivery_date") < CDate(TempDate) Then
        dtAwal = Format(RsRange.Fields("Delivery_date"), "yyyy-MM-dd")
    End If
    
    TempDate = Format(dtAkhir, "yyyy-MM-dd")
    RsRange.MoveLast
    If RsRange.Fields("Delivery_date") > CDate(TempDate) Then
        dtAkhir = Format(RsRange.Fields("Delivery_date"), "yyyy-MM-dd")
    End If
End If
RsRange.Close

End Sub


Private Sub Command1_Click(Index As Integer)
    Me.MousePointer = vbHourglass
    Select Case Index
    Case 0: 'Submit
        If dataError Then Me.MousePointer = vbDefault: Exit Sub
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
        
        If Trim(CboRemarks) <> "" And Trim(LblRemarks.Caption) = "" Then
            
            Dim RSCekRemark As New ADODB.Recordset
            If RSCekRemark.State <> adStateClosed Then RSCekRemark.Close
            RSCekRemark.Open "Select *from Message Where MsgId ='0079'", Db, adOpenDynamic, adLockOptimistic
            If RSCekRemark.EOF = True Then
                LblErrMsg = "[0079] Please select remarks Id!"
                Me.MousePointer = vbDefault
                Exit Sub
            Else
                LblErrMsg = DisplayMsg("0079") 'Please select remarks Id!
                Me.MousePointer = vbDefault
                Exit Sub
            End If
            RSCekRemark.Close
            
        End If
        
        Call simpanMaster
        Call simpanDetail
        Call simpanMaster
    Case 1: 'Clear
        LblErrMsg = ""
        Call isiGridAtas
        If gridAtas.Rows > 1 Then
            Call isiGridBawah
        Else
            Call headerGrid(gridBawah)
        End If
        
        up_FillCboDelivery
    End Select
    Me.MousePointer = vbDefault
End Sub

Function totAmountDo() As Double
Dim rstot As New ADODB.Recordset
    sql = "select sum(amount) as tot from Delivery_Order where DO_NO ='" & cbo(1) & "'"
    Set rstot = Db.Execute(sql)
    
    totAmountDo = IIf(IsNull(rstot(0)), 0, rstot(0))
End Function

Function listPO() As String
Dim rsIsiPO As New ADODB.Recordset
Dim tampungPO As String

    sql = "Select distinct PO_NO from Delivery_ORder where Do_NO = '" & cbo(1) & "'"
    Set rsIsiPO = Db.Execute(sql)
    
    If rsIsiPO.EOF Then
        listPO = ""
    Else
        tampungPO = ""
        Do While Not rsIsiPO.EOF
            tampungPO = tampungPO & Trim(rsIsiPO(0)) & ", "
            sql = "select po_date from orderentry_master where po_no = '" & Trim(rsIsiPO(0)) & "'"
            rsIsiPO.MoveNext
        Loop
        listPO = Left(Trim(tampungPO), Len(Trim(tampungPO)) - 1)
    End If
    Set rsIsiPO = Nothing
End Function

Sub simpanMaster()
Dim tampung As String
Dim rsDOMaster As New ADODB.Recordset

'On Error Resume Next
nilKosong = True
    If cboStatus = "Create" Then
        sql = "Insert into DO_Master (Cust_Code, DO_NO, DO_Date, Amount, List_PO, Reissue_Cls, Revised_Cls, Fix_Cls, WHCode, forwarder_Code,Last_Update, Last_User," & vbCrLf & _
                "   Remarks_cls, Remarks, BC_type, BC40_No, BC40_Date, Delivery_Cls, No_Register) " & vbCrLf & _
                "       values ('" & cbo(0) & "', '" & cbo(1) & "', '" & Format(dtDO, "yyyy-MM-dd") & "', " & totAmountDo & ", '" & listPO & "', 0, 0, 0, '" & Trim(cboWH) & "','" & Trim(cbo(3)) & "', getdate(), '" & userLogin & "'," & vbCrLf & _
                "           '" & Trim(CboRemarks) & "','" & Trim(txtremarks) & "','" & Trim(cboBCType) & "','" & Trim(txtBCNo) & "','" & Format(dtpBCDate, "yyyy-MM-dd") & "','" & Trim(cboDeliveryCls.Text) & "', '" & Trim(txtRegisterNo.Text) & "')"
        Db.Execute sql
        
        '**** Handle No yg Sama saat yg Sama
        If InStr(1, err.Description, "Violation of PRIMARY KEY constraint") > 0 Then
            buatDoBaru
            Call simpanMaster
            Exit Sub
        End If
    
        cboStatus = "Update"
        cbo(1).Enabled = True
        cmdUpdate.Caption = "Update"
        
        tampung = cbo(1)
        Call isiCboDO
        cbo(1) = tampung
        txtDoNO = cbo(1)
    Else
        sql = "Update DO_Master " & _
            "set DO_Date ='" & Format(dtDO, "yyyy-MM-dd") & "', " & _
            "Amount = " & totAmountDo & ", " & _
            "List_PO = '" & listPO & "', " & _
            "WHCode = '" & Trim(cboWH) & "', " & _
            "Forwarder_Code = '" & Trim(cbo(3)) & "', " & _
            "BC_type = '" & Trim(cboBCType.Text) & "', " & _
            "BC40_No = '" & Trim(txtBCNo.Text) & "', " & _
            "BC40_Date = '" & Format(dtpBCDate, "yyyy-MM-dd") & "', " & _
            "Delivery_Cls = '" & Trim(cboDeliveryCls.Text) & "', " & _
            "No_Register ='" & Trim(txtRegisterNo.Text) & "', " & _
            "Last_Update = getdate(), Last_User = '" & userLogin & "',Remarks_Cls ='" & Trim(CboRemarks) & "', " & _
            "Remarks = '" & Trim(txtremarks) & "' " & _
            "where DO_NO = '" & cbo(1) & "' and Cust_Code = '" & cbo(0) & "'"
        Db.Execute sql
        
    End If
    Set rsDOMaster = Nothing
nilKosong = False
End Sub

Function maxDOSeqNo(brs As Integer) As Double
Dim rsmax As New ADODB.Recordset
    
    sql = "Select ISNULL(Max(DOSeq_No),0) + 1  DOSeqNo " & _
        "From Delivery_Order Where " & kondisi(brs, 1)
    Set rsmax = Db.Execute(sql)
    maxDOSeqNo = rsmax!DOSeqNo
End Function

Function kondisi(brs As Integer, Optional doSeq As Byte) As String
With gridBawah
    kondisi = " DO_No='" & Trim(cbo(1)) & _
            "' and PO_NO='" & Trim(.TextMatrix(brs, bteColPONo)) & _
            "' and Seq_No = " & CDbl(Trim(.TextMatrix(brs, bteColSeqNo)))
    
    If doSeq = 0 Then
        kondisi = kondisi & " and DOSeq_No = " & CDbl(Trim(.TextMatrix(brs, bteColDOSeqNo)))
    End If
End With
End Function

Sub simpanDetail()
Dim hapus As Integer, ubah As Integer, simpan As Integer
Dim nilPesan As String
Dim tipeUbah As Integer
Dim tanya
Dim rsCek As New ADODB.Recordset
Dim rsSimpanDO As New ADODB.Recordset
Dim itemCD As String, noPO As String, DeliveryDate As String, seqNo As Double, DOSeqNo As String
Dim Qty As String, tampungQty As String, Price As Double
Dim batasDO As Double, batasInvoice As Double
Dim service As Double

    prosesSimpan = True
    With gridBawah
        '** Utk meliat msg yg dipake
        hapus = 0
        ubah = 0
        simpan = 0
        
        For i = 1 To .Rows - 1
            itemCD = .TextMatrix(i, bteColProdCode)
            noPO = .TextMatrix(i, bteColPONo)
            Qty = .TextMatrix(i, bteColQty)
            DeliveryDate = .TextMatrix(i, bteColDate)
            service = IIf(.TextMatrix(i, bteColService) = "", 0, .TextMatrix(i, bteColService))
            seqNo = .TextMatrix(i, bteColSeqNo)
            DOSeqNo = .TextMatrix(i, bteColDOSeqNo)
            tampungQty = .TextMatrix(i, bteColQtySisaAwal)
            Price = .TextMatrix(i, bteColPrice)
            
            batasDO = nilQty(cbo(0), itemCD, noPO, Format(DeliveryDate, "yyyy-MM-dd"), CDbl(seqNo))
            'Jika Col ke bteColAmount Berisi berarti dr DB
            If tampungQty <> "" Then batasDO = batasDO + CDbl(Qty)
                        
            'Jika ada di Invoice cek Batas DO dan Invoice
            If CDbl(Qty) > batasDO Then
                'Jika tdk ada di Invoice cek Batas DO saja
                LblErrMsg = DisplayMsg(4045) & " " & Format(batasDO, gs_formatQty)
                .Row = i: .Col = bteColQty
                .TopRow = i: .SetFocus
                prosesSimpan = False: Exit Sub
            End If
        Next i
        
        '***** Cek Grid Bawah APakah Ada Delete atau hanya Update/Insert
        For i = 1 To .Rows - 1
            If .TextMatrix(i, bteColUpdate) = "u" Then
                If .TextMatrix(i, bteColSelect) <> "D" Then 'kolom pertama kosong utk update atau insert
                    If .TextMatrix(i, bteColQty) = 0 Then  'jika qty = 0 maka di delete
                        If .TextMatrix(i, bteColDOSeqNo) <> "" Then
                            sql = "delete Delivery_Order where" & kondisi(i)
                            Db.Execute sql
                            hapus = hapus + 1
                        End If
                    Else
                        If Trim(.TextMatrix(i, bteColDOSeqNo)) = "" Then .TextMatrix(i, bteColDOSeqNo) = maxDOSeqNo(i): DOSeqNo = .TextMatrix(i, bteColDOSeqNo)
                        
                        sql = "select DO_NO from Delivery_Order where " & _
                            kondisi(i)
                        Set rsSimpanDO = Db.Execute(sql)
                        
                        If rsSimpanDO.EOF Then 'Jika data tidak ada
                            sql = "insert into Delivery_Order (DO_NO, Item_Code, MakerItem_Code, " & _
                                "Delivery_Date, Delivery_Time, PO_NO, Seq_No, DOSeq_No, Qty, Unit_Cls, SerialNoFrom,SerialNoTo,CtnQty, NetWeight, GrossWeight, Currency_Code, Price, Service,Amount, Lot_No, Revised_Cls, Last_Update, Last_User) " & _
                                "values ('" & Trim(cbo(1)) & "','" & Trim(.TextMatrix(i, bteColProdCode)) & "','" & Trim(.TextMatrix(i, bteColMakerItem)) & "','" & _
                                Format(.TextMatrix(i, bteColDate), "yyyy-MM-dd") & "','" & Format(.TextMatrix(i, bteColTime), "HH:MM") & "','" & Trim(.TextMatrix(i, bteColPONo)) & "'," & _
                                CDbl(.TextMatrix(i, bteColSeqNo)) & "," & CDbl(.TextMatrix(i, bteColDOSeqNo)) & "," & CDbl(.TextMatrix(i, bteColQty)) & ",'" & Trim(.TextMatrix(i, bteColUnitCls)) & "','" & _
                                Trim(.TextMatrix(i, bteColSerialNoFrom)) & "','" & Trim(.TextMatrix(i, bteColSerialNoTo)) & "'," & _
                                CDbl(.TextMatrix(i, bteColCtn)) & "," & CDbl(.TextMatrix(i, bteColNet)) & "," & CDbl(.TextMatrix(i, bteColGross)) & ",'" & _
                                .TextMatrix(i, bteColCurrCls) & "'," & CDbl(.TextMatrix(i, bteColPrice)) & "," & CDbl(.TextMatrix(i, bteColService)) & "," & CDbl(.TextMatrix(i, bteColAmount)) & ",'" & Trim(.TextMatrix(i, bteColLotNo)) & "',0, getdate(), '" & userLogin & "')"
                            Db.Execute sql
                            simpan = simpan + 1
                        Else
                            sql = "update Delivery_Order " & _
                                "set makeritem_code = '" & Trim(.TextMatrix(i, bteColMakerItem)) & "', " & _
                                "Delivery_Date = '" & Format(.TextMatrix(i, bteColDate), "yyyy-MM-dd") & "', " & _
                                "Delivery_Time = '" & Format(.TextMatrix(i, bteColTime), "HH:MM") & "', " & _
                                "qty = " & CDbl(.TextMatrix(i, bteColQty)) & ",  " & _
                                "unit_cls = '" & Trim(.TextMatrix(i, bteColUnitCls)) & "', " & _
                                "SerialNoFrom='" & Trim(.TextMatrix(i, bteColSerialNoFrom)) & "'," & _
                                "SerialNoTo='" & Trim(.TextMatrix(i, bteColSerialNoTo)) & "'," & _
                                "ctnqty = " & CDbl(.TextMatrix(i, bteColCtn)) & ", " & _
                                "netweight = " & CDbl(.TextMatrix(i, bteColNet)) & ", " & _
                                "grossweight = " & CDbl(.TextMatrix(i, bteColGross)) & ", " & _
                                "currency_code ='" & .TextMatrix(i, bteColCurrCls) & "', " & _
                                "price = " & CDbl(.TextMatrix(i, bteColPrice)) & ",  " & _
                                "Service = " & CDbl(.TextMatrix(i, bteColService)) & ",  " & _
                                "Amount = " & CDbl(.TextMatrix(i, bteColAmount)) & ",  " & _
                                "Lot_No = '" & Trim(.TextMatrix(i, bteColLotNo)) & "', " & _
                                "Revised_Cls = 1, Last_Update = getdate(), Last_User = '" & userLogin & "' " & _
                                "where " & kondisi(i)
                            Db.Execute sql
                            ubah = ubah + 1
                        End If
                        Set rsSimpanDO = Nothing
                    End If
                    
                ElseIf .TextMatrix(i, bteColSelect) = "D" And Trim(.TextMatrix(i, bteColDOSeqNo)) <> "" Then 'Utk Hapus
                
                    If IsEmpty(tanya) Then _
                        tanya = MsgBox("Do You Really Want to Delete This Data?", vbQuestion & vbYesNo, "Confirmation")
                    
                    If tanya = vbYes Then
                        sql = "delete Delivery_Order where " & kondisi(i)
                        Db.Execute sql
                    Else
                         prosesSimpan = False: Exit Sub
                    End If
                    Set rsCek = Nothing
                    hapus = hapus + 1
                End If
            Else
                ubah = 1
            End If
        Next i
        
        '******* Ulang itung Jumlah utk Delete
        Call simpanMaster
        '*******
                           
        Call isiGridAtas
        Call isiGridBawah
        
        If hapus > 0 Then
            nilPesan = "1201" 'Hapus
        ElseIf ubah = 0 Then 'Ngga ada pengubahan
            nilPesan = "1000" 'Simpan
        Else
            nilPesan = "1101" 'Ubah
        End If
        LblErrMsg = DisplayMsg(nilPesan)
    End With
    prosesSimpan = False
End Sub

Private Sub cmdReport_Click(Index As Integer)
    
    Me.MousePointer = vbHourglass
    cbo(1) = cbo(1)
    If cbo(1) = "" Then
        LblErrMsg = DisplayMsg(1035)
        If cbo(1).Enabled Then cbo(1).SetFocus
    ElseIf cbo(1).MatchFound = False Then
        LblErrMsg = DisplayMsg(8129)
        If cbo(1).Enabled Then cbo(1).SetFocus
    Else
        Select Case Index
        Case 0
            Call DOPrintStatus("'" & cbo(1) & "'")
            Call DOReport("'" & cbo(1) & "'")
        Case 1
            Call DIReport("'" & cbo(1) & "'")
        End Select
    End If
    Me.MousePointer = vbDefault

End Sub

Private Sub cmdPage_Click(Index As Integer)
With gridAtas
Select Case Index
    Case 0:  'First Page
            .TopRow = 1
    Case 1:  'Prev Page
            If .TopRow < 6 Then
                .TopRow = 1
            Else
                .TopRow = .TopRow - 5
            End If
    Case 2:  'Next Page
        If .TopRow < .Rows - 1 Then .TopRow = .TopRow + 5
    Case 3:  'Bottom Page
        .TopRow = .Rows
End Select
End With
End Sub



Private Sub TxtBCNo_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc("'") Then
    KeyAscii = 0
End If
End Sub

Private Sub txtDO_GotFocus()
    Command1(0).SetFocus
End Sub

Private Sub txtDoNO_GotFocus()
    Command1(0).SetFocus
End Sub

Sub HapusDOMaster()
    sql = "delete DO_Master where DO_No = '" & cbo(1) & "' And Do_No not in (select DO_NO from Delivery_ORder)"
    Db.Execute sql
End Sub

Private Sub CmdSubMenu_Click()
    Call HapusDOMaster
    DoEvents
    frmMainMenu.Show
    DoEvents
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    LblErrMsg.Caption = ErrMsg
End If
End Sub

Private Function CekInvoice() As Boolean
    Dim adoRs As New ADODB.Recordset
    sql = "select  do_no from invoice_detail where do_no = '" & Trim(cbo(1)) & "'"
    adoRs.Open sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
    CekInvoice = Not adoRs.EOF
    adoRs.Close
    Set adoRs = Nothing
End Function

Private Sub IsiComboWH()
    
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
    End With
    
    adoRs.Close
    Set adoRs = Nothing
    
End Sub

Private Sub IsiCboRemarks()
    
    Dim adoRs As New ADODB.Recordset
    
    sql = "select Remarks_cls, Description from Remarks_cls"
    adoRs.Open sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    With CboRemarks
        .clear
        .columnCount = 2
        .ColumnWidths = "60 pt; 180 pt"
        .ListWidth = 240
        .ListRows = 15
        While adoRs.EOF = False
            .AddItem ""
            .List(.ListCount - 1, 0) = Trim(adoRs.Fields("remarks_Cls"))
            .List(.ListCount - 1, 1) = Trim(adoRs.Fields("description"))
            adoRs.MoveNext
        Wend
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
    If gridBawah.Rows > 1 Then
        For Pos = 1 To gridBawah.Rows - 1
            ls_sql = " select * from orderEntry_detail " & _
                " where po_no='" & Trim(gridBawah.TextMatrix(Pos, bteColPONo)) & "' " & _
                " and item_code='" & Trim(gridBawah.TextMatrix(Pos, bteColProdCode)) & "' " & _
                " and Seq_no='" & Trim(gridBawah.TextMatrix(Pos, bteColSeqNo)) & "' "
            If RS.State <> adStateClosed Then RS.Close
            RS.CursorLocation = adUseClient
            RS.Open ls_sql, Db, adOpenKeyset, adLockOptimistic
            If RS.EOF = True Then
                found = True
                Exit For
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

Private Function GetMinSerial(a As String, b As String, C As Double, d As String) As String
Dim RSA As New ADODB.Recordset
Dim sqla As String

sqla = "Select Min(Serial_No) From Serial_Detail Where Item_Code='" & _
            Trim(b) & "' And PO_No ='" & Trim(a) & "' And Po_SeqNo='" & C & _
            "' And (DO_No='" & Trim(d) & "' Or DO_No is null) "

RSA.Open sqla, Db, adOpenForwardOnly, adLockReadOnly

If Not IsNull(RSA(0)) Then
    GetMinSerial = RSA(0)
Else
    GetMinSerial = ""
End If
RSA.Close
End Function


Private Sub TxtForwarder_Change()
Dim rsCari As New ADODB.Recordset
If Trim(TxtForwarder) <> "" Then
    rsCari.Open "Select * From Trade_Master Where Trade_Code='" & Trim(TxtForwarder) & "'", Db, adOpenForwardOnly, adLockReadOnly
    
    If Not rsCari.EOF Then
        Text1 = rsCari(2)
        Text1.locked = True
    Else
        Text1 = ""
        Text1.locked = False
    End If
End If


End Sub

Private Sub txtRemarks_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc("'") Then
    KeyAscii = 0
    End If
End Sub

Private Sub up_FillCboDelivery()
    '## receipt Combo
    cboDeliveryCls.clear
    cboDeliveryCls.columnCount = 2
    cboDeliveryCls.TextColumn = 1
    cboDeliveryCls.AddItem ""
    cboDeliveryCls.List(0, 0) = "D"
    cboDeliveryCls.List(0, 1) = "Delivery"
    cboDeliveryCls.AddItem ""
    cboDeliveryCls.List(1, 0) = "R1"
    cboDeliveryCls.List(1, 1) = "Return"
    cboDeliveryCls.ColumnWidths = "30 pt; 60 pt"
    cboDeliveryCls.ListWidth = 90
    cboDeliveryCls.ListRows = 4
    cboDeliveryCls.ListIndex = 0
End Sub

Private Sub uf_GetNoRegister()
    Dim rsGetNoSeri As New ADODB.Recordset
    
    If Trim(cbo(1).Text) <> "" Then
        sql = "EXEC dbo.sp_GetNoRegister '" & dtDO.Value & "', 'D', '" & userLogin & "' "
                
        If rsGetNoSeri.State <> adStateClosed Then rsGetNoSeri.Close
        rsGetNoSeri.Open sql, Db, adOpenForwardOnly, adLockReadOnly
        
        If Not rsGetNoSeri.EOF Then
           txtRegisterNo.Text = rsGetNoSeri.Fields("No_Register")
        End If
    End If
End Sub
