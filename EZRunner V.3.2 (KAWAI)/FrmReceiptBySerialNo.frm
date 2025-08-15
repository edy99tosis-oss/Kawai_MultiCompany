VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmReceiptBySerialNo 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Receipt By Serial No"
   ClientHeight    =   10335
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15090
   Icon            =   "FrmReceiptBySerialNo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10335
   ScaleWidth      =   15090
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrint 
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
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   37
      Tag             =   "FTTF*/"
      Top             =   9720
      Width           =   1125
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4680
      Top             =   9840
   End
   Begin VB.TextBox txtBarcode 
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
      MaxLength       =   100
      TabIndex        =   18
      Tag             =   "TTFF*/"
      Top             =   3720
      Width           =   3555
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FDDFE3&
      Height          =   2895
      Left            =   240
      TabIndex        =   9
      Tag             =   "TTTF*/"
      Top             =   720
      Width           =   14805
      Begin VB.TextBox lblToWH 
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
         Index           =   2
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   43
         Tag             =   "TTFF*/"
         Top             =   1930
         Width           =   1815
      End
      Begin VB.TextBox lblFromWH 
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
         Index           =   0
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   42
         Tag             =   "TTFF*/"
         Top             =   1560
         Width           =   1815
      End
      Begin VB.OptionButton OptWithoutSerialNo 
         BackColor       =   &H00FDDFE3&
         Caption         =   "Without Serial No"
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
         Left            =   1920
         TabIndex        =   39
         Top             =   200
         Width           =   1935
      End
      Begin VB.OptionButton OptSerialNo 
         BackColor       =   &H00FDDFE3&
         Caption         =   "With Serial No"
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
         Left            =   210
         TabIndex        =   38
         Top             =   200
         Width           =   1575
      End
      Begin VB.TextBox txtBCNo 
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
         Left            =   7080
         MaxLength       =   100
         TabIndex        =   34
         Tag             =   "TTFF*/"
         Top             =   1050
         Width           =   1515
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H0080FFFF&
         Caption         =   "&Search"
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
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   10
         Tag             =   "TTFF*/"
         Top             =   2400
         Width           =   1125
      End
      Begin MSComCtl2.DTPicker DtpFrom 
         Height          =   315
         Left            =   1920
         TabIndex        =   1
         Tag             =   "TTFF*/"
         Top             =   600
         Width           =   1485
         _ExtentX        =   2619
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
         Format          =   134086659
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker DtpTo 
         Height          =   315
         Left            =   3840
         TabIndex        =   11
         Tag             =   "TTFF*/"
         Top             =   600
         Width           =   1515
         _ExtentX        =   2672
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
         Format          =   134086659
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker dtpDNDate 
         Height          =   315
         Left            =   1920
         TabIndex        =   20
         Tag             =   "TTFF*/"
         Top             =   2400
         Width           =   1485
         _ExtentX        =   2619
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
         Format          =   134086659
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker dtpReceiptDate 
         Height          =   315
         Left            =   7080
         TabIndex        =   30
         Tag             =   "TTFF*/"
         Top             =   1980
         Width           =   1515
         _ExtentX        =   2672
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
         Format          =   134086659
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker dtpBCDate 
         Height          =   315
         Left            =   7080
         TabIndex        =   35
         Tag             =   "TTFF*/"
         Top             =   1510
         Width           =   1515
         _ExtentX        =   2672
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
         Format          =   134086659
         CurrentDate     =   37798
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   3480
         X2              =   5280
         Y1              =   2200
         Y2              =   2200
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   3480
         X2              =   5280
         Y1              =   1800
         Y2              =   1800
      End
      Begin MSForms.ComboBox cboToWH 
         Height          =   315
         Left            =   1920
         TabIndex        =   41
         Tag             =   "TTFF*/"
         Top             =   1920
         Width           =   1485
         VariousPropertyBits=   746604571
         MaxLength       =   25
         DisplayStyle    =   3
         Size            =   "2619;556"
         ListRows        =   15
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboFromWH 
         Height          =   315
         Left            =   1920
         TabIndex        =   40
         Tag             =   "TTFF*/"
         Top             =   1500
         Width           =   1485
         VariousPropertyBits=   746604571
         MaxLength       =   25
         DisplayStyle    =   3
         Size            =   "2619;556"
         ListRows        =   15
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblBCDate 
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
         Index           =   2
         Left            =   5760
         TabIndex        =   36
         Tag             =   "TTFF*/"
         Top             =   1530
         Width           =   720
      End
      Begin VB.Label lblBCNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BC No"
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
         Left            =   5760
         TabIndex        =   33
         Tag             =   "TTFF*/"
         Top             =   1080
         Width           =   540
      End
      Begin MSForms.ComboBox cboBCType 
         Height          =   315
         Left            =   7080
         TabIndex        =   32
         Tag             =   "TTFF*/"
         Top             =   600
         Width           =   1485
         VariousPropertyBits=   746604571
         MaxLength       =   25
         DisplayStyle    =   3
         Size            =   "2619;556"
         ListRows        =   15
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblBCType 
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
         Index           =   0
         Left            =   5760
         TabIndex        =   31
         Tag             =   "TTFF*/"
         Top             =   610
         Width           =   735
      End
      Begin VB.Label lblReceiptDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Receipt Date"
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
         Left            =   5760
         TabIndex        =   29
         Tag             =   "TTFF*/"
         Top             =   2000
         Width           =   1095
      End
      Begin VB.Label lblRemainingScan 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1935
         Left            =   11880
         TabIndex        =   26
         Tag             =   "FTTF*/"
         Top             =   840
         Width           =   2355
      End
      Begin VB.Label lblTotalSerial 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1935
         Left            =   9000
         TabIndex        =   25
         Tag             =   "FTTF*/"
         Top             =   840
         Width           =   2685
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H80000000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000A&
         FillColor       =   &H000000FF&
         Height          =   2055
         Left            =   11760
         Tag             =   "FTTF*/"
         Top             =   720
         Width           =   2535
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000A&
         FillColor       =   &H000000FF&
         Height          =   2055
         Left            =   9120
         Tag             =   "FTTF*/"
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Label19 
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
         Index           =   6
         Left            =   240
         TabIndex        =   24
         Tag             =   "TTFF*/"
         Top             =   2400
         Width           =   720
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Warehouse"
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
         Left            =   240
         TabIndex        =   23
         Tag             =   "TTFF*/"
         Top             =   1530
         Width           =   1455
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "REMAINING SCAN"
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
         Left            =   11700
         TabIndex        =   22
         Tag             =   "FTTF*/"
         Top             =   400
         Width           =   2640
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL SERIAL NO"
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
         Left            =   9120
         TabIndex        =   21
         Tag             =   "FTTF*/"
         Top             =   400
         Width           =   2520
      End
      Begin VB.Label Label19 
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
         Index           =   2
         Left            =   240
         TabIndex        =   16
         Tag             =   "TTFF*/"
         Top             =   610
         Width           =   720
      End
      Begin VB.Label Label1 
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
         Index           =   0
         Left            =   3480
         TabIndex        =   15
         Tag             =   "TTFF*/"
         Top             =   650
         Width           =   210
      End
      Begin VB.Label Label19 
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
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Tag             =   "TTFF*/"
         Top             =   1080
         Width           =   540
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Warehouose"
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
         Left            =   240
         TabIndex        =   13
         Tag             =   "TTFF*/"
         Top             =   2000
         Width           =   1335
      End
      Begin MSForms.ComboBox cboDNo 
         Height          =   315
         Left            =   1920
         TabIndex        =   12
         Tag             =   "TTFF*/"
         Top             =   1050
         Width           =   3405
         VariousPropertyBits=   746604571
         MaxLength       =   25
         DisplayStyle    =   3
         Size            =   "6006;556"
         ListRows        =   15
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Delete"
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
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   6
      Tag             =   "FTTF*/"
      Top             =   9720
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.CommandButton cmd_clear 
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
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "FTTF*/"
      Top             =   9720
      Width           =   1125
   End
   Begin VB.CommandButton cmd_sub_menu 
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
      TabIndex        =   4
      Tag             =   "TTFF*/"
      Top             =   9720
      Width           =   1125
   End
   Begin VB.CommandButton cmd_submit 
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
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "FTTF*/"
      Top             =   9720
      Width           =   1125
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   585
      Left            =   120
      TabIndex        =   0
      Tag             =   "TTTF*/"
      Top             =   9000
      Width           =   14835
      Begin VB.Label lbl_pesan 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "lbl_pesan"
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
         Left            =   0
         TabIndex        =   2
         Tag             =   "TTTF*/"
         Top             =   240
         Width           =   14715
      End
   End
   Begin EZRunnerv3.Anchor Anchor1 
      Left            =   2040
      Top             =   9720
      _ExtentX        =   847
      _ExtentY        =   820
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   435
      Left            =   13080
      TabIndex        =   7
      TabStop         =   0   'False
      Tag             =   "FTTF*/"
      Top             =   120
      Width           =   1845
      _extentx        =   3254
      _extenty        =   767
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   4605
      Left            =   120
      TabIndex        =   17
      TabStop         =   0   'False
      Tag             =   "TTTF*/"
      Top             =   4200
      Width           =   14775
      _cx             =   26061
      _cy             =   8123
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
      ExplorerBar     =   1
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
   Begin VB.TextBox Text1 
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
      Left            =   9000
      MaxLength       =   100
      TabIndex        =   27
      Tag             =   "TTFF*/"
      Top             =   9480
      Visible         =   0   'False
      Width           =   1275
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp 
      Height          =   375
      Left            =   2880
      TabIndex        =   28
      Top             =   9720
      Visible         =   0   'False
      Width           =   495
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   873
      _cy             =   661
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Scan Barcode No"
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
      TabIndex        =   19
      Tag             =   "TTTF*/"
      Top             =   3750
      Width           =   1470
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Receipt By Serial No"
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
      TabIndex        =   8
      Tag             =   "TTTF*/"
      Top             =   240
      Width           =   14565
   End
End
Attribute VB_Name = "FrmReceiptBySerialNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bteColSelect As Byte
Dim bteColBarcodeNo As Byte
Dim bteColSerialNo As Byte
Dim bteColItemCode As Byte
Dim bteColDescription As Byte
Dim bteColQty As Byte
Dim btecolDNNo As Byte
Dim btecolRSstatus As Byte
Dim SupplierCode As String, ItemCode As String, DNNo As String, WHFrom As String, WHTo As String, CurrencyCode As String
Dim Qty As Integer
Dim Price As Double, Amount As Double
Dim RSGetPrice As New Recordset
Dim RSGetCurr As New Recordset
Dim l_stock_location As String, SupplyCls As String, UnitCls As String, SJNo As String
Dim db2 As New ADODB.Connection
Dim iScan As Integer, iTotal As Integer, iTempScan As Integer, iScanCls As Integer
Dim Time As Long
Dim validate As Boolean


Private Sub up_Header()

    bteColSelect = 0
    bteColBarcodeNo = 1
    bteColSerialNo = 2
    bteColItemCode = 3
    bteColDescription = 4
    bteColQty = 5
    btecolRSstatus = 6
    
    With Grid
        .ColS = 7
        .Rows = 1
        
        .TextMatrix(0, bteColSelect) = ""
        .TextMatrix(0, bteColBarcodeNo) = "Barcode No"
        .TextMatrix(0, bteColSerialNo) = "Serial No"
        .TextMatrix(0, bteColItemCode) = "Item Code"
        .TextMatrix(0, bteColDescription) = "Description"
        .TextMatrix(0, bteColQty) = "Qty"
        .TextMatrix(0, btecolRSstatus) = "Status"
            
         .ColWidth(bteColSelect) = 300
         .ColWidth(bteColBarcodeNo) = 3250
         .ColWidth(bteColSerialNo) = 3000
         .ColWidth(bteColItemCode) = 2000
         .ColWidth(bteColDescription) = 3000
         .ColWidth(bteColQty) = 1000
         .ColWidth(btecolRSstatus) = 750
    
         .ColAlignment(bteColSelect) = flexAlignCenterCenter
         .ColAlignment(bteColBarcodeNo) = flexAlignCenterCenter
         .ColAlignment(bteColSerialNo) = flexAlignCenterCenter
         .ColAlignment(bteColItemCode) = flexAlignCenterCenter
         .ColAlignment(bteColDescription) = flexAlignCenterCenter
         .ColAlignment(bteColQty) = flexAlignCenterCenter
         .ColAlignment(btecolRSstatus) = flexAlignCenterCenter
              
        
    End With
End Sub

Private Sub cboDNo_Click()
    up_getWH
    up_Header
    cboBCType.Text = ""
    txtBCNo.Text = ""
    txtBarcode.Text = ""
    lblTotalSerial.Caption = 0
    lblRemainingScan.Caption = 0
    lbl_pesan = ""
    
    dtpBCDate.Value = Now
    dtpReceiptDate.Value = Now
    dtpDNDate.Value = Now
End Sub

Private Sub cboFromWH_Click()
cboFromWH = cboFromWH
    If cboFromWH.MatchFound Then
        lblFromWH(0) = cboFromWH.Column(1)
        lbl_pesan = ""
    Else
        lblFromWH(0) = ""
        lbl_pesan = DisplayMsg(4018)
    End If
End Sub

Private Sub cboToWH_Click()
cboToWH = cboToWH
    If cboToWH.MatchFound Then
        lblToWH(2) = cboToWH.Column(1)
        lbl_pesan = ""
    Else
        lblToWH(2) = ""
        lbl_pesan = DisplayMsg(4018)
    End If
End Sub

Private Sub cmd_clear_Click()
    up_Clear
End Sub

Private Sub cmd_sub_menu_Click()
    DoEvents
    frmMainMenu.Show
    DoEvents
    Unload Me
End Sub

Private Sub Cmd_Submit_Click()
Dim s As Integer, d As Integer, j  As Integer
Dim l_curr As String, sql_del As String, l_amount As String, l_qty As String, L_price As String, l_unit_cls As String, sql_prod As String
Dim RS As New ADODB.Recordset, ls_sql As String
    
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

    cmd_submit.Enabled = False
    
    If hakUpdate(Me.Name) = 0 Then lbl_pesan = DisplayMsg(3008): cmd_submit.Enabled = True: Exit Sub
    
    ls_sql = " SELECT Trade_Code, Trade_Name FROM Trade_Master " & vbCrLf & _
            " WHERE Trade_Code = '" & cboToWH & "' "
    If RS.State = adStateOpen Then RS.Close
    Set RS = Db.Execute(ls_sql)

    If Not RS.EOF Then
        l_stock_location = Trim(RS!trade_name)
    End If

    RS.Close
    
    lbl_pesan = up_ValidateDateRange(dtpDNDate.Value, True)
    If lbl_pesan.Caption <> "" Then cmd_submit.Enabled = True: cmd_submit.Enabled = True: Exit Sub
    
    Dim ls_ClosingMonth As String
    Dim ls_ClosingYear As String
    ls_ClosingMonth = uf_GetLastClosing("month")
    ls_ClosingYear = uf_GetLastClosing("year")
    
'    #Validate date Range
    lbl_pesan = up_ValidateDateRange(dtpDNDate.Value, True)
    If lbl_pesan <> "" Then cmd_submit.Enabled = True: cmd_submit.Enabled = True:   Exit Sub
    
    up_Validate
            
    If validate = True Then
        validate = False
        Exit Sub
    End If
    
    '#BEGIN TRANS
    db2.BeginTrans
    
    On Error GoTo errHandler
    
    Me.MousePointer = vbHourglass
    
    Set cmd = New ADODB.Command
    Set RSSuply = New ADODB.Recordset
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_SupplyScan_Sel"
    
    cmd.Parameters.append cmd.CreateParameter("FromwH", adVarChar, adParamInput, 15, RTrim(cboFromWH.Text))
    cmd.Parameters.append cmd.CreateParameter("ToWH", adVarChar, adParamInput, 15, RTrim(cboToWH.Text))
    cmd.Parameters.append cmd.CreateParameter("DateFrom", adDBTime, adParamInput, , DtpFrom)
    cmd.Parameters.append cmd.CreateParameter("DateTo", adDBTime, adParamInput, , DtpTo)
    cmd.Parameters.append cmd.CreateParameter("DNNo", adVarChar, adParamInput, 25, RTrim(cboDNo.Text))
    cmd.Parameters.append cmd.CreateParameter("Type", adVarChar, adParamInput, 1, "3")
    
    Set RSSupply = cmd.Execute
    
           
'        i = 1
'        With Grid
            If RSSupply.EOF = False Then
                
                While Not RSSupply.EOF
                    ItemCode = Trim(RSSupply("ChildItem_Code"))
                    SupplyCls = Trim(RSSupply("Supply_Cls"))
                    UnitCls = Trim(RSSupply("childunit_cls"))
                    CurrencyCode = up_GetCurrency(ItemCode)
                    Price = up_GetPrice(ItemCode)
                    Qty = Trim(RSSupply("ChildRequirement_Qty"))
                    Amount = Price * Qty
                    SJNo = Trim(RSSupply("SJNo"))
                    
                    Dim strSQL As String
                    Dim RSIns As ADODB.Recordset
                    Dim prm As ADODB.Parameter
                    
                    Set cmd = New ADODB.Command
                    cmd.CommandType = adCmdStoredProc
                    cmd.CommandTimeout = 0
                    cmd.ActiveConnection = Db
                    cmd.CommandText = "sp_ReceiptSerialNoInsertUpdate"
                    
                    Set prm1 = cmd.CreateParameter("FromWarehouse_Code", adVarChar, adParamInput, 50, RTrim(cboFromWH.Text))
                    cmd.Parameters.append prm1
                    Set prm2 = cmd.CreateParameter("ToWarehouse_Code", adVarChar, adParamInput, 50, RTrim(cboToWH.Text))
                    cmd.Parameters.append prm2
                    Set prm3 = cmd.CreateParameter("ChildSupply_date", adDate, adParamInput, , dtpReceiptDate.Value)
                    cmd.Parameters.append prm3
                    Set prm4 = cmd.CreateParameter("ChildItem_Code", adVarChar, adParamInput, 50, RTrim(ItemCode))
                    cmd.Parameters.append prm4
                    Set prm5 = cmd.CreateParameter("ChildUnit_Cls", adVarChar, adParamInput, 50, UnitCls)
                    cmd.Parameters.append prm5
                    Set prm6 = cmd.CreateParameter("Currency_Code", adVarChar, adParamInput, 50, CurrencyCode)
                    cmd.Parameters.append prm6
                    
                    Set prm7 = cmd.CreateParameter("Price", adDecimal, adParamInput, , Price)
                    prm7.Precision = 38
                    prm7.NumericScale = 4
                    cmd.Parameters.append prm7
                    
                    Set prm8 = cmd.CreateParameter("ChildeItemQty", adDecimal, adParamInput, , Qty)
                    prm8.Precision = 38
                    prm8.NumericScale = 4
                    cmd.Parameters.append prm8
                    
                    Set prm9 = cmd.CreateParameter("Amount", adDecimal, adParamInput, , Amount)
                    prm9.Precision = 38
                    prm9.NumericScale = 4
                    cmd.Parameters.append prm9
                    
                    Set prm10 = cmd.CreateParameter("SJNo", adVarChar, adParamInput, 50, RTrim(cboDNo.Text))
                    cmd.Parameters.append prm10
                    Set prm11 = cmd.CreateParameter("BCType", adVarChar, adParamInput, 50, RTrim(cboBCType.Text))
                    cmd.Parameters.append prm11
                    Set prm12 = cmd.CreateParameter("BCNo", adVarChar, adParamInput, 50, RTrim(txtBCNo.Text))
                    cmd.Parameters.append prm12
                    Set prm13 = cmd.CreateParameter("BC40_Date", adDate, adParamInput, , dtpBCDate.Value)
                    cmd.Parameters.append prm13
                    Set prm14 = cmd.CreateParameter("User", adVarChar, adParamInput, 50, userLogin)
                    cmd.Parameters.append prm14
                        
                    Set RSIns = cmd.Execute
                
                    '#Check if item influence the stock or not
'                    Call up_UpdateStockMaster(Format(dtpDNDate.Value, "yyyy-MM-dd"), ls_ClosingMonth, ls_ClosingYear, Trim(cboFromWH), Trim(cboToWH), Trim(ItemCode), CDbl(Qty), "S1", Trim(l_stock_location), "", "I", "", "", False, False, True, Db)
                    uf_UpdateSupplyStockMaster
                    
                    uf_UpdateReceiptStockMaster
                    
                    '#Update Stock Receipt
'                    Call up_UpdateStockMaster(Format(dtpDNDate.Value, "yyyy-MM-dd"), ls_ClosingMonth, ls_ClosingYear, Trim(cboFromWH), Trim(cboToWH), Trim(ItemCode), CDbl(Qty), "R", Trim(l_stock_location), "", "I", "", "", False, False, True, Db)
                RSSupply.MoveNext
                                
                Wend
                
                Set cmd = New ADODB.Command
                Set rsUpd = New ADODB.Recordset
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 0
                cmd.ActiveConnection = Db
                cmd.CommandText = "sp_SupplyScan_updStatus"
                
                cmd.Parameters.append cmd.CreateParameter("Receipt_User", adVarChar, adParamInput, 25, userLogin)
                cmd.Parameters.append cmd.CreateParameter("ReceiptDate", adDate, adParamInput, , dtpReceiptDate.Value)
                cmd.Parameters.append cmd.CreateParameter("SJNo", adVarChar, adParamInput, 50, RTrim(cboDNo.Text))
                
                Set rsUpd = cmd.Execute
                
                
                db2.CommitTrans
                
                up_GridLoad
                
                lbl_pesan.Caption = DisplayMsg(1000) '"Insert data success !"
            Else
                db2.CommitTrans
                lbl_pesan.Caption = DisplayMsg(9003) '"Insert data success !"
                
            End If
            
           
            
            
            
    
    
    
    Me.MousePointer = vbDefault
    
    
ErrExit:
    
    Exit Sub
    
errHandler:
    cmd_submit.Enabled = True
    db2.RollbackTrans
    lbl_pesan.Caption = "[" & err.number & "] " & err.Description
    Me.MousePointer = vbDefault
    err.clear
    Resume ErrExit

End Sub


Private Sub cmdPrint_Click()
    up_Print
End Sub

Private Sub cmdSearch_Click()
    up_GridLoad
End Sub

Private Sub DtpFrom_Change()
    If CDate(DtpFrom.Value) > (DtpTo.Value) Then
        lbl_pesan.Caption = DisplayMsg("4068") 'Start Date must be lower than End Date                ' "Start Date must be lower than " & DTPicker2akhir.Value & " !!!"
        DtpFrom.SetFocus
    Else
        up_FillComboDNNo
        If lbl_pesan <> "" Then Exit Sub
        up_FillComboBC
        lbl_pesan = ""
    End If
End Sub

Private Sub DtpTo_Change()
     If CDate(DtpTo.Value) > CDate(DtpFrom.Value) Then
        lbl_pesan.Caption = DisplayMsg("4066") ''End Date must be higher than Start Date
        DtpTo.SetFocus
    Else
        up_FillComboDNNo
        up_FillComboBC
        lbl_pesan = ""
    End If
End Sub

Private Sub Form_Load()
 CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
    
    If db2.State <> adStateClosed Then db2.Close
    db2.Open Db.ConnectionString
    
    Call up_Header
    Call up_FillComboDNNo
    up_FillComboBC
    Call up_Clear
    
    If StatusAdmin = 1 Then
        lblReceiptDate(7).Visible = True
        lblBCType(0).Visible = True
        lblBCNo(1).Visible = True
        lblBCDate(2).Visible = True
        
        cboBCType.Visible = True
        txtBCNo.Visible = True
        dtpBCDate.Visible = True
        dtpReceiptDate.Visible = True
       
    Else
        lblReceiptDate(7).Visible = False
'        lblBCType(0).Visible = False
'        lblBCNo(1).Visible = False
'        lblBCDate(2).Visible = False
'
'        cboBCType.Visible = False
'        txtBCNo.Visible = False
'        dtpBCDate.Visible = False
        dtpReceiptDate.Visible = False
    End If
    
    Time = 0
    
    With Anchor1
          .RegString = "AnchorCtrl,Positions," & Me.Name & "0|0"
          .DoInit
    End With
End Sub

Private Sub up_Clear()
        
    cboFromWH.Text = ""
    cboToWH.Text = ""
    cboDNo.Text = ""
    txtBarcode.Text = ""
    lblTotalSerial.Caption = 0
    lblRemainingScan.Caption = 0
    lblFromWH(0).Text = ""
    lblToWH(2).Text = ""
    lbl_pesan = ""
    
    DtpFrom.Value = DateSerial(Year(Now), Month(Now), 1)
    DtpTo.Value = Now()
    dtpDNDate.Value = Now()
    dtpReceiptDate.Value = Now()
    dtpBCDate.Value = Now()
    
    If OptWithoutSerialNo.Value = True Then
        cmd_submit.Enabled = True
    Else
        cmd_submit.Enabled = False
    End If
    
    up_FillComboBC
    up_FillComboWH
    up_Header
        
End Sub

Private Sub up_FillComboDNNo()
Dim sql As String
Dim RS As New Recordset
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_DNNo_Sel"
    
    cmd.Parameters.append cmd.CreateParameter("FromwH", adVarChar, adParamInput, 15, "")
    cmd.Parameters.append cmd.CreateParameter("ToWH", adVarChar, adParamInput, 15, "")
    cmd.Parameters.append cmd.CreateParameter("DateFrom", adDate, adParamInput, , Format(DtpFrom.Value, "YYYY-MM-DD"))
    cmd.Parameters.append cmd.CreateParameter("DateTo", adDate, adParamInput, , Format(DtpTo.Value, "YYYY-MM-DD"))
    cmd.Parameters.append cmd.CreateParameter("Type", adVarChar, adParamInput, 1, "")
    If OptSerialNo.Value = True Then
        cmd.Parameters.append cmd.CreateParameter("SerialNo", adVarChar, adParamInput, 1, "1")
    ElseIf OptWithoutSerialNo.Value = True Then
        cmd.Parameters.append cmd.CreateParameter("SerialNo", adVarChar, adParamInput, 1, "0")
    Else
        lbl_pesan = DisplayMsg(9017) & " Type Serial No "
        Exit Sub
    End If
    
    Set RS = cmd.Execute
    
    If RS.EOF = False Then

        With cboDNo
            .clear
            .columnCount = 1
            .ColumnWidths = "165pt"
            .ListWidth = 165
            .ListRows = 10
        
            i = 0
            
            Do While Not RS.EOF
                .AddItem
                .List(i, 0) = Trim(RS("SJ_No") & "")
                
                RS.MoveNext
                i = i + 1
            Loop
            
          .ListIndex = -1
        End With
    Else
        cboDNo.clear
    End If
    
End Sub

Private Sub up_getWH()
Dim sql As String
Dim RS As New Recordset

    sql = "EXEC sp_SupplyScan_GetHeader '" & cboDNo.Text & "'"
    Set RS = Db.Execute(sql)
    
    If RS.EOF = False Then
        cboFromWH.Text = Trim(RS("From_WH"))
        cboToWH.Text = Trim(RS("To_WH"))
        dtpDNDate.Value = Format(RS("SJ_Date"), "YYYY-MMM-DD")
    End If
End Sub

Private Sub up_FillComboBC()
Dim ls_sql As String
Dim rs_combo As New ADODB.Recordset
Dim i As Long


cboBCType.columnCount = 1
cboBCType.clear

ls_sql = "select BC_Type from BC_master"
rs_combo.Open ls_sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
i = 0

Do While Not rs_combo.EOF
cboBCType.AddItem rs_combo("BC_Type")
rs_combo.MoveNext
Loop

cboBCType.ColumnWidths = "75"
cboBCType.ListWidth = 75
cboBCType.ListRows = 7

End Sub

Private Sub up_FillComboWH()
Dim sql As String
Dim RS As New Recordset
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_WH_Sel"
    
    Set RS = cmd.Execute

    With cboFromWH
        .clear
        .columnCount = 2
        .ColumnWidths = "50pt;130pt"
        .ListWidth = 180
        .ListRows = 15
    
        i = 0
        
        Do While Not RS.EOF
            .AddItem
            .List(i, 0) = Trim(RS("WH_Code") & "")
            .List(i, 1) = Trim(RS("WH_Name") & "")
            
            RS.MoveNext
            i = i + 1
        Loop
        
        .ListIndex = -1
    End With
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_WH_Sel"
    
    Set RS = cmd.Execute

    With cboToWH
        .clear
        .columnCount = 2
        .ColumnWidths = "50pt;130pt"
        .ListWidth = 180
        .ListRows = 15
    
        i = 0
        
        Do While Not RS.EOF
            .AddItem
            .List(i, 0) = Trim(RS("WH_Code") & "")
            .List(i, 1) = Trim(RS("WH_Name") & "")
            
            RS.MoveNext
            i = i + 1
        Loop
        
        .ListIndex = -1
    End With
    
End Sub

Private Sub up_GridLoad()
Dim sql As String
Dim cmd As ADODB.Command
Dim li_Row As Integer
Dim PSStatus As Integer
Dim bcType As String, bcNo As String
Dim bcDate As Date, receiptDate As Date

Me.MousePointer = vbHourglass
     
up_Header
lbl_pesan = ""
bcType = ""
bcNo = ""
bcDate = Now
receiptDate = Now
iScan = 0
iScanCls = 0
iTempScan = 0
PSStatus = 0
        
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_SupplyScan_Sel"
    
    cmd.Parameters.append cmd.CreateParameter("FromwH", adVarChar, adParamInput, 15, RTrim(cboFromWH.Text))
    cmd.Parameters.append cmd.CreateParameter("ToWH", adVarChar, adParamInput, 15, RTrim(cboToWH.Text))
    cmd.Parameters.append cmd.CreateParameter("DateFrom", adDBTime, adParamInput, , DtpFrom)
    cmd.Parameters.append cmd.CreateParameter("DateTo", adDBTime, adParamInput, , DtpTo)
    cmd.Parameters.append cmd.CreateParameter("DNNo", adVarChar, adParamInput, 25, RTrim(cboDNo.Text))
    cmd.Parameters.append cmd.CreateParameter("Type", adVarChar, adParamInput, 1, "2")
    
    Set RS = cmd.Execute
        
        i = 1
        With Grid
            If RS.EOF = False Then
                While Not RS.EOF
                    .Rows = .Rows + 1
                    
                    .Cell(flexcpChecked, i, ColCheck) = flexUnchecked
                    .Cell(flexcpBackColor, i, ColCheck) = vbWhite
                    
                    .Cell(flexcpChecked, i, btecolRSstatus) = flexUnchecked
                    .Cell(flexcpBackColor, i, btecolRSstatus) = vbWhite
                    .Cell(flexcpPictureAlignment, i, btecolRSstatus) = flexAlignCenterCenter
                    
                    If RS("PSStatus") = 1 Then
                        .TextMatrix(i, btecolRSstatus) = flexChecked
                        .Cell(flexcpChecked, i, btecolRSstatus) = flexChecked
                        .TextMatrix(i, bteColSelect) = flexChecked
                        .Cell(flexcpChecked, i, ColCheck) = flexChecked
                        Grid.Cell(flexcpBackColor, i, ColCheck, i, btecolRSstatus) = vbGreen
                        iScan = iScan + 1
                        
                        If OptWithoutSerialNo.Value = True Then
                            cmd_submit.Enabled = True
                        Else
                            cmd_submit.Enabled = False
                        End If
    
                        PSStatus = 1
                    Else
                        cmd_submit.Enabled = True
                    End If
                    
                    If RS("Scan_Cls") = 1 Then
                        .TextMatrix(i, bteColSelect) = flexChecked
                        .Cell(flexcpChecked, i, ColCheck) = flexChecked
                        Grid.Cell(flexcpBackColor, i, ColCheck, i, btecolRSstatus) = vbGreen
                        iScanCls = iScanCls + 1
                        iTempScan = iScanCls
                    End If
                    
                    .TextMatrix(i, bteColSelect) = ""
                    .TextMatrix(i, bteColBarcodeNo) = IIf(IsNull(Trim(RS("Barcode_No"))) = True, "", Trim(RS("Barcode_No")))
                    .TextMatrix(i, bteColSerialNo) = IIf(IsNull(Trim(RS("Serial_No"))) = True, "", Trim(RS("Serial_No")))
                    .TextMatrix(i, bteColItemCode) = IIf(IsNull(Trim(RS("Item_Code"))) = True, "", Trim(RS("Item_Code")))
                    .TextMatrix(i, bteColDescription) = IIf(IsNull(Trim(RS("Description"))) = True, "", Trim(RS("Description")))
                    .TextMatrix(i, bteColQty) = IIf(IsNull(Trim(RS("Qty"))) = True, 0, Trim(RS("Qty")))
                    .TextMatrix(i, btecolRSstatus) = ""
                    .Cell(flexcpAlignment, i, bteColBarcodeNo, i, bteColDescription) = flexAlignLeftCenter
                    
                    i = i + 1
                    
                    bcType = IIf(IsNull(Trim(RS("BC_Type"))) = True, "", Trim(RS("BC_Type")))
                    bcNo = IIf(IsNull(Trim(RS("BC40_No"))) = True, "", Trim(RS("BC40_No")))
                    bcDate = IIf(IsNull(Trim(RS("BC40_Date"))) = True, Now, Trim(RS("BC40_Date")))
                    receiptDate = IIf(IsNull(Trim(RS("ChildSupply_date"))) = True, Now, Trim(RS("ChildSupply_date")))
                    dtpDNDate.Value = Trim(RS("SupplyDate"))
                    
                RS.MoveNext
                Wend
            End If
        End With
        
        cboBCType.Text = bcType
        txtBCNo.Text = bcNo
        dtpBCDate.Value = bcDate
        dtpReceiptDate.Value = receiptDate
        
        iTotal = i - 1
        lblTotalSerial.Caption = iTotal
        lblRemainingScan.Caption = iTotal - iScan
        
        If iTotal - iScanCls = 0 Then
            cmd_submit.Enabled = True
        Else
            If OptWithoutSerialNo.Value = True Then
                cmd_submit.Enabled = True
            Else
                cmd_submit.Enabled = False
            End If
            
        End If
        
        If PSStatus = 1 Then
            cmd_submit.Enabled = False
            iTotal = i - 1
            lblTotalSerial.Caption = iTotal
            lblRemainingScan.Caption = iTotal - iScan
        Else
            iTotal = i - 1
            lblTotalSerial.Caption = iTotal
            lblRemainingScan.Caption = iTotal - iScanCls
        End If
        
        
  Me.MousePointer = vbDefault
  
End Sub

Function up_GetPrice(ItemCode As String) As Double
       
    sql = "SELECT TOP 1 isnull(price,0) Price FROM Price_Master WHERE " & _
        "Item_Code='" & ItemCode & _
        "' and start_date<='" & Format(DtpFrom, "yyyymmdd") & _
        "' and end_date>='" & Format(DtpTo, "yyyymmdd") & _
        "' order by trade_code desc, priority_cls desc"
        
    Set RSGetPrice = Db.Execute(sql)
    
    If RSGetPrice.EOF = False Then
        up_GetPrice = Format(RSGetPrice("Price"), gs_formatAmountIDR)
    End If
    
End Function

Function up_GetCurrency(ItemCode As String) As String
       
    sql = "SELECT TOP 1 Currency_Code ,isnull(price,0) Price FROM Price_Master WHERE " & _
        "Item_Code='" & ItemCode & _
        "' and start_date<='" & Format(DtpFrom, "yyyymmdd") & _
        "' and end_date>='" & Format(DtpTo, "yyyymmdd") & _
        "' order by trade_code desc, priority_cls desc"
        
    Set RSGetCurr = Db.Execute(sql)
    
    If RSGetCurr.EOF = False Then
        up_GetCurrency = IIf(IsNull(RSGetCurr("Currency_Code")), "", RSGetCurr("Currency_Code"))
    End If
    
End Function

Private Sub cek_SeqNo()
Dim sql As String
Dim rsseqno As New Recordset
Dim RSTempSeqNo As New Recordset

    sql = "SELECT Seq_No FROM Part_Supply WHERE SJNo ='" & cboDNo.Text & "'"
    Set rsseqno = Db.Execute(sql)
    
    If rsseqno.EOF = True Then
        
        sql = "SELECT (MAX(Seq_No)+1) Seq_No FROM Part_Supply "
        Set RSTempSeqNo = Db.Execute(sql)
        
        If rsseqno.EOF = True Then
            TmpSeqNo = Trim(RSTempSeqNo("Seq_No"))
        End If
        
    End If
End Sub

Private Sub OptSerialNo_Click()
'cmd_submit.Enabled = False
up_Clear
up_FillComboDNNo
End Sub

Private Sub OptWithoutSerialNo_Click()
up_Clear
up_FillComboDNNo
End Sub

Private Sub Timer1_Timer()
Time = Time + 1

Text1.Text = Time

If Time = 20 Then
    Time = 0
    txtBarcode.Text = ""
End If

End Sub


Private Sub txtBarcode_KeyPress(KeyAscii As Integer)
Dim RS As ADODB.Recordset
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter

KeyAscii = Asc(UCase(Chr(KeyAscii)))

    For i = 1 To Grid.Rows - 1
        
        If Grid.Cell(flexcpChecked, i, ColCheck) = flexUnchecked Then
            If LTrim(txtBarcode.Text) = LTrim(Grid.TextMatrix(i, bteColBarcodeNo)) Then

                
                If iTempScan = 0 Then
                    iTempScan = iTotal - 1
                    lblRemainingScan.Caption = iTempScan
                Else
                    iTempScan = iTempScan - 1
                    lblRemainingScan.Caption = iTempScan
                End If
                
'                'Validasi Check Barcode No Sudah di Supply atau belum 20250130
'                 If CheckAlreadeyScan(Trim(cboDNo.Text), txtBarcode.Text) = True Then
'                    lbl_pesan = "Barcode No Already Supply"
'
'                    cmd_submit.Enabled = False
'                    wmp.URL = (App.path & "\Incorrect.mp3")
'                    Exit Sub
'                End If
                
                Grid.Cell(flexcpChecked, i, ColCheck) = flexChecked
                Grid.Cell(flexcpBackColor, i, ColCheck, i, btecolRSstatus) = vbGreen
                lbl_pesan = ""

                wmp.URL = (App.path & "\Correct.mp3")
                               
                Set cmd = New ADODB.Command
                cmd.CommandType = adCmdStoredProc
                cmd.CommandTimeout = 0
                cmd.ActiveConnection = Db
                cmd.CommandText = "sp_SupplyScan_UpdCls"
                
                cmd.Parameters.append cmd.CreateParameter("SJNo", adVarChar, adParamInput, 25, Trim(cboDNo.Text))
                cmd.Parameters.append cmd.CreateParameter("BarcodeNo", adVarChar, adParamInput, 100, txtBarcode.Text)
                cmd.Parameters.append cmd.CreateParameter("LastUpdate", adVarChar, adParamInput, 25, userLogin)
                                
                Set RS = cmd.Execute

                txtBarcode.Text = ""

                Time = 0
               
                If iTempScan = 0 Then cmd_submit.Enabled = True
                
                Exit Sub
            Else
                lbl_pesan = "Invalid Barcode No"
            End If
        Else
            If txtBarcode.Text = Grid.TextMatrix(i, bteColBarcodeNo) Then
                lbl_pesan = "Barcode No Already Scan"
            Else
                lbl_pesan = "Invalid Barcode No"
            End If
        End If
    Next i
    
    If lbl_pesan <> "" Then
        wmp.URL = (App.path & "\Incorrect.mp3")
    End If
        
End Sub

Sub uf_UpdateSupplyStockMaster()
    Dim RS As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim prm As ADODB.Parameter
    Dim li_Row As Integer

    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_update_stock_EZR"
    
    cmd.Parameters.append cmd.CreateParameter("TransDate", adDBTime, adParamInput, , Now)
    cmd.Parameters.append cmd.CreateParameter("WHCode", adVarChar, adParamInput, 15, cboFromWH.Text)
    cmd.Parameters.append cmd.CreateParameter("ItemCode", adVarChar, adParamInput, 25, ItemCode)
    cmd.Parameters.append cmd.CreateParameter("Status", adVarChar, adParamInput, 10, "S")
    Set prm = cmd.CreateParameter("Qty", adNumeric, adParamInput, , Qty)
    prm.Precision = 18
    prm.NumericScale = 5
    cmd.Parameters.append prm
    
    
    Set RS = cmd.Execute
    
End Sub

Sub uf_UpdateReceiptStockMaster()
    Dim RS As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim prm As ADODB.Parameter
    Dim li_Row As Integer

    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_update_stock_EZR"
    
    cmd.Parameters.append cmd.CreateParameter("TransDate", adDBTime, adParamInput, , Now)
    cmd.Parameters.append cmd.CreateParameter("WHCode", adVarChar, adParamInput, 15, cboToWH.Text)
    cmd.Parameters.append cmd.CreateParameter("ItemCode", adVarChar, adParamInput, 25, ItemCode)
    cmd.Parameters.append cmd.CreateParameter("Status", adVarChar, adParamInput, 10, "R")
    Set prm = cmd.CreateParameter("Qty", adNumeric, adParamInput, , Qty)
    prm.Precision = 18
    prm.NumericScale = 5
    cmd.Parameters.append prm
    
    
    Set RS = cmd.Execute
    
End Sub

Private Sub up_Print()
    Dim xlapp As New Excel.application
    Dim li_Row As Integer
    Dim RS As New ADODB.Recordset
    
    Dim ls_BarcodeNo As String
    Dim ls_SerialNo As String
    Dim ls_ItemCode As String
    Dim ls_description As String
    Dim ls_Qty As String
        
    ls_BarcodeNo = "A"
    ls_SerialNo = "B"
    ls_ItemCode = "C"
    ls_description = "D"
    ls_Qty = "E"
    
    Me.MousePointer = vbHourglass

    sql = "EXEC sp_SupplyScan_Sel '" & RTrim(cboFromWH.Text) & "', '" & RTrim(cboToWH.Text) & "', '" & DtpFrom & "', '" & DtpTo & "', " & vbCrLf & _
          " '" & RTrim(cboDNo.Text) & "', '" & 2 & "' "
            
    If RS.State = adStateOpen Then RS.Close
    RS.CursorLocation = adUseClient
    RS.Open sql, Db, adOpenKeyset, adLockOptimistic

    If Not RS.EOF Then
    
    Screen.MousePointer = vbHourglass
    With xlapp

    .Workbooks.Add
    .Range("A1") = "RECEIPT BY SERIAL NO"
    .Range("A3") = "DN Date"
    .Range("A4") = "DN No"
    .Range("A5") = "From WH"
    .Range("A6") = "To WH"
    .Range("A7") = "DN Date"
    .Range("A1:A7").Font.Bold = True
    
    .Range("D3") = "BC Type"
    .Range("D4") = "BC No"
    .Range("D5") = "BC Date"
    .Range("D6") = "Receipt Date"
    .Range("D3:D6").Font.Bold = True
        
    .Range("B3") = ": " & DtpFrom.Value
    .Range("B4") = ": " & cboDNo.Text
    .Range("B5") = ": " & cboDNo.Text
    .Range("B6") = ": " & cboFromWH.Text
    .Range("B7") = ": " & cboToWH.Text
    
    .Range("E3") = ": " & cboBCType
    .Range("E4") = ": " & txtBCNo
    .Range("E5") = ": " & dtpBCDate
    .Range("E6") = ": " & dtpReceiptDate
    
    
    .Range("A9") = "Barcode No"
    .Range("B9") = "Serial No"
    .Range("C9") = "Item Code"
    .Range("D9") = "Description"
    .Range("E9") = "Qty"
    .Range("A9:E9").Font.Bold = True
    .Range("A9:E9").horizontalAlignment = xlCenter

    .Range("A1:E1").Merge
    .Range("A1").Font.Size = 14
    .Range("A1").Font.Size = 18
    .Range("A1").Font.Bold = True
    .Range("A1").horizontalAlignment = xlCenter
    
    .ActiveSheet.Cells(1, 1).columnWidth = 21
    .ActiveSheet.Cells(1, 2).columnWidth = 13
    .ActiveSheet.Cells(1, 3).columnWidth = 14
    .ActiveSheet.Cells(1, 4).columnWidth = 25
    .ActiveSheet.Cells(1, 5).columnWidth = 12
    

    Row = 10

Dim jumlah As Double
    Do While Not RS.EOF
        .Range(ls_BarcodeNo & Row) = Trim(RS!Barcode_No)
        .Range(ls_SerialNo & Row) = Trim(RS!Serial_No)
        .Range(ls_ItemCode & Row) = "'" + Trim(RS!Item_Code)
        .Range(ls_description & Row) = Format(RS!Description)
        .Range(ls_Qty & Row) = (RS!Qty)
        
        Row = Row + 1
        RS.MoveNext
    Loop
    
    .Range(ls_BarcodeNo & 9, ls_Qty & Row - 1).Borders(xlEdgeTop).LineStyle = xlContinuous
    .Range(ls_BarcodeNo & 9, ls_Qty & Row - 1).Borders(xlEdgeBottom).LineStyle = xlContinuous
    .Range(ls_BarcodeNo & 9, ls_Qty & Row - 1).Borders(xlEdgeLeft).LineStyle = xlContinuous
    .Range(ls_BarcodeNo & 9, ls_Qty & Row - 1).Borders(xlEdgeRight).LineStyle = xlContinuous
    .Range(ls_BarcodeNo & 9, ls_Qty & Row - 1).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    .Range(ls_BarcodeNo & 9, ls_Qty & Row - 1).Borders(xlInsideVertical).LineStyle = xlContinuous
    
    
    .WindowState = xlMaximized
    .Visible = True
End With

Else
    lbl_pesan = DisplayMsg(4006)
End If

Screen.MousePointer = vbDefault
Me.MousePointer = vbDefault
   
    
End Sub

Private Sub up_Validate()
    If OptWithoutSerialNo.Value = True Then
         If cboBCType.Text = "" Then
              lbl_pesan = DisplayMsg(9017) & " BCType !": cboBCType.SetFocus:
              validate = True
         ElseIf txtBCNo.Text = "" Then
              lbl_pesan = DisplayMsg("0001") & " BCNo !": txtBCNo.SetFocus:
              validate = True
         End If
     End If
End Sub

'Public Function CheckAlreadeyScan(ByVal SJNo As String, BarcodeNo As String) As Boolean
'Dim RS As New ADODB.Recordset
'Dim sql As String
'
'sql = "EXEC dbo.sp_ReceiptSupplyScan_Check '" & SJNo & "', '" & BarcodeNo & "' "
'
'If RS.State <> adStateClosed Then RS.Close
'RS.Open sql, Db, adOpenForwardOnly, adLockReadOnly
'
'If Not RS.EOF Then
'    If RS.Fields("Scan_Cls").Value = "1" Then
'        CheckAlreadeyScan = True
'    Else
'        CheckAlreadeyScan = False
'    End If
'End If
'
'End Function

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With Grid
        If Col = bteColSelect Or Col = btecolRSstatus Then
            .Cell(flexcpChecked, Row, bteColSelect) = 1
            .Cell(flexcpChecked, Row, btecolRSstatus) = 1
        End If
    End With
End Sub

