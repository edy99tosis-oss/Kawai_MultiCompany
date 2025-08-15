VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmPart_Rec 
   BackColor       =   &H00FDDFE3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Part (Material) Receipt Scheduled"
   ClientHeight    =   10455
   ClientLeft      =   2040
   ClientTop       =   2130
   ClientWidth     =   15120
   Icon            =   "FrmPart_Rec.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10455
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtNoSeri 
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
      Left            =   6840
      MaxLength       =   4
      TabIndex        =   86
      Tag             =   "TTFF*/"
      Top             =   8970
      Width           =   1005
   End
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
      Left            =   9120
      MaxLength       =   25
      TabIndex        =   84
      Tag             =   "TTFF*/"
      Top             =   8970
      Width           =   1725
   End
   Begin VB.TextBox txtRec_Status 
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
      Height          =   315
      Left            =   120
      MaxLength       =   25
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmduploadbc 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Upload BC"
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
      Index           =   6
      Left            =   6930
      Style           =   1  'Graphical
      TabIndex        =   81
      Top             =   10020
      Width           =   1485
   End
   Begin VB.CommandButton cmdsj 
      BackColor       =   &H0080FFFF&
      Caption         =   "Update SJ"
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
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   80
      Top             =   10020
      Width           =   1185
   End
   Begin VB.CommandButton CmdupdateBC 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Update BC"
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
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   10020
      Width           =   1185
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   13080
      TabIndex        =   78
      Top             =   150
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   714
   End
   Begin VB.TextBox TxtBCNo 
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
      Left            =   5760
      TabIndex        =   75
      Top             =   7440
      Width           =   1275
   End
   Begin VB.TextBox txtSj 
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
      Left            =   1170
      TabIndex        =   74
      Top             =   7440
      Width           =   1785
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1695
      Left            =   345
      TabIndex        =   59
      Top             =   870
      Width           =   14595
      Begin VB.CommandButton cmd_Browser 
         Caption         =   "..."
         Height          =   300
         Left            =   6960
         TabIndex        =   82
         Top             =   270
         Width           =   300
      End
      Begin VB.CommandButton cmdCari 
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
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   1185
         Width           =   1185
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
         Left            =   9285
         MaxLength       =   25
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   1230
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
         Left            =   11820
         Style           =   1  'Graphical
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1125
      End
      Begin VB.TextBox txtWHSubcon 
         Appearance      =   0  'Flat
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
         Height          =   255
         Left            =   10920
         Locked          =   -1  'True
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   300
         Width           =   3015
      End
      Begin VB.TextBox Lblsupp 
         Appearance      =   0  'Flat
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
         Height          =   255
         Left            =   3420
         Locked          =   -1  'True
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   300
         Width           =   3510
      End
      Begin MSComCtl2.DTPicker Tgl1 
         Height          =   345
         Left            =   1785
         TabIndex        =   1
         Top             =   720
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   609
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
         Format          =   130940931
         CurrentDate     =   37868
      End
      Begin MSComCtl2.DTPicker Tgl2 
         Height          =   345
         Left            =   3720
         TabIndex        =   2
         Top             =   720
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   609
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
         Format          =   130940931
         CurrentDate     =   37868
      End
      Begin VB.Label LblPart 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Part No To Search"
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
         Left            =   7440
         TabIndex        =   71
         Top             =   1290
         Width           =   1560
      End
      Begin MSForms.ComboBox CboPart 
         Height          =   315
         Index           =   1
         Left            =   1785
         TabIndex        =   3
         Top             =   1200
         Width           =   3450
         VariousPropertyBits=   746604571
         MaxLength       =   35
         DisplayStyle    =   3
         Size            =   "6085;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label LblPart 
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
         Index           =   1
         Left            =   285
         TabIndex        =   66
         Top             =   1275
         Width           =   525
      End
      Begin VB.Line Line7 
         X1              =   10920
         X2              =   13920
         Y1              =   585
         Y2              =   585
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Con WH Code"
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
         Index           =   1
         Left            =   7440
         TabIndex        =   65
         Top             =   315
         Width           =   1785
      End
      Begin MSForms.ComboBox cmbbox_warehouse 
         Height          =   315
         Left            =   9285
         TabIndex        =   4
         Top             =   270
         Width           =   1515
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2672;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         Value           =   "cmbbox_warehouse"
         BorderColor     =   0
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label14 
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
         Left            =   3420
         TabIndex        =   64
         Top             =   795
         Width           =   165
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PO Issue Date"
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
         Left            =   285
         TabIndex        =   63
         Top             =   795
         Width           =   1230
      End
      Begin VB.Label LblPart 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Code "
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
         Left            =   285
         TabIndex        =   62
         Top             =   330
         Width           =   1275
      End
      Begin MSForms.ComboBox CboPart 
         Height          =   315
         Index           =   0
         Left            =   1785
         TabIndex        =   0
         Top             =   240
         Width           =   1515
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2672;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Line Line4 
         X1              =   3420
         X2              =   6920
         Y1              =   585
         Y2              =   585
      End
   End
   Begin VB.TextBox txtPackage 
      Alignment       =   1  'Right Justify
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
      Left            =   6045
      MaxLength       =   12
      TabIndex        =   18
      Top             =   10530
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.TextBox txtBC40 
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
      Left            =   13020
      MaxLength       =   30
      TabIndex        =   16
      Top             =   10530
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton CmdBC40 
      BackColor       =   &H0080FFFF&
      Caption         =   "B.C 4.0"
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
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   9960
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton cmdClear 
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
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   10020
      Width           =   1185
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
      Left            =   12445
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   10020
      Width           =   1185
   End
   Begin VB.TextBox LblItemName 
      Appearance      =   0  'Flat
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
      Height          =   255
      Left            =   2235
      Locked          =   -1  'True
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   8430
      Width           =   3495
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   345
      TabIndex        =   46
      Top             =   9330
      Width           =   14595
      Begin VB.Label LblErr 
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
         Height          =   255
         Left            =   105
         TabIndex        =   47
         Top             =   195
         Width           =   14370
      End
   End
   Begin VB.CommandButton CmdData 
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
      Left            =   7140
      Style           =   1  'Graphical
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   10020
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton CmdData 
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
      Left            =   5850
      Style           =   1  'Graphical
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   10020
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton CmdData 
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
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   10020
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton CmdData 
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
      Left            =   3270
      Style           =   1  'Graphical
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   10020
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.TextBox CboItem 
      BackColor       =   &H80000018&
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
      Left            =   465
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   8400
      Width           =   1695
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
      Left            =   13740
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   10020
      Width           =   1185
   End
   Begin VB.CommandButton CmdMenu 
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   10020
      Width           =   1185
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
      Height          =   315
      Left            =   1800
      MaxLength       =   50
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   8970
      Width           =   4065
   End
   Begin VB.TextBox txtqty 
      Alignment       =   1  'Right Justify
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
      Left            =   8325
      MaxLength       =   12
      TabIndex        =   10
      Top             =   8400
      Width           =   1200
   End
   Begin VB.TextBox txtamount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
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
      Left            =   12885
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   8400
      Width           =   1935
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   4650
      Left            =   345
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   2700
      Width           =   14595
      _cx             =   25744
      _cy             =   8202
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
      BackColorFixed  =   12640511
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483624
      BackColorAlternate=   -2147483624
      GridColor       =   8421504
      GridColorFixed  =   12582912
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
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
   Begin MSComCtl2.DTPicker TglReceipt 
      Height          =   345
      Left            =   13425
      TabIndex        =   7
      Top             =   7455
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   609
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
      Format          =   130940931
      CurrentDate     =   37868
   End
   Begin VB.CheckBox ChkComplete 
      BackColor       =   &H00FDDFE3&
      Caption         =   "Complete"
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
      Left            =   10065
      TabIndex        =   20
      Top             =   10575
      Visible         =   0   'False
      Width           =   1185
   End
   Begin MSComCtl2.DTPicker DtBCDate 
      Height          =   345
      Left            =   7800
      TabIndex        =   5
      Top             =   7440
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   609
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
      Format          =   130940931
      CurrentDate     =   37868
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No. Seri"
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
      Left            =   6000
      TabIndex        =   87
      Top             =   9030
      Width           =   690
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Register No."
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
      Left            =   8040
      TabIndex        =   85
      Top             =   9030
      Width           =   1050
   End
   Begin VB.Label LblPart 
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
      Index           =   9
      Left            =   3000
      TabIndex        =   77
      Top             =   7530
      Width           =   825
   End
   Begin MSForms.ComboBox CbotypeBC 
      Height          =   345
      Left            =   3780
      TabIndex        =   76
      Top             =   7440
      Width           =   1365
      VariousPropertyBits=   746604571
      MaxLength       =   15
      DisplayStyle    =   3
      Size            =   "2408;609"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label15 
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
      Left            =   7080
      TabIndex        =   72
      Top             =   7530
      Width           =   720
   End
   Begin VB.Label LblPart 
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
      Index           =   7
      Left            =   5190
      TabIndex        =   68
      Top             =   7530
      Width           =   645
   End
   Begin MSForms.ComboBox cboPackage 
      Height          =   315
      Left            =   6975
      TabIndex        =   19
      Top             =   10530
      Visible         =   0   'False
      Width           =   1005
      VariousPropertyBits=   746604571
      MaxLength       =   4
      DisplayStyle    =   3
      Size            =   "1773;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblPackage 
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
      Left            =   8055
      TabIndex        =   67
      Top             =   10560
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Label LblQty 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   7140
      TabIndex        =   58
      Top             =   10080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label LblItemCode 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ItemCode"
      Height          =   195
      Left            =   3270
      TabIndex        =   57
      Top             =   10110
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label LblRecDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rec date"
      Height          =   195
      Left            =   4560
      TabIndex        =   56
      Top             =   10110
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label LblRecCls 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rec Cls"
      Height          =   195
      Left            =   5850
      TabIndex        =   55
      Top             =   10110
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label LblWHCode 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WHCode"
      Height          =   195
      Left            =   7140
      TabIndex        =   54
      Top             =   10110
      Visible         =   0   'False
      Width           =   660
   End
   Begin MSForms.ComboBox cbocurr 
      Height          =   315
      Left            =   10380
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   8400
      Width           =   735
      VariousPropertyBits=   746604575
      BackColor       =   -2147483624
      MaxLength       =   2
      DisplayStyle    =   3
      Size            =   "1296;556"
      BoundColumn     =   2
      MatchEntry      =   1
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.ComboBox cboPrice 
      Height          =   315
      Left            =   11175
      TabIndex        =   13
      Top             =   8400
      Width           =   1635
      VariousPropertyBits=   612386843
      DisplayStyle    =   3
      Size            =   "2884;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Line Line6 
      Visible         =   0   'False
      X1              =   8055
      X2              =   9570
      Y1              =   10830
      Y2              =   10830
   End
   Begin VB.Label lblTransport 
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
      Left            =   13335
      TabIndex        =   53
      Top             =   9000
      Width           =   1530
   End
   Begin VB.Line Line5 
      X1              =   13320
      X2              =   14835
      Y1              =   9270
      Y2              =   9270
   End
   Begin VB.Label LblPart 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Package"
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
      Left            =   5070
      TabIndex        =   52
      Top             =   10590
      Visible         =   0   'False
      Width           =   720
   End
   Begin MSForms.ComboBox cboTransport 
      Height          =   315
      Left            =   12345
      TabIndex        =   17
      Top             =   8970
      Width           =   900
      VariousPropertyBits=   746604571
      MaxLength       =   2
      DisplayStyle    =   3
      Size            =   "1587;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label LblPart 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transport by"
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
      Left            =   11055
      TabIndex        =   51
      Top             =   9030
      Width           =   1095
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FDDFE3&
      Caption         =   "BC 40 No."
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
      Left            =   11745
      TabIndex        =   50
      Top             =   10590
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Part (Material) Receipt Scheduled"
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
      Left            =   345
      TabIndex        =   49
      Top             =   293
      Width           =   14595
   End
   Begin VB.Line Line3 
      X1              =   2235
      X2              =   5715
      Y1              =   8700
      Y2              =   8700
   End
   Begin VB.Line Line2 
      X1              =   11070
      X2              =   12120
      Y1              =   7740
      Y2              =   7740
   End
   Begin VB.Label LblRec 
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
      Left            =   11070
      TabIndex        =   48
      Top             =   7500
      Width           =   1065
   End
   Begin VB.Label Label10 
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
      Left            =   12240
      TabIndex        =   45
      Top             =   7530
      Width           =   1095
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Receipt Cls"
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
      Left            =   9360
      TabIndex        =   44
      Top             =   7530
      Width           =   960
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   7080
      X2              =   8220
      Y1              =   8670
      Y2              =   8670
   End
   Begin VB.Label LblAddress 
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
      Left            =   7080
      TabIndex        =   43
      Top             =   8430
      Visible         =   0   'False
      Width           =   1155
   End
   Begin MSForms.ComboBox CboWHCode 
      Height          =   315
      Left            =   5805
      TabIndex        =   9
      Top             =   8400
      Width           =   1200
      VariousPropertyBits=   612386843
      MaxLength       =   15
      DisplayStyle    =   3
      Size            =   "2117;556"
      ListRows        =   15
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox CboRecCls 
      Height          =   315
      Left            =   10380
      TabIndex        =   6
      Top             =   7440
      Width           =   645
      VariousPropertyBits=   746604571
      MaxLength       =   6
      DisplayStyle    =   3
      Size            =   "1138;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label LblPart 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WH Code"
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
      Left            =   6000
      TabIndex        =   42
      Top             =   7980
      Width           =   795
   End
   Begin VB.Label Label4 
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
      Left            =   7305
      TabIndex        =   41
      Top             =   7980
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
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
      Left            =   2235
      TabIndex        =   40
      Top             =   7980
      Width           =   960
   End
   Begin VB.Label LblPart 
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
      Index           =   5
      Left            =   405
      TabIndex        =   39
      Top             =   9030
      Width           =   765
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00A6D2FF&
      Height          =   585
      Index           =   2
      Left            =   345
      Top             =   8265
      Width           =   14595
   End
   Begin VB.Label Label12 
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
      Left            =   13275
      TabIndex        =   38
      Top             =   7980
      Width           =   1140
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Curr"
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
      Left            =   10545
      TabIndex        =   37
      Top             =   7980
      Width           =   390
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
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
      Left            =   11790
      TabIndex        =   36
      Top             =   7980
      Width           =   420
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qty"
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
      Left            =   8775
      TabIndex        =   35
      Top             =   7980
      Width           =   300
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code"
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
      Left            =   465
      TabIndex        =   34
      Top             =   7980
      Width           =   1155
   End
   Begin MSForms.ComboBox Cbounit 
      Height          =   315
      Left            =   9585
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   8400
      Width           =   735
      VariousPropertyBits=   746604575
      BackColor       =   -2147483624
      MaxLength       =   2
      DisplayStyle    =   3
      Size            =   "1296;556"
      BoundColumn     =   2
      MatchEntry      =   1
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unit"
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
      Left            =   9780
      TabIndex        =   33
      Top             =   7980
      Width           =   330
   End
   Begin VB.Label LblPart 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Srt Jln No"
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
      Left            =   330
      TabIndex        =   32
      Top             =   7530
      Width           =   825
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00A6D2FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   2
      Left            =   345
      Top             =   7905
      Width           =   14595
   End
End
Attribute VB_Name = "FrmPart_Rec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rst As Recordset, RsW As Recordset, RsI As Recordset
Dim SQLT As String, SqlW As String, SqlI As String
Dim RsPM As Recordset, RsPD As Recordset, RsPr As Recordset
Dim SqlPM As String, SqlPD As String
Dim baru As Boolean
Dim Pos As Integer, jml As Integer
Dim StaErr As Boolean, nErr As Integer, StrWDel As String, nOK As Integer
Dim HakU As Integer
Dim staBaru As String
Dim OrderQty As Double, QTYRec As Double
Dim SeqID As Long
Dim nPrice As String
Dim seqNo As New clsMRP
Dim DateActual, receiptDate As Date
Dim validate As Boolean


Dim KeyProd As String
Dim tampungQty As Double
Dim blnFix As Integer, thnFix As Integer
Dim newDb As New ADODB.Connection
Dim provisionCls As String, l_SJNo As String
Dim xcurr As String, blnshow As Boolean

Dim bteColSelect As Byte
Dim bteColProdCod As Byte
Dim bteColPartNo As Byte
Dim bteColDesc As Byte
Dim bteColDate As Byte
Dim bteColCls As Byte
Dim bteColWHCode As Byte
Dim bteColOrder As Byte
Dim bteColReceipt As Byte
Dim bteColRemain As Byte
Dim bteColUnit As Byte
Dim bteColCurr As Byte
Dim bteColPrice As Byte
Dim bteColAmount As Byte
Dim bteColPriceService As Byte
Dim bteColAmountService As Byte
Dim bteColItemCode As Byte
Dim bteColDelDate As Byte
Dim bteColAddress As Byte
Dim bteColRecCls As Byte
Dim bteColRecDate As Byte
Dim bteColRecUnit As Byte
Dim bteColRecCurr As Byte
Dim bteColRem As Byte
Dim bteColQtyOrder As Byte
Dim bteColQtyRec As Byte
Dim bteColRecPrice As Byte
Dim bteColRecWHCode As Byte
Dim bteColUnitCls As Byte
Dim bteColCurrCode As Byte
Dim bteColComplete As Byte
Dim bteColComplteCls As Byte
Dim bteColSeqNo As Byte
Dim bteColProvision As Byte
Dim bteColBC40 As Byte
Dim bteColTransport As Byte
Dim bteColPackageQty As Byte
Dim bteColPackage As Byte
Dim bteColBCDate As Byte
Dim bteColBctype As Byte
Dim bteColRecStatus As Byte
Dim bteColNoRegister As Byte
Dim bteColNoSeri As Byte

Dim bteHakPrice As Byte
Dim dblQty As Double, dblQtyLama As Double ' Add for Qty Receipt Validation

Sub Header()
    Dim C As Byte
    
    bteColSelect = 0
    bteColProdCod = 1
    bteColPartNo = 2
    bteColDesc = 3
    bteColDate = 4
    bteColWHCode = 5
    bteColCls = 6
    bteColOrder = 7
    bteColReceipt = 8
    bteColRemain = 9
    bteColUnit = 10
    bteColComplete = 11
    bteColCurr = 12
    bteColPrice = 13
    bteColAmount = 14
    bteColPriceService = 15
    bteColAmountService = 16
    bteColItemCode = 17
    bteColDelDate = 18
    bteColAddress = 19
    bteColRecCls = 20
    bteColRecDate = 21
    bteColRecUnit = 22
    bteColRecCurr = 23
    bteColRem = 24
    bteColQtyOrder = 25
    bteColQtyRec = 26
    bteColRecPrice = 27
    bteColRecWHCode = 28
    bteColUnitCls = 29
    bteColCurrCode = 30
    bteColBC40 = 31
    bteColComplteCls = 32
    bteColSeqNo = 33
    bteColProvision = 34
    bteColTransport = 35
    bteColPackageQty = 36
    bteColPackage = 37
    bteColBCDate = 38
    bteColBctype = 39
    bteColRecStatus = 40
    bteColNoRegister = 41
    bteColNoSeri = 42
    
    With grid
        .ColS = 43
        .Rows = 1
        
        .TextMatrix(0, bteColSelect) = ""
        .TextMatrix(0, bteColProdCod) = "Product Code"
        .TextMatrix(0, bteColPartNo) = "Part Number"
        .TextMatrix(0, bteColDesc) = "Description"
        .TextMatrix(0, bteColDate) = "Delivery Date"
        .TextMatrix(0, bteColCls) = "Cls"
        .TextMatrix(0, bteColWHCode) = "WH Code"
        .TextMatrix(0, bteColOrder) = "Order"
        .TextMatrix(0, bteColReceipt) = "Receipt"
        .TextMatrix(0, bteColRemain) = "Remaining"
        .TextMatrix(0, bteColUnit) = "Unit"
        .TextMatrix(0, bteColCurr) = "Curr"
        .TextMatrix(0, bteColPrice) = "Price"
        .TextMatrix(0, bteColAmount) = "Amount"
        .TextMatrix(0, bteColPriceService) = "Price Service"
        .TextMatrix(0, bteColAmountService) = "Amount Service"
        .TextMatrix(0, bteColItemCode) = "ItemCode"
        .TextMatrix(0, bteColDelDate) = "Deldate"
        .TextMatrix(0, bteColAddress) = "Address"
        .TextMatrix(0, bteColRecCls) = "RecCls"
        .TextMatrix(0, bteColRecDate) = "RecDate"
        .TextMatrix(0, bteColRecUnit) = "Unit"
        .TextMatrix(0, bteColRecCurr) = "Curr"
        .TextMatrix(0, bteColRem) = "Remarks"
        .TextMatrix(0, bteColQtyOrder) = "OrderQTY"
        .TextMatrix(0, bteColQtyRec) = "RecQTY"
        .TextMatrix(0, bteColRecPrice) = "Price"
        .TextMatrix(0, bteColRecWHCode) = "WH Code"
        .TextMatrix(0, bteColUnitCls) = "Unit Cls"
        .TextMatrix(0, bteColCurrCode) = "Curr Code"
        .TextMatrix(0, bteColComplete) = "Complete"
        .TextMatrix(0, bteColComplteCls) = "Complete"
        .TextMatrix(0, bteColSeqNo) = "Seqno"
        .TextMatrix(0, bteColProvision) = "Provision"
        .TextMatrix(0, bteColBC40) = "BC No."
        .TextMatrix(0, bteColBctype) = "BC Type"
        .TextMatrix(0, bteColRecStatus) = "Receipt Status"
        .TextMatrix(0, bteColNoRegister) = "No. Register"
        .TextMatrix(0, bteColNoSeri) = "No."
        
        .Cell(flexcpAlignment, 0, 0, 0, .ColS - 1) = flexAlignCenterCenter
        .ColAlignment(bteColProdCod) = flexAlignLeftCenter
        .ColAlignment(bteColPartNo) = flexAlignLeftCenter
        .ColAlignment(bteColDesc) = flexAlignLeftCenter
        .ColAlignment(bteColDate) = flexAlignLeftCenter
        .ColAlignment(bteColCls) = flexAlignCenterCenter
        .ColAlignment(bteColOrder) = flexAlignRightCenter
        .ColAlignment(bteColReceipt) = flexAlignRightCenter
        .ColAlignment(bteColRemain) = flexAlignRightCenter
        .ColAlignment(bteColUnit) = flexAlignCenterCenter
        .ColAlignment(bteColCurr) = flexAlignCenterCenter
        .ColAlignment(bteColPrice) = flexAlignRightCenter
        .ColAlignment(bteColAmount) = flexAlignRightCenter
        .ColAlignment(bteColPriceService) = flexAlignRightCenter
        .ColAlignment(bteColAmountService) = flexAlignRightCenter
        .ColAlignment(bteColRem) = flexAlignLeftCenter
        .ColAlignment(bteColBC40) = flexAlignLeftCenter
        .ColAlignment(bteColBctype) = flexAlignLeftCenter
        .ColAlignment(bteColRecStatus) = flexAlignLeftCenter
        .ColAlignment(bteColNoRegister) = flexAlignLeftCenter
         
        .ColWidth(bteColSelect) = 300
        .ColWidth(bteColProdCod) = 1500
        .ColWidth(bteColPartNo) = 1500
        .ColWidth(bteColDesc) = 2500
        .ColWidth(bteColDate) = 1400
        .ColWidth(bteColCls) = 500
        .ColWidth(bteColOrder) = 1200
        .ColWidth(bteColReceipt) = 1200
        .ColWidth(bteColRemain) = 1200
        .ColWidth(bteColUnit) = 500
        .ColWidth(bteColCurr) = 500
        .ColWidth(bteColPrice) = 1500
        .ColWidth(bteColAmount) = 1500
        .ColWidth(bteColPriceService) = 1500
        .ColWidth(bteColAmountService) = 1500
        .ColWidth(bteColComplete) = 1000
        .ColWidth(bteColRem) = 1500
        .ColWidth(bteColBC40) = 1500
        .ColWidth(bteColBctype) = 1500
        .ColWidth(bteColRecStatus) = 1500
        .ColWidth(bteColNoRegister) = 1700
        .ColWidth(bteColNoSeri) = 500
        
        .ColHidden(bteColItemCode) = True
        .ColHidden(bteColDelDate) = True
        .ColHidden(bteColAddress) = True
        .ColHidden(bteColRecCls) = True
        .ColHidden(bteColRecDate) = True
        .ColHidden(bteColRecUnit) = True
        .ColHidden(bteColRecCurr) = True
        .ColHidden(bteColQtyOrder) = True
        .ColHidden(bteColQtyRec) = True
        .ColHidden(bteColRecPrice) = True
        .ColHidden(bteColRecWHCode) = True
        .ColHidden(bteColUnitCls) = True
        .ColHidden(bteColCurrCode) = True
        .ColHidden(bteColComplteCls) = True
        .ColHidden(bteColSeqNo) = True
        .ColHidden(bteColProvision) = True
        .ColHidden(bteColTransport) = True
        .ColHidden(bteColPackageQty) = True
        .ColHidden(bteColPackage) = True
        .ColHidden(bteColBC40) = True
        .ColHidden(bteColBCDate) = True
        .ColHidden(bteColBctype) = True
        .ColHidden(bteColRecStatus) = True
        
        .ColHidden(bteColCurr) = (bteHakPrice = 0)
        .ColHidden(bteColPrice) = (bteHakPrice = 0)
        .ColHidden(bteColAmount) = (bteHakPrice = 0)
        .ColHidden(bteColPriceService) = (bteHakPrice = 0)
        .ColHidden(bteColAmountService) = (bteHakPrice = 0)
        
        .EditMaxLength = 1
    End With
End Sub

Private Sub cbocurr_Change()
    cbocurr_Click
    If cbocurr.ListCount > 0 Then
        cbocurr = Trim(cbocurr)
        If cbocurr.MatchFound Then
            xcurr = cbocurr.List(, 0)
            blnshow = False
            'Call browseprice(CboItem, TglReceipt, CboPart(0), xcurr)
        Else
            xcurr = ""
        End If
    End If
End Sub

Private Sub cbocurr_Click()
    CboPart(0) = CboPart(0)
    cbounit = cbounit
End Sub

Private Sub cbocurr_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub CboItem_Change()
    Dim RSX As Recordset
    Dim RsI2 As Recordset
    
    If Trim(cboitem) <> "" Then
        LblItemName = Trim$(uf_GetItemDescription(Trim$(cboitem)))
        Lbladdress = ""
        Set RsI2 = Db.Execute("Select WH_code,unit_cls from Item_master where item_code='" & Trim(cboitem) & "'")
        If Not RsI2.EOF Then
            cboWhCode = Trim$(RsI2!wh_code)
            CboWHCode_Click
            Set RSX = Db.Execute("Select unit_cls UC,Currency_code CC, Coalesce(Price,0) PC from Price_master where item_code='" & Trim$(cboitem) & "'")
            If Not RSX.EOF Then
                cbocurr.Text = uf_GetCurrencyDescription(Trim(RSX!CC))
                cbocurr.TextColumn = 2
                cbocurr_Click
                cboprice.Text = RSX!pc
            End If
            Call filterCboUnit(RsI2!Unit_cls, cbounit)
            cbounit.TextColumn = 2
            cbounit.Text = uf_GetUnitDescription(Trim(RsI2!Unit_cls))
        End If
        blnshow = False
       ' browseprice Trim$(CboItem), Format(TglReceipt, "YYYY-MM-DD"), Trim$(CboPart(0)), xcurr
    End If
    Exit Sub
X:
    Lbladdress = ""
End Sub

Private Sub GetSubConWarehouseCode()
Dim RS As New ADODB.Recordset
With cmbbox_warehouse
    .clear
    .columnCount = 3
    .ListWidth = "360"
    .ColumnWidths = "60pt;300pt;0pt"
    .Text = ""
    RS.Open "select wh_code, wh_name, stockcontrol_cls from warehouse_master where adm_group = '" & Trim(CboPart(0)) & "' ", Db, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not RS.EOF Then
        While Not RS.EOF
            .AddItem ""
            .Column(0, .ListCount - 1) = Trim(RS.Fields("wh_code"))
            .Column(1, .ListCount - 1) = Trim(RS.Fields("wh_name"))
            .Column(2, .ListCount - 1) = Trim(RS.Fields("stockcontrol_cls"))
            RS.MoveNext
        Wend
        .ListIndex = 0
    End If
    RS.Close
End With
End Sub

Private Function uf_GetSubConStatus(ls_TradeCode As String) As String
Dim RS As New ADODB.Recordset
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
RS.Open "select trade_cls from trade_master where trade_code='" & Trim(ls_TradeCode) & "'", Db, adOpenKeyset, adLockOptimistic
If RS.EOF = False Then
    uf_GetSubConStatus = Trim(RS!trade_cls & "")
Else
    uf_GetSubConStatus = ""
End If
If RS.State = 1 Then RS.Close
End Function

Private Sub cboPackage_Change()
    If cboPackage.MatchFound Then
        lblPackage = cboPackage.Column(1)
    Else
        lblPackage = ""
    End If
End Sub

Private Sub CboPart_Change(Index As Integer)
    If Index = 0 Then
        CboPart(0) = CboPart(0)
        lblSupp = ""
        CbotypeBC = ""
        Header
        GetSubConWarehouseCode
        GetBCType
        If CboPart(0).MatchFound = True Then lblSupp = CboPart(0).List(CboPart(0).ListIndex, 1)
        Tgl2_Click
    End If
    If Index = 1 Then
        CboPart(1) = CboPart(1)
        Header
        'If CboPart(1).MatchFound = True Then display
    End If
End Sub

Private Sub CboPart_Click(Index As Integer)
    If Index = 0 Then Tgl2_Click
    If Index = 1 Then
        CboPart(1) = CboPart(1)
        Header
        'If CboPart(1).MatchFound = True Then display
    End If
End Sub

Private Sub CboPart_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub cboprice_Change()
    If Trim(cboprice) = "" Then cboprice.Text = Format(0, gs_formatPrice): Exit Sub
    If CDbl(cboprice) = 0 Then txtamount = Format(0, gs_formatPrice): Exit Sub
    If Left(Trim(cboprice), 1) = "," Then _
        cboprice.Text = Right(Trim(cboprice), Len(Trim(cboprice)) - 1)
        If Trim(cboprice) <> "" And IsNumeric(cboprice) = True Then
            If Trim(txtQty) <> "" And IsNumeric(txtQty) = True Then
                txtamount.Text = Format(CDbl(cboprice) * CDbl(txtQty.Text), gs_formatAmount)
            Else
                txtamount = Format(0, gs_formatAmount)
            End If
        Else
            txtamount = Format(0, gs_formatAmount)
        End If
End Sub

Private Sub cboprice_Click()
    If CDbl(cboprice) = 0 Then txtamount = Format(0, gs_formatAmount): Exit Sub
    If Trim(txtQty) = "" Then txtamount = Format(0, gs_formatAmount): Exit Sub
    If CDbl(txtQty) > 0 And CDbl(cboprice) > 0 Then
        txtamount = CDbl(txtQty) * CDbl(cboprice)
        If Round(CDbl(txtamount)) / CDbl(txtamount) = 1 Then
            txtamount = Format(CDbl(txtamount), gs_formatAmount)
        Else
            txtamount = Format(CDbl(txtamount), gs_formatAmount)
        End If
    Else
        txtamount = Format(0, gs_formatAmount)
    End If
End Sub

Private Sub CboPrice_KeyPress(KeyAscii As MSForms.ReturnInteger)
 If InStr(1, "0123456789.", Chr(KeyAscii)) <= 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub cboPrice_LostFocus()
    If IsNumeric(cboprice) = True Then
        cboprice.Text = Format(cboprice, gs_formatPrice)
    Else
        cboprice.Text = Format(0, gs_formatPrice)
    End If
End Sub

Private Sub CboRecCls_Change()
    CboRecCls_Click
End Sub

Private Sub CboRecCls_Click()
    CboRecCls = CboRecCls
    If CboRecCls.MatchFound Then LblRec = CboRecCls.List(CboRecCls.ListIndex, 1)
End Sub

Private Sub CboRecCls_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub cboTransport_Change()
    If cboTransport.MatchFound Then
        lblTransport = cboTransport.Column(1)
    Else
        lblTransport = ""
    End If
End Sub
Private Sub settingcombo()
Dim ls_sql As String
Dim rs_combo As New ADODB.Recordset
Dim i As Integer


CbotypeBC.columnCount = 1
CbotypeBC.clear

ls_sql = "select BC_Type from BC_master"
rs_combo.Open ls_sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
i = 0

Do While Not rs_combo.EOF
CbotypeBC.AddItem rs_combo("BC_Type")
rs_combo.MoveNext
Loop

CbotypeBC.ColumnWidths = "90"
CbotypeBC.ListWidth = 90
CbotypeBC.ListRows = 7


End Sub

Private Sub CbotypeBC_Change()
    
    LblErr.Caption = ""
    
    up_GetNoSeri (cboitem.Text)
    
    If CbotypeBC.Text = "4.0" Or CbotypeBC.Text = "2.6.2" Or CbotypeBC.Text = "2.3" Then
        txtRec_Status.Text = "02"
    Else
        txtRec_Status.Text = "01"
    End If
    
    TxtSj_LostFocus
    
    
End Sub

Private Sub Cbounit_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub CboWHCode_Change()
    CboWHCode_Click
End Sub

Private Sub CboWHCode_Click()
    Dim RsIA As Recordset
    cboWhCode = cboWhCode
    cboitem = cboitem
    Lbladdress = ""
    If Trim$(cboWhCode) <> "" Then
        Set RsIA = Db.Execute("Select Address from item_master where item_code='" & cboitem & "' and WH_code='" & cboWhCode & "'")
      '  If Not RsIA.EOF Then Lbladdress = IIf(IsNull(Trim$(RsIA!Address)), "", Trim$(RsIA!Address))
    End If
End Sub

Private Sub CboWHCode_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub cmbbox_warehouse_Change()
    If cmbbox_warehouse.MatchFound Then
        txtWHSubcon = cmbbox_warehouse.Column(1, cmbbox_warehouse.ListIndex)
    Else
        txtWHSubcon = ""
    End If
End Sub

Private Sub cmd_Browser_Click()

  
  Me.MousePointer = vbHourglass
  frm_BrowseSupp.getItemCode = CboPart(0).Text
  frm_BrowseSupp.Show 1
  CboPart(0).Text = frm_BrowseSupp.getPartNumber
  CboPart(0).SetFocus
  
  Me.MousePointer = vbDefault
End Sub

Private Sub CmdBC40_Click()
    
    Me.MousePointer = vbHourglass
    
    If grid.Rows > 1 And grid.TextMatrix(grid.Row, bteColBC40) <> "" Then
        BC40 grid.TextMatrix(grid.Row, bteColBC40)
    Else
        LblErr = DisplayMsg("0013")
    End If
    
    
    Me.MousePointer = vbDefault
    
End Sub

Private Sub cmdCancel_Click()
    LblErr = ""
    txtSearch = ""
    display grid.Row
    Kosong 1
End Sub

Private Sub CmdCari_Click()
    If CboPart(1).MatchFound = True Then display
End Sub

Private Sub cmdClear_Click()
    Header
    cboitem = ""
    cmbbox_warehouse = ""
    cboWhCode = ""
    LblItemName = ""
    LblRec = ""
    LblRecCls = ""
    lblWHCode = ""
    lblItemCode = ""
    lblQty = ""
    LblRecDate = ""
    Lbladdress = ""
    txtQty = ""
    cbounit = ""
    cbocurr = ""
    cboprice.Text = ""
    txtamount = ""
    TxtSj = ""
    CboRecCls = ""
    TglReceipt = Now()
    DtBCDate = Now()
    txtremarks = ""
    CboPart(0) = ""
    Tgl1 = Now()
    Tgl2 = Now()
    CboPart(1) = ""
    txtBCNo = ""
    cboTransport = ""
    cboPackage = ""
    txtPackage = ""
    LblErr = ""
    txtSearch = ""
    CbotypeBC.Text = ""
    txtRegisterNo.Text = ""
    txtNoSeri.Text = ""
End Sub

Private Sub CmdMenu_Click()
    rst.Close
    Set rst = Nothing
    RsW.Close
    Set RsW = Nothing
    frmMainMenu.Show
    Unload Me
    DoEvents
End Sub

Private Sub cmdSearch_Click()
    Dim i As Double
    On Error Resume Next
    
    LblErr = ""
    
    If txtSearch = "" Or grid.Rows = 1 Then txtSearch.SetFocus: Exit Sub
    'If Grid.Row = Grid.Rows - 1 Then i = 1 Else i = Grid.Row + 1
    
    Do
'        Select Case cboSearch.ListIndex
'        Case 0
            grid.Col = bteColProdCod
            If UCase(Mid(grid.TextMatrix(i, bteColProdCod), 1, Len(txtSearch))) = UCase(txtSearch) Then
                Exit Do
            End If
'        Case 1
'            Grid.Col = bteColDesc
'            If InStr(UCase(Grid.TextMatrix(i, bteColDesc)), UCase(txtSearch)) <> 0 Then
'                Exit Do
'            End If
'        End Select
        i = i + 1
        If i = grid.Rows - 1 Then
            txtSearch = ""
            i = 2
            LblErr = DisplayMsg(8012)
            Exit Do
        End If
    Loop
    
    grid.Row = i
    grid.TopRow = i
    grid.SetFocus

End Sub

Private Sub cmdsj_Click()

With frmupdate

    .txtsupplier.Text = CboPart(0).Text
    .txtpo.Text = CboPart(1).Text
    
    .txtsjno.Text = TxtSj.Text
    .lblSJNo.Text = TxtSj.Text
    .lblsjdate.Text = TxtSj.Text
    .sjdate.Value = TglReceipt.Value
    .lblsjdate.Text = Format(TglReceipt.Value, "dd MMM yyyy")
    
    .Show 1

End With

End Sub

Private Sub CmdSubmit_Click()

    Dim strS As Integer, Jawab As Integer, CekS As Boolean, CekD As Boolean
    Dim strD As Integer, ie As Integer
    Dim CC As Boolean, ubah As Integer
    Dim totalQty As Double
    Dim strSQL As String
    
    Dim FixMonth As String * 2
    Dim FixYear As String * 4
    Dim rsProv As New ADODB.Recordset
    
    Dim icg As Long
    Dim tampungBln As String
    
    ' Add Validation For Over Qty -- KAWAI 20090527
    If baru = False Then
        totalQty = dblQty + dblQtyLama
    Else
        totalQty = dblQty
    End If
    
    If txtQty = "" Then
    LblErr = DisplayMsg(1068)
    Exit Sub
    End If
    
    If totalQty < CDbl(txtQty) Then
        LblErr = " [0000] - Invalid Quantity Input (Over from Order)"
        Exit Sub
    End If
    
        
    
    ' ---------------
    Me.MousePointer = vbHourglass
    LblErr = ""
    
    CekS = False
    CekD = False
    StaErr = False
    ubah = 0

    If HakU = 0 Then LblErr = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
    
    'If up_Validasi_Item(Trim(CboItem.Text)) = False Then up_SendEmail
    
    '#Check Warehouse Subcon not blank when subcon
    If uf_GetSubConStatus(CboPart(0)) = "3" And Trim(cmbbox_warehouse) = "" Then
        LblErr = DisplayMsg(8100): Me.MousePointer = vbDefault: Exit Sub
    End If
    
    tampungBln = seqNo.blnAkhir()
    blnFix = Split(tampungBln, ",")(0)
    thnFix = Split(tampungBln, ",")(1)
    
    Db.BeginTrans
    For icg = 1 To grid.Rows - 1
        If Trim$(grid.Cell(flexcpChecked, icg, bteColProdCod)) <> "" Or Trim$(grid.Cell(flexcpChecked, icg, bteColPartNo)) <> "" Then
         If grid.Cell(flexcpChecked, icg, bteColComplete) <> Val(grid.TextMatrix(icg, bteColComplteCls)) Then
            If grid.Cell(flexcpChecked, icg, bteColComplete) = flexChecked Then
                Db.Execute "Update Purchaseorder_detail with (updlock) set Complete_cls='1', Last_Update = getdate(), Last_User = '" & userLogin & "' " & _
                    "where PO_no='" & Trim$(CboPart(1)) & "' and item_code='" & Trim$(grid.TextMatrix(icg, bteColProdCod)) & "'"
            ElseIf grid.Cell(flexcpChecked, icg, bteColComplete) = flexUnchecked Then
                Db.Execute "Update Purchaseorder_detail with (updlock) set Complete_cls='0', Last_Update = getdate(), Last_User = '" & userLogin & "' " & _
                    "where PO_no='" & Trim$(CboPart(1)) & "' and item_code='" & Trim$(grid.TextMatrix(icg, bteColProdCod)) & "'"
            End If
            ubah = ubah + 1
         End If
        End If
    Next
    
    If err.number = 0 Then
        Db.CommitTrans
    Else
        Db.RollbackTrans
        err.clear
    End If
    
    strS = grid.FindRow("S", , bteColSelect, False)
    strD = grid.FindRow("D", , bteColSelect, False, False)
    
    If baru = False Then
        
        '# Yudha 2007-12-11
        '# Cek AP Invoice, jika sudah dibuat tidak bisa di-update
        Dim adoInv As New ADODB.Recordset
        sql = "select invoice_no from invoicesupplier_detail where receiptseq_no = '" & SeqID & "'"
        adoInv.Open sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
        If Not adoInv.EOF Then
            LblErr = DisplayMsg(4110) & " (Invoice No : " & Trim(adoInv.Fields("invoice_no")) & ")": Me.MousePointer = vbDefault: adoInv.Close: Exit Sub
        End If
        adoInv.Close
        
        If strD > 0 Then
            CekD = True
            Dim ipo As Long, Pos As Long, jum As Long
            jum = 0
            For ipo = 1 To grid.Rows - 1
                Pos = grid.FindRow("D", ipo, bteColSelect, False, False)
                If Pos > 0 Then jum = jum + 1
            Next
            If jum = 1 Then
                Pos = grid.FindRow("D", , bteColSelect, False, False)
                If grid.TextMatrix(Pos, bteColDelDate) = "" Then grid.TextMatrix(Pos, bteColSelect) = "": LblErr = DisplayMsg(1202): Me.MousePointer = vbDefault: Exit Sub
            End If
            Jawab = MsgBox("Do you really want to Delete this Record", vbInformation + vbYesNo + vbDefaultButton2, "Confirmation")
        End If
        If Jawab = vbYes Then DataGrid
        
        If strS > 0 And cek Then
                        
            Dim li_diff As Integer
            li_diff = DateDiff("M", uf_GetLastClosing("fulldate"), TglReceipt)
            If li_diff > 2 Or li_diff < 0 Then LblErr = DisplayMsg(8022): Me.MousePointer = vbDefault: Exit Sub
            
            DataGrid
            
            If StaErr = False Then
                                
                Db.BeginTrans
'                If (QTYRec - CDbl(LblQty)) + CDbl(txtqty) >= OrderQty Then
'                    Db.Execute "Update Purchaseorder_detail with (updlock) set Complete_cls='1', Last_Update = getdate(), Last_User = '" & userLogin & "' " & _
'                        "where PO_no='" & Trim$(CboPart(1)) & "' and item_code='" & Trim$(cboitem) & "'"
'                Else
'                    Db.Execute "Update Purchaseorder_detail with (updlock) set Complete_cls='0', Last_Update = getdate(), Last_User = '" & userLogin & "' " & _
'                        "where PO_no='" & Trim$(CboPart(1)) & "' and item_code='" & Trim$(cboitem) & "'"
'                End If
                                
                Db.Execute "update purchaseorder_detail with (updlock) set complete_cls = " & _
                    "case when " & _
                        "(select isnull(sum(qty), 0) from part_receipt where receipt_cls = 'R' and po_no = purchaseorder_detail.po_no and item_code = purchaseorder_detail.item_code) - " & _
                        "(select isnull(sum(qty), 0) from part_receipt where receipt_cls = 'R1' and po_no = purchaseorder_detail.po_no and item_code = purchaseorder_detail.item_code) >= qty " & _
                    "then 1 else 0 end " & _
                    "where po_no='" & Trim$(CboPart(1)) & "' and item_code='" & Trim$(cboitem) & "' "
                
                If err.number = 0 Then
                    Db.CommitTrans
                Else
                    Db.RollbackTrans
                    err.clear
                End If
                
                baru = True
                display grid.Row
                Kosong 1
                cboitem = ""
                txtQty = Format(0, gs_formatQty)
                cbounit = ""
                cbocurr = ""
                cboprice.Text = Format(0, gs_formatPrice)
                txtamount = Format(0, gs_formatAmount)
                txtBC40 = ""
                TxtSj = ""
                CbotypeBC.Text = ""
                txtBCNo = ""
                txtRegisterNo = ""
                
                Dim IK As Long
                For IK = 1 To grid.Rows - 1
                    grid.TextMatrix(IK, bteColSelect) = ""
                Next
            
            Else
            
                LblErr = DisplayMsg(1102)
            
            End If
        
        Else
            
            If ubah > 0 Then LblErr = DisplayMsg(1101)
           
            
        
        End If
        
        Dim PRec As Integer
        If CekD Then
            If Jawab = vbYes Then
                If StaErr = False Then
                    LblErr = DisplayMsg(1201)
                    baru = True
                Else
                    LblErr = DisplayMsg(1202)
                End If
            ElseIf Jawab = vbNo Then
                LblErr = ""
                baru = True
                Kosong 1
                Dim Ikd As Long
                For Ikd = 1 To grid.Rows - 1
                    grid.TextMatrix(Ikd, bteColSelect) = ""
                Next
            End If
        End If
        
    Else
        
        Dim SqlU As String, PosRec As Integer
        
'        If txtNoSeri.Text = "" Then LblErr = DisplayMsg("0001") & " No Seri ! ":  txtNoSeri.SetFocus: Exit Sub
        
        If cek Then
            
            KeyProd = seqNo.keyReceipt
            SqlU = "insert into Part_Receipt (Seq_No, Supplier_Code, PO_No, Warehouse_Code, Address, Receipt_Cls, Receipt_Date, " & _
                "Item_Code, Qty, Unit_Cls, Currency_Code, Price, Amount, SuratJalan_No, ProductionResult_Cls, DailySeq_No, Remarks, " & _
                "Last_Update, Last_User,BC40_No, BC40_Date, BC_Type,Transport_Cls, Package_Qty, Package_Cls, Receipt_Status, No_Register, no_seri) " & _
                "Values (" & KeyProd & ",'" & Trim$(CboPart(0)) & "','" & Trim$(CboPart(1)) & "','" & Trim$(cboWhCode) & "','" & Trim$(Lbladdress) & "','" & Trim$(CboRecCls) & " ','" & Format(TglReceipt.Value, "yyyy-mm-dd") & "','" & _
                Trim$(cboitem) & "'," & CDbl(txtQty) & ",'" & Trim$(cbounit.List(cbounit.ListIndex, 0)) & "','" & Trim$(cbocurr.List(cbocurr.ListIndex, 0)) & "'," & _
                CDbl(cboprice) & "," & Round(CDbl(txtamount), gi_decimalDigitAmount) & ",'" & _
                Trim$(TxtSj) & "','0',null,'" & Trim$(txtremarks) & "', getdate(), '" & userLogin & "', '" & Trim(txtBCNo) & "','" & Format(DtBCDate.Value, "yyyy-mm-dd") & "', '" & Trim(CbotypeBC.Text) & "',  '" & Trim(cboTransport) & "', " & CDbl(txtPackage) & ", '" & Trim(cboPackage) & "' , '" & Trim(txtRec_Status.Text) & "', '" & Trim(txtRegisterNo.Text) & "', '" & Trim(txtNoSeri.Text) & "')"

            Dim rsc As Recordset
            Db.BeginTrans
            
            'On Error GoTo Errhandle
            Db.Execute SqlU
                
Errhandle:
            If InStr(1, err.Description, "Violation of PRIMARY KEY constraint") > 0 Then
                err.clear
                        KeyProd = seqNo.keyReceipt
                        SqlU = "insert into Part_Receipt (Seq_No, Supplier_Code, PO_No, Warehouse_Code, Address, Receipt_Cls, Receipt_Date, " & _
                            "Item_Code, Qty, Unit_Cls, Currency_Code, Price, Amount, SuratJalan_No, ProductionResult_Cls, DailySeq_No, Remarks, " & _
                            "Last_Update, Last_User,BC40_No, BC40_Date, BC_type, Transport_Cls, Package_Qty, Package_Cls, Receipt_Status, No_Register, no_seri) " & _
                            "Values (" & KeyProd & ",'" & Trim$(CboPart(0)) & "','" & Trim$(CboPart(1)) & "','" & Trim$(cboWhCode) & "','" & Trim$(Lbladdress) & "','" & Trim$(CboRecCls) & " ','" & Format(TglReceipt.Value, "yyyy-mm-dd") & "','" & _
                            Trim$(cboitem) & "'," & CDbl(txtQty) & ",'" & Trim$(cbounit.List(cbounit.ListIndex, 0)) & "','" & Trim$(cbocurr.List(cbocurr.ListIndex, 0)) & "'," & _
                            CDbl(cboprice) & "," & Round(CDbl(txtamount), gi_decimalDigitAmount) & "'," & _
                            Trim$(TxtSj) & "','0',null,'" & Trim$(txtremarks) & "'," & _
                            "getdate(), '" & userLogin & "', '" & Trim(txtBCNo) & "', '" & Format(DtBCDate.Value, "yyyy-mm-dd") & "','" & Trim(CbotypeBC.Text) & "' ,'" & Trim(cboTransport) & "', " & CDbl(txtPackage) & ", '" & Trim(cboPackage) & "', '" & Trim(txtRec_Status.Text) & "', '" & Trim(txtRegisterNo.Text) & "', '" & Trim(txtNoSeri.Text) & "')"
                        Db.Execute SqlU
                        If InStr(1, err.Description, "Violation of PRIMARY KEY constraint") > 0 Then GoTo Errhandle
            End If
            
            If err.number = 0 Then
                
                '*********** Proses Insert ke Supply **************
                sql = "Select Provision_Cls from Item_MaSter where ITem_Code = '" & cboitem & "'"
                Set rsProv = Db.Execute(sql)
                If Not rsProv.EOF Then provisionCls = Trim(rsProv(0)) Else provisionCls = ""
    
'                If provisionCls = "01" Then
'                    KeyProd = "R" & KeyProd
'                    Call inputSupply(CboItem, CDbl(txtQty))
'                End If
                '*************************************************
                
                '#Herfin 20070606
                '#Process Konsumsi Subcon
                '===============================================
                If uf_GetSubConStatus(Trim(CboPart(0))) = "3" Then
                    KeyProd = "R" & KeyProd
                    Call up_SubConInputConsumption(cboitem, CDbl(txtQty), Trim(cmbbox_warehouse))
                End If
                '===============================================
                
                If txtRec_Status.Text = "01" Then
                    ProsesStock 1, Trim$(cboitem), Trim$(cboWhCode), Trim$(cboWhCode), Format(TglReceipt, "YYYYMM"), Trim$(txtQty), ""
                End If
                
                Db.CommitTrans
            
            Else
                
                Db.RollbackTrans
                LblErr = err.Description
                err.clear
                Me.MousePointer = vbDefault
                Exit Sub
            
            End If
                
            Db.BeginTrans
'            If CDbl(txtqty) + QTYRec >= OrderQty Then
'                Db.Execute "Update Purchaseorder_detail with (updlock) set Complete_cls='1', Last_Update = getdate(), Last_User = '" & userLogin & "' " & _
'                    "where PO_no='" & Trim$(CboPart(1)) & "' and item_code='" & Trim$(CboItem) & "'"
'            End If
            
            Db.Execute "update purchaseorder_detail with (updlock) set complete_cls = " & _
                "case when " & _
                    "(select isnull(sum(qty), 0) from part_receipt where receipt_cls = 'R' and po_no = purchaseorder_detail.po_no and item_code = purchaseorder_detail.item_code) - " & _
                    "(select isnull(sum(qty), 0) from part_receipt where receipt_cls = 'R1' and po_no = purchaseorder_detail.po_no and item_code = purchaseorder_detail.item_code) >= qty " & _
                "then 1 else 0 end " & _
                "where po_no='" & Trim$(CboPart(1)) & "' and item_code='" & Trim$(cboitem) & "' "

            If err.number = 0 Then
                LblErr = DisplayMsg(1000)
                Db.CommitTrans
            Else
                LblErr = err.Description
                Db.RollbackTrans
                err.clear
            End If
            
            Kosong 1
            CboPart(0).SetFocus
        
            baru = True
            display grid.Row
            SqlU = ""
        
        Else
            
            If ubah > 0 Then LblErr = DisplayMsg(1101)
        
        End If
            
    End If
    
    Me.MousePointer = vbDefault
End Sub

Function isiPrice(ItemCode As String, tglDO As String, currCode As String) As String
    Dim rsPrice As New ADODB.Recordset
    sql = "select top 1 currency_code,isnull(price,0) Price from price_master where " & _
        "item_code='" & ItemCode & _
        "' and price_cls='01' " & _
        "and start_date<='" & Format(tglDO, "yyyymmdd") & _
        "' and end_date>='" & Format(tglDO, "yyyymmdd") & _
        "' order by trade_code desc, priority_cls desc"
        
    Set rsPrice = newDb.Execute(sql)
    If rsPrice.EOF Then
        isiPrice = currCode & ",0"
    Else
        isiPrice = Trim(rsPrice(0)) & "," & Trim(rsPrice(1))
    End If
    Set rsPrice = Nothing
End Function

Sub up_SubConInputConsumption(ibu As String, Qty As Double, ls_WhCode As String)

Dim rsAnak As New ADODB.Recordset, rsc As New ADODB.Recordset
Dim fromWHCode As String, fromAddress As String, toWHCode As String, toAddress As String
Dim UnitCls As String, currCD As String, Price As Double, Amount As Double
Dim itemAnak As String, nilPrice As String, currAnak As String, qtyAnak As Double
Dim stockWH As String, stockItem As String, Sn As Double

    '*********Update Supply Anak2nya diambil dr BOM MaSter ***********
    sql = "Select c.Manufacture_Code as Factory_Code,a.Item_Code,a.Qty as qtyAnak,a.Unit_Cls," & _
        "b.WH_Code,b.Address,b.Stockcontrol_Cls as stockItem,(select Stockcontrol_Cls from warehouse_master where wh_code='" & ls_WhCode & "') as stockWH, b.Provision_Cls " & _
        "from BOM_Master a,Item_Master b, Item_Master c " & _
        "where a.Item_Code = b.Item_Code " & _
        "And a.Parent_ItemCode = c.Item_Code " & _
        "  " & _
        "And a.Parent_ItemCode = '" & ibu & _
        "' And Start_Date <='" & Format(TglReceipt, "yyyyMMdd") & _
        "' And End_Date >= '" & Format(TglReceipt, "yyyyMMdd") & "' order by a.Item_Code"
    Set rsAnak = newDb.Execute(sql)

    If Not rsAnak.EOF Then
        Do While Not rsAnak.EOF
            fromWHCode = rsAnak("Wh_Code")
            fromAddress = IIf(IsNull(rsAnak("Address")), "", rsAnak("Address"))
            toWHCode = IIf(IsNull(rsAnak("Factory_Code")), "", rsAnak("Factory_Code"))
            itemAnak = rsAnak("Item_Code")
            qtyAnak = rsAnak("QtyAnak") * CDbl(Qty)
            UnitCls = rsAnak("Unit_Cls")
            nilPrice = isiPrice(itemAnak, Format(TglReceipt, "yyyy-MM-dd"), currCD)
            currAnak = Split(nilPrice, ",")(0)
            Price = Split(nilPrice, ",")(1)
            Amount = CDbl(qtyAnak) * CDbl(Price)
            stockWH = rsAnak("StockWH")
            stockItem = rsAnak("StockItem")
            
            Set rsc = Db.Execute("Select isnull(Max(seq_No),0)+1 From Part_Supply")
            Sn = rsc(0)
            rsc.Close
            
            sql = "insert into Part_Supply(FromWarehouse_Code,From_Address,ToWarehouse_Code,ChildSupply_date,ChildItem_Code,Supply_Cls," & _
                "ChildRequirement_Qty,ChildUnit_Cls,Currency_Code,Price,Amount,ParentItem_Code,Lot_No,Remarks,SubConPartReceipt_SeqNo,Do_NO," & _
                "Last_Update,Last_User) " & _
                "values ('" & ls_WhCode & "','" & "" & "','" & ls_WhCode & "','" & Format(TglReceipt, "yyyy-MM-dd") & "','" & itemAnak & "','S'," & _
                CDbl(qtyAnak) & ",'" & UnitCls & "','" & currAnak & "'," & Price & "," & Amount & ",'" & cboitem & "','','" & txtremarks & "','" & KeyProd & "', ''," & _
                "getdate(),'" & userLogin & "')"
            Db.Execute sql
            
            If stockWH = "01" And stockItem = "01" Then _
                Call seqNo.updateStock(ls_WhCode, itemAnak, qtyAnak, "", Format(TglReceipt, "yyyy-MM-dd"), blnFix, thnFix, Db, "Supply", 0, 1)
            If Not (rsAnak.EOF) Then rsAnak.MoveNext
        '******************
        Loop
        '******************
End If
End Sub

Sub inputSupply(ibu As String, Qty As Double, Optional UnitAnak As String, Optional TglRequirement As String, Optional tampungParent As String)

Dim rsAnak As New ADODB.Recordset, rsc As New ADODB.Recordset
Dim fromWHCode As String, fromAddress As String, toWHCode As String, toAddress As String
Dim UnitCls As String, currCD As String, Price As Double, Amount As Double
Dim itemAnak As String, nilPrice As String, currAnak As String, qtyAnak As Double
Dim stockWH As String, stockItem As String, Sn As Double

    '*********Update Supply Anak2nya diambil dr BOM MaSter ***********
    sql = "Select c.Manufacture_Code as Factory_Code,a.Item_Code,a.Qty as qtyAnak,a.Unit_Cls," & _
        "b.WH_Code,b.Address,b.Stockcontrol_Cls as stockItem,d.Stockcontrol_Cls as stockWH, b.Provision_Cls " & _
        "from BOM_Master a,Item_Master b, Item_Master c, Warehouse_Master d " & _
        "where a.Item_Code = b.Item_Code " & _
        "And a.Parent_ItemCode = c.Item_Code " & _
        "And b.WH_Code = d.WH_Code " & _
        "And a.Parent_ItemCode = '" & ibu & _
        "' And Start_Date <='" & Format(TglReceipt, "yyyyMMdd") & _
        "' And End_Date >= '" & Format(TglReceipt, "yyyyMMdd") & "' order by a.Item_Code"
    Set rsAnak = newDb.Execute(sql)

    If Not rsAnak.EOF Then
        Do While Not rsAnak.EOF
            
            fromWHCode = rsAnak("Wh_Code")
            fromAddress = IIf(IsNull(rsAnak("Address")), "", rsAnak("Address"))
            toWHCode = IIf(IsNull(rsAnak("Factory_Code")), "", rsAnak("Factory_Code"))
            itemAnak = rsAnak("Item_Code")
            qtyAnak = rsAnak("QtyAnak") * CDbl(Qty)
            UnitCls = rsAnak("Unit_Cls")
            nilPrice = isiPrice(itemAnak, Format(TglReceipt, "yyyy-MM-dd"), currCD)
            currAnak = Split(nilPrice, ",")(0)
            Price = Split(nilPrice, ",")(1)
            Amount = CDbl(qtyAnak) * CDbl(Price)
            stockWH = rsAnak("StockWH")
            stockItem = rsAnak("StockItem")
            
            Set rsc = Db.Execute("Select isnull(Max(seq_No),0)+1 From Part_Supply")
            Sn = rsc(0)
            rsc.Close
            
            sql = "insert into Part_Supply(FromWarehouse_Code,From_Address,ToWarehouse_Code,ChildSupply_date,ChildItem_Code,Supply_Cls," & _
                "ChildRequirement_Qty,ChildUnit_Cls,Currency_Code,Price,Amount,ParentItem_Code,Lot_No,Production_Date,Remarks,Do_NO," & _
                "Last_Update,Last_User) " & _
                "values ('" & fromWHCode & "','" & fromAddress & "','" & toWHCode & "','" & Format(TglReceipt, "yyyy-MM-dd") & "','" & itemAnak & "','S'," & _
                CDbl(qtyAnak) & ",'" & UnitCls & "','" & currAnak & "'," & Price & "," & Amount & ",'" & cboitem & "','',Null,'" & txtremarks & "','" & KeyProd & "'," & _
                "getdate(),'" & userLogin & "')"
            Db.Execute sql
            
            If stockWH = "01" And stockItem = "01" Then _
                Call seqNo.updateStock(fromWHCode, itemAnak, qtyAnak, "", Format(TglReceipt, "yyyy-MM-dd"), blnFix, thnFix, Db, "Supply", 0, 1)
            If Not (rsAnak.EOF) Then rsAnak.MoveNext
        '******************
        Loop
        '******************
    End If
End Sub


Private Sub Command1_Click()
MsgBox dblQty
MsgBox dblQtyLama
End Sub

Private Sub ComboBox1_Change()

End Sub

Private Sub CmdupdateBC_Click()
        Dim strSQL As String
    strSQL = " update Part_Receipt " & vbCrLf & _
                  " set BC40_No='" & Trim(txtBCNo.Text) & "',BC_Type='" & Trim(CbotypeBC.Text) & "',BC40_Date='" & Format(DtBCDate.Value, "yyyy-mm-dd") & "',Last_Update=getdate(),Last_User='" & Trim(userLogin) & "' " & vbCrLf & _
                  " Where Supplier_Code='" & Trim(CboPart(0).Text) & "' and SuratJalan_No='" & Trim(TxtSj.Text) & "' " & vbCrLf & _
                  "  "
    Db.Execute strSQL
    
    LblErr.Caption = "Update BC Nomor Success!!"
    
End Sub

Private Sub cmduploadbc_Click(Index As Integer)
 'Case 6:
    
'        If cbocust.Text = "" Then
'
'            lblErrMsg.Caption = "Please Select Customer Code !"
'            Exit Sub
'
'        ElseIf cbopono.Text = "" Then
'
'            lblErrMsg.Caption = "Please Select or Create SI/PO No !"
'            Exit Sub
'
'        ElseIf uf_validasi_upload(Trim(cbocust.Text), Trim(cbopono.Text)) = False Then
'            lblErrMsg.Caption = "Data SI/PO No : " & Trim(cbopono.Text) & " Has Been Upload !!"
'            Exit Sub
'
'        End If
'
'        FrmUploadDetailItem.txtTradeCode.Text = Trim(cbocust.Text)
'        FrmUploadDetailItem.txtpono.Text = Trim(cbopono.Text)
        FrmUploadBC.Show 1
       
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        cmdSearch_Click
    End If
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
Dim ir As Long, strUnit As String
CtrlMenu1.FormName = Me.Name
Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
HakU = hakUpdate(Me.Name)

If newDb.State <> adStateClosed Then newDb.Close
newDb.Open Db.ConnectionString

bteHakPrice = (hakPrice(Me.Name))
cbocurr.Visible = (bteHakPrice = 1)
cboprice.Visible = (bteHakPrice = 1)
txtamount.Visible = (bteHakPrice = 1)
Label7.Visible = (bteHakPrice = 1)
Label11.Visible = (bteHakPrice = 1)
Label12.Visible = (bteHakPrice = 1)

LblRec = ""
baru = True
Header
staBaru = ""

Call settingcombo

'##Tampilkan Combo Customer code dari trade_master
SQLT = "Select rtrim(trade_code) as TC,Trade_name as TN, Address1 as A from trade_master where trade_cls in ('2', '3') order by trade_code"

Set rst = New Recordset
rst.Open SQLT, Db, adOpenKeyset, adLockOptimistic
CboPart(0).clear
CboPart(0).columnCount = 2
CboPart(0).TextColumn = 1
ir = 0
While Not rst.EOF
    CboPart(0).AddItem ""
    CboPart(0).List(ir, 0) = rst!TC
    CboPart(0).List(ir, 1) = Trim$(rst!TN)
    ir = ir + 1
    rst.MoveNext
Wend
CboPart(0).ColumnWidths = "60 pt; 300 pt"
CboPart(0).ListWidth = 360
CboPart(0).ListRows = 15

'##Tampilkan Warehouse code dari warehouse_master
SqlW = "Select rtrim(wh_code) as WC,wh_name as WN from warehouse_master order by wh_code"

Set RsW = New Recordset
RsW.Open SqlW, Db, adOpenKeyset, adLockOptimistic
cboWhCode.clear
cboWhCode.columnCount = 2
cboWhCode.TextColumn = 1
ir = 0
While Not RsW.EOF
    cboWhCode.AddItem ""
    cboWhCode.List(ir, 0) = RsW!wC
    cboWhCode.List(ir, 1) = Trim$(RsW!wn)
    ir = ir + 1
    RsW.MoveNext
Wend
cboWhCode.ColumnWidths = "60 pt; 180 pt"
cboWhCode.ListWidth = 240
cboWhCode.ListRows = 15

'## receipt Combo
CboRecCls.clear
CboRecCls.columnCount = 2
CboRecCls.TextColumn = 1
CboRecCls.AddItem ""
CboRecCls.List(0, 0) = "R"
CboRecCls.List(0, 1) = "Receipt"
CboRecCls.AddItem ""
CboRecCls.List(1, 0) = "R1"
CboRecCls.List(1, 1) = "Return"
CboRecCls.ColumnWidths = "30 pt; 60 pt"
CboRecCls.ListWidth = 90
CboRecCls.ListRows = 4

'## Unit Combo
Call up_FillCombo(cbounit, "unit_cls")
cbounit.TextColumn = 2
'## Curr Combo
'isiCboUnitCurr cbocurr, isiCurr, 0, 4
Call up_FillCombo(cbocurr, "curr_cls")
cbocurr.TextColumn = 2

Tgl1 = Now()
Tgl2 = Now()
TglReceipt = Now()
DtBCDate = Now()

cmbbox_warehouse = ""
Call up_FillCombo(cboTransport, "Transport_Cls")
With cboTransport
    .ColumnWidths = "30pt,90pt"
    .ListWidth = 120
End With
Call up_FillCombo(cboPackage, "Package_Cls")
With cboPackage
    .ColumnWidths = "60pt,90pt"
    .ListWidth = 150
End With
End Sub

Sub display(Optional lngRow As Long = 1)
Dim ig As Long, id As Long
Dim RsD As Recordset, SQLd As String
Dim Rsd2 As Recordset

Dim Ord As Double, nrecs As Double, nHl As Double

'SQLd = " select  purchaseorder_detail.complete_cls as cek, purchaseorder_detail.item_code,Item_Master.item_name,makeritem_code,  " & vbCrLf & _
'              "         purchaseorder_master.delivery_date,purchaseorder_detail.qty, " & vbCrLf & _
'              "         isnull(receiptqty,0)-isnull(returnqty,0) RecQty, coalesce(rec.Price,purchaseorder_detail.Price, 0) Price, " & vbCrLf & _
'              "         isnull(purchaseorder_detail.Amount, 0) Amount, isnull(purchaseorder_detail.Price_Service, 0) Price_Service,  " & vbCrLf & _
'              "         isnull(purchaseorder_detail.Amount_Service, 0) Amount_Service, WH_code,purchaseorder_master.WHTo,purchaseorder_detail.Unit_cls, " & vbCrLf & _
'              "         purchaseorder_detail.Currency_code   " & vbCrLf & _
'              "         from purchaseorder_master  " & vbCrLf & _
'              "         inner join purchaseorder_detail on purchaseorder_master.PO_no=purchaseorder_detail.PO_no     " & vbCrLf & _
'              "         inner join item_master on purchaseorder_detail.item_code = item_master.item_code     " & vbCrLf & _
'              "         left join (Select item_code,Sum(Qty) as ReceiptQty ,Price=max(Price) " & vbCrLf & _
'              "         from part_receipt   "
'
'SQLd = SQLd + "         where Supplier_code='" & CboPart(0) & "'   and PO_no='" & CboPart(1) & "' and receipt_cls ='R'     " & vbCrLf & _
'              "         group by item_code)rec on purchaseorder_detail.item_code=rec.item_code    left join (Select item_code,Sum(Qty) as ReturnQty  " & vbCrLf & _
'              " from part_receipt where Supplier_code='" & CboPart(0) & "'  and PO_no='" & CboPart(1) & "' and receipt_cls ='R1'  " & vbCrLf & _
'              " group by item_code)ret on purchaseorder_detail.item_code=ret.item_code  where  purchaseorder_master.Supplier_code='" & CboPart(0) & "' and purchaseorder_detail.PO_no='" & CboPart(1) & "' " & vbCrLf

        SQLd = " SELECT PurchaseOrder_Detail.Complete_Cls AS cek , " & vbCrLf & _
                      "         PurchaseOrder_Detail.Item_Code , " & vbCrLf & _
                      "         Item_Master.Item_Name , " & vbCrLf & _
                      "         MakerItem_Code , " & vbCrLf & _
                      "         PurchaseOrder_Master.Delivery_Date , " & vbCrLf & _
                      "         PurchaseOrder_Detail.Qty , " & vbCrLf & _
                      "         ISNULL(ReceiptQty, 0) - ISNULL(ReturnQty, 0) RecQty , " & vbCrLf & _
                      "         COALESCE(rec.Price, PurchaseOrder_Detail.Price, 0) Price , " & vbCrLf & _
                      "         ISNULL(PurchaseOrder_Detail.Amount, 0) Amount , " & vbCrLf & _
                      "         ISNULL(PurchaseOrder_Detail.Price_Service, 0) Price_Service , " & vbCrLf & _
                      "         ISNULL(PurchaseOrder_Detail.Amount_Service, 0) Amount_Service , "
        
        SQLd = SQLd + "         WH_Code , " & vbCrLf & _
                      "         PurchaseOrder_Master.WHTo , " & vbCrLf & _
                      "         PurchaseOrder_Detail.Unit_Cls , " & vbCrLf & _
                      "         PurchaseOrder_Detail.Currency_Code, '02' Receipt_Status " & vbCrLf & _
                      "  FROM   PurchaseOrder_Master " & vbCrLf & _
                      "         INNER JOIN PurchaseOrder_Detail ON PurchaseOrder_Master.PO_No = PurchaseOrder_Detail.PO_No " & vbCrLf & _
                      "         INNER JOIN Item_Master ON PurchaseOrder_Detail.Item_Code = Item_Master.Item_Code " & vbCrLf & _
                      "         LEFT JOIN ( SELECT  Item_Code , " & vbCrLf & _
                      "                             SUM(Qty) AS ReceiptQty , " & vbCrLf & _
                      "                             Price = MAX(Price) " & vbCrLf & _
                      "                     FROM    Part_Receipt "
        
        SQLd = SQLd + "                     WHERE   Supplier_Code = '" & CboPart(0) & "' " & vbCrLf & _
                      "                             AND PO_No = '" & CboPart(1) & "' " & vbCrLf & _
                      "                             AND Receipt_Cls = 'R' " & vbCrLf & _
                      "                     GROUP BY Item_Code " & vbCrLf & _
                      "                   ) rec ON PurchaseOrder_Detail.Item_Code = rec.Item_Code " & vbCrLf & _
                      "         LEFT JOIN ( SELECT  Item_Code , " & vbCrLf & _
                      "                             SUM(Qty) AS ReturnQty " & vbCrLf & _
                      "                     FROM    Part_Receipt " & vbCrLf & _
                      "                     WHERE   Supplier_Code ='" & CboPart(0) & "' " & vbCrLf & _
                      "                             AND PO_No = '" & CboPart(1) & "' " & vbCrLf & _
                      "                             AND Receipt_Cls = 'R1' "
        
        SQLd = SQLd + "                     GROUP BY Item_Code " & vbCrLf & _
                      "                   ) ret ON PurchaseOrder_Detail.Item_Code = ret.Item_Code " & vbCrLf & _
                      "  WHERE  PurchaseOrder_Master.Supplier_Code = '" & CboPart(0) & "' " & vbCrLf & _
                      "         AND PurchaseOrder_Detail.PO_No = '" & CboPart(1) & "' "


Set RsD = Db.Execute(SQLd)
Set Rsd2 = New ADODB.Recordset

With grid
ig = 0
Header
While Not RsD.EOF
        ig = ig + 1
        .GridLines = flexGridRaised
        .AddItem ""
        .TextMatrix(ig, bteColProdCod) = RsD!Item_Code
            .Cell(flexcpAlignment, ig, bteColProdCod) = flexAlignLeftCenter
        .TextMatrix(ig, bteColPartNo) = Trim$(RsD!MakerItem_Code)
            .Cell(flexcpAlignment, ig, bteColPartNo) = flexAlignLeftCenter
        'If (RsD!item_name) <> Null Then
        .TextMatrix(ig, bteColDesc) = IIf(IsNull(Trim$(RsD!item_name)), "", Trim$(RsD!item_name))
        'End If
            .Cell(flexcpAlignment, ig, bteColDesc) = flexAlignLeftCenter
        .TextMatrix(ig, bteColDate) = Format(RsD!delivery_Date, "dd mmm yyyy")
            .Cell(flexcpAlignment, ig, bteColDate) = flexAlignLeftCenter
        .TextMatrix(ig, bteColOrder) = Format(RsD!Qty, gs_formatQty)
            .Cell(flexcpAlignment, ig, bteColOrder) = flexAlignRightCenter
        If IsNull(RsD!recqty) Then
            .TextMatrix(ig, bteColReceipt) = Format(0, gs_formatAmount)
            nrecs = 0
        Else
            If InStr(1, CDbl(RsD!recqty), ".") = 0 Then
                    .TextMatrix(ig, bteColReceipt) = Format(RsD!recqty, gs_formatQty)
                ElseIf InStr(1, CDbl(RsD!recqty), ".") > 0 Then
                    If Split(CDbl(RsD!recqty), ".")(1) <> "" Then .TextMatrix(ig, bteColReceipt) = Format(RsD!recqty, gs_formatQty)
            End If
            nrecs = RsD!recqty
        End If
        .Cell(flexcpAlignment, ig, bteColReceipt) = flexAlignRightCenter
        
        Ord = RsD!Qty
        
        If RsD!cek <> "1" Or IsNull(RsD!cek) Then
            nHl = Ord - nrecs
        Else
            nHl = 0
        End If
        
        If nHl <= 0 Then
            .TextMatrix(ig, bteColRemain) = Format(0, gs_formatQty)
        Else
            If InStr(1, CDbl(nHl), ".") = 0 Then
                    .TextMatrix(ig, bteColRemain) = Format(nHl, gs_formatQty)
            ElseIf InStr(1, CDbl(nHl), ".") > 0 Then
                    If Split(CDbl(nHl), ".")(1) <> "" Then .TextMatrix(ig, bteColRemain) = Format(nHl, gs_formatQty)
            End If
            
        End If
        .Cell(flexcpAlignment, ig, bteColRemain) = flexAlignRightCenter
        
        .TextMatrix(ig, bteColUnit) = uf_GetUnitDescription(Trim$(RsD!Unit_cls))
        .Cell(flexcpAlignment, ig, bteColUnit) = flexAlignCenterCenter
        
        .TextMatrix(ig, bteColCurr) = uf_GetCurrencyDescription(Trim(Trim$(RsD!currency_code)))
        .Cell(flexcpAlignment, ig, bteColCurr) = flexAlignCenterCenter
        
        .TextMatrix(ig, bteColPrice) = Format(RsD!Price, gs_formatPrice)
        .Cell(flexcpAlignment, ig, bteColPrice) = flexAlignRightCenter
                
        .TextMatrix(ig, bteColAmount) = Format(RsD!Amount, gs_formatAmount)
        .Cell(flexcpAlignment, ig, bteColAmount) = flexAlignRightCenter

        .TextMatrix(ig, bteColPriceService) = Format(RsD!Price_Service, gs_formatPrice)
        .Cell(flexcpAlignment, ig, bteColPriceService) = flexAlignRightCenter

        .TextMatrix(ig, bteColAmountService) = Format(RsD!amount_service, gs_formatAmount)
        .Cell(flexcpAlignment, ig, bteColAmountService) = flexAlignRightCenter
        
        'hide
        .TextMatrix(ig, bteColRecStatus) = RsD!Receipt_Status
        .Cell(flexcpAlignment, ig, bteColRecStatus) = flexAlignLeftCenter
        
        If Trim(RsD!WHTo & "") = "" Then
            .TextMatrix(ig, bteColRecWHCode) = RsD!wh_code
        Else
            .TextMatrix(ig, bteColRecWHCode) = RsD!WHTo
        End If
        .TextMatrix(ig, bteColUnitCls) = RsD!Unit_cls
        .TextMatrix(ig, bteColCurrCode) = RsD!currency_code
        .TextMatrix(ig, bteColComplete) = ""
        If IsNull(RsD!cek) Or RsD!cek = 0 Or RsD!cek = "" Then
            .Cell(flexcpChecked, ig, bteColComplete) = flexUnchecked
            .TextMatrix(ig, bteColComplteCls) = 2
        Else
            .Cell(flexcpChecked, ig, bteColComplete) = flexChecked
            .TextMatrix(ig, bteColComplteCls) = 1
        End If
        
        .Cell(flexcpBackColor, ig, bteColProdCod, ig, .ColS - 1) = &HE0E0E0
                
'        SQLd = "Select pr.Seq_No as SeqNo, pr.Supplier_code, pr.PO_no, pr.warehouse_code, pr.Address, pr.receipt_cls, " & _
'            "pr.receipt_date, pr.item_code, pr.qty, pr.unit_cls, pr.currency_code, " & _
'            "coalesce(pr.price,pd.price, 0) price, isnull(pd.amount, 0) amount, isnull(pd.price_service, 0) price_service, isnull(pd.amount_service, 0) amount_service, " & _
'            "pr.Suratjalan_no , pr.Remarks, po.delivery_date, im.Provision_Cls, pr.bc40_no, pr.BC40_Date,pr.BC_Type,pr.Transport_Cls, pr.Package_Qty, pr.Package_Cls " & _
'            "from part_receipt pr " & _
'            "inner join purchaseorder_master po on pr.PO_No = po.po_no " & _
'            "inner join purchaseorder_detail pd on pr.PO_No = pd.po_no and pr.item_code = pd.item_code " & _
'            "inner join Item_master im on pr.Item_Code = im.Item_Code " & _
'            "where pr.Supplier_code ='" & CboPart(0) & "' " & _
'            "and pr.PO_no = '" & CboPart(1) & "' " & _
'            "and pr.item_code = '" & RsD!Item_Code & "' " & _
'            "and (ProductionResult_cls = '0' or ProductionResult_cls is null)"

        SQLd = " SELECT  pr.Seq_No AS SeqNo , " & vbCrLf & _
                      "         pr.Supplier_Code , " & vbCrLf & _
                      "         pr.PO_No , " & vbCrLf & _
                      "         pr.Warehouse_Code , " & vbCrLf & _
                      "         pr.Address , " & vbCrLf & _
                      "         pr.Receipt_Cls , " & vbCrLf & _
                      "         pr.Receipt_Date , " & vbCrLf & _
                      "         pr.Item_Code , " & vbCrLf & _
                      "         pr.Qty , " & vbCrLf & _
                      "         pr.Unit_Cls , " & vbCrLf & _
                      "         pr.Currency_Code , "
        
        SQLd = SQLd + "         COALESCE(pr.Price, pd.Price, 0) price , " & vbCrLf & _
                      "         ISNULL(pd.Amount, 0) amount , " & vbCrLf & _
                      "         ISNULL(pd.Price_Service, 0) price_service , " & vbCrLf & _
                      "         ISNULL(pd.Amount_Service, 0) amount_service , " & vbCrLf & _
                      "         pr.SuratJalan_No , " & vbCrLf & _
                      "         pr.Remarks , " & vbCrLf & _
                      "         po.Delivery_Date , " & vbCrLf & _
                      "         im.Provision_Cls , " & vbCrLf & _
                      "         pr.BC40_No , " & vbCrLf & _
                      "         pr.BC40_Date , " & vbCrLf & _
                      "         pr.BC_Type , "
        
        SQLd = SQLd + "         pr.Transport_Cls , " & vbCrLf & _
                      "         pr.Package_Qty , " & vbCrLf & _
                      "         pr.Package_Cls, " & vbCrLf & _
                      "         ISNULL(pr.Receipt_Status, '02')Receipt_Status, " & vbCrLf & _
                      "         ISNULL(pr.No_Register, '')No_Register, " & vbCrLf & _
                      "         ISNULL(pr.No_Seri, '')No_Seri " & vbCrLf & _
                      " FROM    Part_Receipt pr " & vbCrLf & _
                      "         INNER JOIN PurchaseOrder_Master po ON pr.PO_No = po.PO_No " & vbCrLf & _
                      "         INNER JOIN PurchaseOrder_Detail pd ON pr.PO_No = pd.PO_No " & vbCrLf & _
                      "                                               AND pr.Item_Code = pd.Item_Code " & vbCrLf & _
                      "         INNER JOIN Item_Master im ON pr.Item_Code = im.Item_Code " & vbCrLf & _
                      " WHERE   pr.Supplier_Code = '" & CboPart(0) & "'  " & vbCrLf & _
                      "         AND pr.PO_No = '" & CboPart(1) & "' "
        
        SQLd = SQLd + "         AND pr.Item_Code = '" & RsD!Item_Code & "' " & vbCrLf & _
                      "         AND ( ProductionResult_Cls = '0' " & vbCrLf & _
                      "               OR ProductionResult_Cls IS NULL " & vbCrLf & _
                      "             ) "

        If Rsd2.State = adStateOpen Then Rsd2.Close
        Rsd2.Open SQLd, Db, adOpenKeyset
        id = 0
        
        While Not Rsd2.EOF
            id = id + 1
            .AddItem ""
            .TextMatrix(ig + id, bteColProdCod) = " "
            .TextMatrix(ig + id, bteColPartNo) = " "
            .TextMatrix(ig + id, bteColDesc) = Trim$(Rsd2!SuratJalan_No)
            .Cell(flexcpAlignment, ig + id, bteColDesc) = flexAlignLeftCenter
            .TextMatrix(ig + id, bteColDate) = Format(Rsd2!Receipt_Date, "dd mmm yyyy")
            .Cell(flexcpAlignment, ig + id, bteColDate) = flexAlignLeftCenter
            .TextMatrix(ig + id, bteColCls) = Trim(Rsd2!receipt_cls)
            .TextMatrix(ig + id, bteColBC40) = Trim$(Rsd2!BC40_No & "")
            .TextMatrix(ig + id, bteColBctype) = Trim(Rsd2!BC_Type & "")
            .TextMatrix(ig + id, bteColBCDate) = Format(Rsd2!BC40_Date, "dd mmm yyyy")
            .TextMatrix(ig + id, bteColTransport) = Trim$(Rsd2!Transport_Cls & "")
            .TextMatrix(ig + id, bteColPackageQty) = Val(Rsd2!Package_Qty & "")
            .TextMatrix(ig + id, bteColPackage) = Trim$(Rsd2!Package_Cls & "")
            If IsNull(Rsd2!Qty) Then
                .TextMatrix(ig + id, bteColReceipt) = Format(0, gs_formatQty)
            Else
                If InStr(1, CDbl(Rsd2!Qty), ".") = 0 Then
                    .TextMatrix(ig + id, bteColReceipt) = Format(Rsd2!Qty, gs_formatQty)
                ElseIf InStr(1, CDbl(Rsd2!Qty), ".") > 0 Then
                    If Split(CDbl(Rsd2!Qty), ".")(1) <> "" Then .TextMatrix(ig + id, bteColReceipt) = Format(Rsd2!Qty, gs_formatQty)
                End If
                
            End If
            .Cell(flexcpAlignment, ig + id, bteColReceipt) = flexAlignRightCenter
            
            .TextMatrix(ig + id, bteColUnit) = uf_GetUnitDescription(Trim$(RsD!Unit_cls))
            .Cell(flexcpAlignment, ig + id, bteColUnit) = flexAlignCenterCenter
            
            .TextMatrix(ig + id, bteColCurr) = uf_GetCurrencyDescription(Trim$(RsD!currency_code))
            .Cell(flexcpAlignment, ig, bteColCurr) = flexAlignCenterCenter
        
            .TextMatrix(ig + id, bteColPrice) = Format(Rsd2!Price, gs_formatPrice)
            .Cell(flexcpAlignment, ig + id, bteColPrice) = flexAlignRightCenter
            
            .TextMatrix(ig + id, bteColAmount) = Format(Rsd2!Amount, gs_formatAmount)
            .Cell(flexcpAlignment, ig + id, bteColAmount) = flexAlignRightCenter
            
            .TextMatrix(ig + id, bteColPriceService) = Format(Rsd2!Price_Service, gs_formatPrice)
            .Cell(flexcpAlignment, ig + id, bteColPriceService) = flexAlignRightCenter
            
            .TextMatrix(ig + id, bteColAmountService) = Format(Rsd2!amount_service, gs_formatAmount)
            .Cell(flexcpAlignment, ig + id, bteColAmountService) = flexAlignRightCenter
            
            '#----
            .TextMatrix(ig + id, bteColItemCode) = Trim$(RsD!Item_Code) 'Hide
            .TextMatrix(ig + id, bteColDelDate) = Trim$(RsD!delivery_Date) 'Hide
            '#----
            .TextMatrix(ig + id, bteColWHCode) = Trim$(Rsd2!Warehouse_Code) 'Hide
            .TextMatrix(ig + id, bteColAddress) = Trim$(Rsd2!Address) 'Hide
            .TextMatrix(ig + id, bteColRecCls) = Trim$(Rsd2!receipt_cls) 'Hide
            .TextMatrix(ig + id, bteColRecDate) = Trim$(Rsd2!Receipt_Date) 'Hide
            .TextMatrix(ig + id, bteColRecUnit) = Trim$(Rsd2!Unit_cls) 'Hide
            .TextMatrix(ig + id, bteColRecCurr) = Trim$(Rsd2!currency_code) 'Hide
            .TextMatrix(ig + id, bteColRem) = Trim$(Rsd2!Remarks)
            '#----
            .TextMatrix(ig + id, bteColQtyOrder) = Trim$(RsD!Qty) 'Order QTY Hide
            .TextMatrix(ig + id, bteColQtyRec) = Trim$(RsD!recqty) 'Order QTY Hide
            .TextMatrix(ig + id, bteColRecPrice) = Trim$(RsD!Price) 'Order QTY Hide
            .TextMatrix(ig + id, bteColSeqNo) = Rsd2!seqNo 'Order QTY Hide
            .TextMatrix(ig + id, bteColProvision) = Rsd2!Provision_Cls 'Provision Cls
            .TextMatrix(ig + id, bteColRecStatus) = Rsd2!Receipt_Status 'Rec Status
            .TextMatrix(ig + id, bteColNoRegister) = Rsd2!No_Register 'Rec Status
            .TextMatrix(ig + id, bteColNoSeri) = Rsd2!No_Seri 'Rec Status
                        
            .Cell(flexcpBackColor, ig + id, bteColProdCod, ig + id, bteColRecDate) = &H80000018
            
'            .MergeRow(ig + id) = True
'            .MergeCells = flexMergeRestrictRows
            
            Rsd2.MoveNext
        Wend
        ig = ig + id
        .Cell(flexcpBackColor, bteColProdCod, bteColSelect, ig, bteColSelect) = vbWhite
    RsD.MoveNext
Wend
End With

If lngRow < grid.Rows Then grid.Row = lngRow: grid.TopRow = lngRow
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If newDb.State <> adStateClosed Then newDb.Close
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim StrGrid As String
Dim AdaS As Boolean, brs As Integer, id As Long
Dim d As Integer

StrGrid = grid.Text
AdaS = False
'Pakai False

If StrGrid = "S" Then
    For id = 1 To grid.Rows - 1
        If id <> Row Then grid.TextMatrix(id, bteColSelect) = ""
    Next id
    If Trim(grid.TextMatrix(grid.Row, bteColProdCod)) = "" Then
        brs = 0
        OrderQty = 0
        baru = False
        cboitem = grid.TextMatrix(grid.Row, bteColItemCode)
        TxtSj = grid.TextMatrix(grid.Row, bteColDesc)
        txtBCNo = grid.TextMatrix(grid.Row, bteColBC40)
        DtBCDate = IIf(grid.TextMatrix(grid.Row, bteColBCDate) = "", grid.TextMatrix(grid.Row, bteColRecDate), grid.TextMatrix(grid.Row, bteColBCDate))
        CbotypeBC = grid.TextMatrix(grid.Row, bteColBctype)
        CboRecCls = grid.TextMatrix(grid.Row, bteColRecCls)
        TglReceipt = grid.TextMatrix(grid.Row, bteColRecDate)
        cboWhCode = grid.TextMatrix(grid.Row, bteColWHCode)
        
        If IsNull(grid.TextMatrix(grid.Row, bteColReceipt)) Then
            txtQty = Format(0, gs_formatQty)
        Else
            If InStr(1, CDbl(grid.TextMatrix(grid.Row, bteColReceipt)), ".") = 0 Then
                txtQty = Format(grid.TextMatrix(grid.Row, bteColReceipt), gs_formatQty)
            ElseIf InStr(1, CDbl(grid.TextMatrix(grid.Row, bteColReceipt)), ".") > 0 Then
                If Split(CDbl(grid.TextMatrix(grid.Row, bteColReceipt)), ".")(1) <> "" Then txtQty = Format(grid.TextMatrix(grid.Row, bteColReceipt), gs_formatQty)
            End If
        End If
        tampungQty = CDbl(txtQty)
        
        cbounit = uf_GetUnitDescription(Trim$(grid.TextMatrix(grid.Row, bteColRecUnit)))
        cbocurr = uf_GetCurrencyDescription(Trim$(grid.TextMatrix(grid.Row, bteColRecCurr)))
        cboprice.Text = Format(CDbl(grid.TextMatrix(grid.Row, bteColPrice)) + CDbl(grid.TextMatrix(grid.Row, bteColPriceService)), gs_formatPrice)
        txtamount = Format(CDbl(txtQty) * CDbl(cboprice), gs_formatAmount)
        txtremarks = Trim$(grid.TextMatrix(grid.Row, bteColRem))
        txtBCNo = Trim$(grid.TextMatrix(grid.Row, bteColBC40))
        CbotypeBC = Trim(grid.TextMatrix(grid.Row, bteColBctype))
        cboTransport = Trim$(grid.TextMatrix(grid.Row, bteColTransport))
        txtPackage = Val(grid.TextMatrix(grid.Row, bteColPackageQty))
        cboPackage = Trim$(grid.TextMatrix(grid.Row, bteColPackage))
        
        'Label u/ update
        LblRecCls = grid.TextMatrix(grid.Row, bteColRecCls)
        LblRecDate = grid.TextMatrix(grid.Row, bteColRecDate)
        lblWHCode = grid.TextMatrix(grid.Row, bteColWHCode)
        lblItemCode = grid.TextMatrix(grid.Row, bteColItemCode)
        lblQty = CDbl(grid.TextMatrix(grid.Row, bteColReceipt))
        OrderQty = grid.TextMatrix(grid.Row, bteColQtyOrder)
        SeqID = grid.TextMatrix(grid.Row, bteColSeqNo)
        provisionCls = grid.TextMatrix(grid.Row, bteColProvision)
        QTYRec = grid.TextMatrix(grid.Row, bteColQtyRec)
        If CbotypeBC.Text = "4.0" Then
            txtRec_Status.Text = grid.TextMatrix(grid.Row, bteColRecStatus)
        ElseIf CbotypeBC.Text = "2.6.2" Then
            txtRec_Status.Text = grid.TextMatrix(grid.Row, bteColRecStatus)
        ElseIf CbotypeBC.Text = "2.3" Then
            txtRec_Status.Text = grid.TextMatrix(grid.Row, bteColRecStatus)
        Else
            txtRec_Status.Text = "01"
        End If
        
        txtRegisterNo.Text = grid.TextMatrix(grid.Row, bteColNoRegister)
        txtNoSeri.Text = grid.TextMatrix(grid.Row, bteColNoSeri)
        
        LblErr = ""
        nPrice = grid.TextMatrix(grid.Row, bteColRecPrice)
  Else
        brs = 0
        OrderQty = 0
        baru = True
        cboitem = grid.TextMatrix(grid.Row, bteColProdCod)
        OrderQty = grid.TextMatrix(grid.Row, bteColOrder)
        QTYRec = grid.TextMatrix(grid.Row, bteColReceipt)
        If CbotypeBC.Text = "4.0" Then
            txtRec_Status.Text = grid.TextMatrix(grid.Row, bteColRecStatus)
        ElseIf CbotypeBC.Text = "2.6.2" Then
            txtRec_Status.Text = grid.TextMatrix(grid.Row, bteColRecStatus)
        ElseIf CbotypeBC.Text = "2.3" Then
            txtRec_Status.Text = grid.TextMatrix(grid.Row, bteColRecStatus)
        Else
            txtRec_Status.Text = "01"
        End If
        'txtRegisterNo.Text = Grid.TextMatrix(Grid.Row, bteColNoRegister)
        
        CboRecCls.ListIndex = 0
        cboprice.Text = Format(CDbl(grid.TextMatrix(grid.Row, bteColPrice)) + CDbl(grid.TextMatrix(grid.Row, bteColPriceService)), gs_formatPrice)
        
        up_GetNoSeri (cboitem.Text)
        
        GoTo x2
x2:
        nPrice = CDbl(grid.TextMatrix(grid.Row, bteColPrice)) + CDbl(grid.TextMatrix(grid.Row, bteColPriceService))
        cboWhCode = Trim$(grid.TextMatrix(grid.Row, bteColRecWHCode))
        cbounit = uf_GetUnitDescription(Trim$(grid.TextMatrix(grid.Row, bteColUnitCls)))
        cbocurr = uf_GetCurrencyDescription(Trim$(grid.TextMatrix(grid.Row, bteColCurrCode)))
        If CDbl(QTYRec) <> 0 And CDbl(OrderQty) <> 0 Then
            If CDbl(OrderQty) < CDbl(QTYRec) Then
                txtQty = Format(0, gs_formatQty)
            Else
                txtQty = CDbl(OrderQty) - CDbl(QTYRec)
            End If
        ElseIf CDbl(OrderQty) <> 0 And CDbl(QTYRec) = 0 Then
            txtQty = CDbl(OrderQty)
        End If
        txtQty = Format(txtQty, gs_formatQty)
        txtPackage.Text = Format(0, gs_formatBox)
        LblErr = ""
  End If
    
ElseIf StrGrid = "D" Then
    For id = 1 To grid.Rows - 1
        'Jika ada S maka , hapus yg S
        If grid.TextMatrix(id, bteColSelect) = "S" Then grid.TextMatrix(id, bteColSelect) = "": Exit For
    Next id
    LblErr = ""
    Kosong 1
    baru = False
Else
    If Col = bteColComplete Then
        If grid.Cell(flexcpChecked, Row, bteColComplete) = 1 Then
            grid.TextMatrix(Row, bteColRemain) = Format(0, gs_formatQty)
        Else
            grid.TextMatrix(Row, bteColRemain) = CDbl(grid.TextMatrix(Row, bteColOrder)) - CDbl(grid.TextMatrix(Row, bteColReceipt))
        End If
    End If
End If

If Not IsNumeric(txtQty.Text) Then txtQty.Text = 0
dblQtyLama = CDbl(txtQty)
dblQty = 0
d = grid.Row
dblQty = CDbl(IIf(grid.TextMatrix(d, bteColRemain) = "", 0, grid.TextMatrix(d, bteColRemain)))
Do While grid.TextMatrix(d, bteColRemain) = "" And d > 0
    d = d - 1
Loop 'Until grid.TextMatrix(d, bteColRemain) <> ""
If d > 0 And dblQty = 0 Then
    dblQty = CDbl(IIf(grid.TextMatrix(d, bteColRemain) = "", 0, grid.TextMatrix(d, bteColRemain)))
End If
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim pesanTgl As String
If Col <> bteColSelect And Col <> bteColComplete Then Cancel = 1
If grid.Cell(flexcpBackColor, Row, 1) = &H80000018 Or Col = bteColComplete Then
    pesanTgl = up_ValidateDateRange(Format(grid.TextMatrix(Row, bteColDate), "yyyy-MM-dd"), True)
    If pesanTgl <> "" Then LblErr = pesanTgl: Cancel = 1 Else LblErr = ""
End If
End Sub

Private Sub Grid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
If grid.Col = bteColSelect Then
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii <> Asc("D") And KeyAscii <> Asc("S") And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyReturn Then
        KeyAscii = 0
    End If
    If KeyAscii = vbKeyEscape Then KeyAscii = 0
    If KeyAscii = Asc(".") Then KeyAscii = 0
End If
End Sub
Function NamaItem(StrItem As String) As String
Dim rsin As Recordset
Set rsin = Db.Execute("Select item_name from item_master where item_code='" & StrItem & "'")
If Not rsin.EOF Then
    NamaItem = rsin!item_name
Else
    NamaItem = ""
End If
rsin.Close
Set rsin = Nothing

End Function

Function MakerItem(StrItem As String) As String
Dim RsMI As Recordset
Set RsMI = Db.Execute("Select makeritem_code from item_master where item_code='" & StrItem & "'")
If Not RsMI.EOF Then
    MakerItem = RsMI!MakerItem_Code
Else
    MakerItem = ""
End If
RsMI.Close
Set RsMI = Nothing
End Function


Private Sub TglReceipt_Change()
'If cboCurr.ListCount > 0 Then
'    cboCurr = Trim(cboCurr)
'    If cboCurr.MatchFound Then
'        xcurr = cboCurr.List(, 0)
'        blnshow = False
'        Call BrowsePrice(cboitem, TglReceipt, CboPart(0), xcurr)
'    Else
'        xcurr = ""
'        LblErr = DisplayMsg(4005): cboCurr.SetFocus
'    End If
'End If
End Sub

Private Sub TxtBCNo_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtNoSeri_KeyPress(KeyAscii As Integer)
    ' Hanya izinkan angka (48-57) dan tombol Backspace (8)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> 8 Then
        KeyAscii = 0 ' Batalkan input
    End If
End Sub

Private Sub txtPackage_KeyPress(KeyAscii As Integer)
    If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub txtPackage_LostFocus()
If IsNumeric(txtPackage) = True Then
    txtPackage.Text = Format(txtPackage.Text, gs_formatBox)
Else
    txtPackage.Text = Format(0, gs_formatBox)
End If
End Sub

Private Sub txtqty_Change()

If InStr(1, txtQty.Text, ",") = 1 Then txtQty.Text = Mid(txtQty.Text, 2, Len(txtQty.Text))
If InStr(1, txtQty.Text, ".") = 1 Then txtQty = Format(0, gs_formatQty)

If Trim(txtQty) = "" Then txtamount = Format(0, gs_formatAmount): Exit Sub
If Trim(cboprice) = "" Then txtamount = Format(0, gs_formatAmount): Exit Sub

If IsNumeric(txtQty) = True And IsNumeric(cboprice) = True Then
    If CDbl(txtQty) > 0 And CDbl(cboprice) > 0 Then
        txtamount = CDbl(txtQty) * CDbl(cboprice)
        If Round(CDbl(txtamount)) / CDbl(txtamount) = 1 Then
            txtamount = Format(CDbl(txtamount), gs_formatAmount)
        Else
            txtamount = Format(CDbl(txtamount), gs_formatAmount)
        End If
    Else
        txtamount = Format(0, gs_formatAmount)
    End If
Else
    txtamount = Format(0, gs_formatAmount)
End If

txtamount = Format(CDbl(txtamount), gs_formatAmount)
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub up_SubConDeleteConsumption(nmDb As Connection, nmRS As ADODB.Recordset, SubConSeqNo As String, blnFix As Integer, thnFix As Integer)
    Dim sql As String
    
    With nmRS
        Do While Not .EOF
            If !stockItem = "01" And !stockWH = "01" Then
                Call seqNo.updateStock(!FromWarehouse_Code, _
                !childitem_code, !ChildRequirement_qty, "", Format(!childsupply_date, "yyyy-MM-dd"), _
                (blnFix), (thnFix), nmDb, "Supply", 0, 0)
            End If
            .MoveNext
        Loop
    End With
    
    sql = "delete  Part_Supply where SubConPartReceipt_SeqNo in (" & Trim(SubConSeqNo) & ")"
    nmDb.Execute sql
End Sub

Sub DataGrid()
Dim kode As String, Sta As String
Dim strSQL As String
On Error Resume Next
Dim PosS As Integer
Dim rsProv As New ADODB.Recordset

Dim rsDelSubCon As New ADODB.Recordset

PosS = grid.FindRow("S", , bteColSelect, False)

If PosS > 0 Then
    Db.BeginTrans
    
    kode = Trim$(grid.TextMatrix(PosS, 1))
    
    strSQL = "update part_receipt " & _
        "set Supplier_code='" & Trim$(CboPart(0).Text) & "', " & _
        "PO_no='" & Trim$(CboPart(1).Text) & "', " & _
        "warehouse_code='" & Trim$(cboWhCode) & "' ," & _
        "receipt_cls='" & Trim$(CboRecCls) & "' ," & _
        "receipt_date='" & Format(TglReceipt, "yyyy-mm-dd") & "', " & _
        "item_code='" & Trim$(cboitem) & "'," & _
        "qty=" & CDbl(txtQty) & "," & _
        "BC40_No='" & Trim(txtBCNo) & "'," & _
        "BC40_Date='" & Format(DtBCDate, "yyyy-mm-dd") & "', " & _
        "BC_type ='" & Trim(CbotypeBC.Text) & "'," & _
        "Transport_Cls='" & Trim(cboTransport) & "'," & _
        "Package_Qty=" & CDbl(txtPackage) & "," & _
        "Package_Cls='" & Trim(cboPackage) & "'," & _
        "Unit_cls='" & Trim$(cbounit.List(cbounit.ListIndex, 0)) & "', " & _
        "currency_code='" & Trim$(cbocurr.List(cbocurr.ListIndex, 0)) & "'," & _
        "price=" & CDbl(cboprice) & ", " & _
        "amount=" & Trim$(CDbl(txtamount)) & "," & _
        "suratjalan_no='" & Trim$(TxtSj) & "', " & _
        "No_Seri='" & Trim$(txtNoSeri.Text) & "', " & _
        "Remarks='" & Trim$(txtremarks) & "'," & _
        "Last_Update = getdate(), " & _
        "Last_User = '" & userLogin & "' " & _
        "Where Seq_No='" & SeqID & "'"

    Db.Execute (strSQL)

    If err.number <> 0 Then
        StaErr = True
        err.clear
        Db.RollbackTrans
    Else
        StaErr = False
         
        
        If txtRec_Status.Text = "01" Then
            If Format(LblRecDate, "YYYYMM") <> Format(TglReceipt, "YYYYMM") Then
                If Trim$(lblWHCode) <> Trim$(cboWhCode.List(cboWhCode.ListIndex, 0)) Then
                    ProsesStock 2, Trim$(cboitem), Trim$(lblWHCode), Trim$(lblWHCode), Format(LblRecDate, "YYYYMM"), 0, CDbl(lblQty) 'Update yg lama
                    ProsesStock 2, Trim$(cboitem), Trim$(cboWhCode.List(cboWhCode.ListIndex, 0)), Trim$(cboWhCode.List(cboWhCode.ListIndex, 0)), Format(TglReceipt, "YYYYMM"), CDbl(txtQty), 0 'Update yg baru
                Else
                    ProsesStock 2, Trim$(cboitem), Trim$(lblWHCode), Trim$(lblWHCode), Format(LblRecDate, "YYYYMM"), 0, CDbl(lblQty) 'Update yg lama
                    ProsesStock 2, Trim$(cboitem), Trim$(lblWHCode), Trim$(lblWHCode), Format(TglReceipt, "YYYYMM"), CDbl(txtQty), 0 'Update yg baru
                End If
            Else
                If Trim$(lblWHCode) <> Trim$(cboWhCode.List(cboWhCode.ListIndex, 0)) Then
                    ProsesStock 2, Trim$(cboitem), Trim$(lblWHCode), Trim$(lblWHCode), Format(LblRecDate, "YYYYMM"), 0, CDbl(lblQty) 'Update yg lama
                    ProsesStock 2, Trim$(cboitem), Trim$(cboWhCode.List(cboWhCode.ListIndex, 0)), Trim$(cboWhCode.List(cboWhCode.ListIndex, 0)), Format(TglReceipt, "YYYYMM"), CDbl(txtQty), 0 'Update yg baru
                Else
                    ProsesStock 2, Trim$(cboitem), Trim$(lblWHCode), Trim$(cboWhCode), Format(TglReceipt, "YYYYMM"), CDbl(txtQty), CDbl(lblQty)
                End If
            End If
        End If
        
        'Get Consumption Subcon
        'Herfin 20070606
        '===================================================================
        sql = "Select Supp.*,Item.StockControl_Cls stockItem, " & _
            "Isnull(WH.StockControl_Cls,'01') StockWH " & _
            "from Part_Supply Supp Inner Join Item_Master Item on Supp.ChildItem_Code = Item.Item_Code " & _
            "   Left Join Warehouse_Master WH  Supp.FromWarehouse_Code = WH.WH_Code" & _
            "where 'A'='A' "
        
        sql = sql & " and SubConPartReceipt_SeqNo in('R" & Trim(SeqID) & "')"
        
        If rsDelSubCon.State = 1 Then rsDelSubCon.Close
        rsDelSubCon.Open sql, Db, adOpenKeyset, adLockOptimistic
        
        '===================================================================
        
        '*********** Proses Update ke Supply **************
        
        KeyProd = "R" & SeqID
        Call seqNo.HapusDataSupp(Db, "'" & KeyProd & "'", blnFix, thnFix)
        
        sql = "Select Provision_Cls from Item_MaSter where ITem_Code = '" & cboitem & "'"
        Set rsProv = Db.Execute(sql)
        If Not rsProv.EOF Then provisionCls = Trim(rsProv(0)) Else provisionCls = ""
            
        'Cancel Supply - Update for KAWAI
        'If provisionCls = "01" Then Call inputSupply(CboItem, CDbl(txtQty), 0)
        '*************************************************
        
               
        'Delete Consumption Subcon
        'Herfin 20070606
        '===================================================================
        If uf_GetSubConStatus(Trim(CboPart(0))) = "3" Then
            Call up_SubConDeleteConsumption(Db, rsDelSubCon, "'R" & SeqID & "'", (blnFix), (thnFix))
            If rsDelSubCon.State = 1 Then rsDelSubCon.Close
            Call up_SubConInputConsumption(cboitem, CDbl(txtQty), Trim(cmbbox_warehouse))
        End If
        '===================================================================
        
    End If
    
    Db.CommitTrans
    Exit Sub
End If

nErr = 0
nOK = 0
StrWDel = ""

Db.BeginTrans
    
Dim ix As Long
Dim strKey As String
Dim strKey2 As String

strKey = ""
strKey2 = ""
For ix = 1 To grid.Rows - 1
    kode = Trim$(grid.TextMatrix(ix, bteColCurr))
    LblRecCls = grid.TextMatrix(ix, bteColRecCls)
    LblRecDate = grid.TextMatrix(ix, bteColRecDate)
    lblWHCode = grid.TextMatrix(ix, bteColWHCode)
    lblItemCode = grid.TextMatrix(ix, bteColItemCode)
    lblQty = grid.TextMatrix(ix, bteColReceipt)
    provisionCls = grid.TextMatrix(ix, bteColProvision)
    Sta = Trim$(grid.TextMatrix(ix, bteColSelect))
    If Sta = "D" Then
        strSQL = "Delete from part_receipt where Seq_No='" & grid.TextMatrix(ix, bteColSeqNo) & "'"
        If strSQL <> "" Then Db.Execute strSQL
        
'        Db.Execute "Update Purchaseorder_detail with (updlock) set Complete_cls='0', Last_Update = getdate(), Last_User = '" & userLogin & "' " & _
'            "where PO_no='" & Trim$(CboPart(1)) & "' and item_code='" & Trim$(lblItemCode) & "'"
        
        Db.Execute "update purchaseorder_detail with (updlock) set complete_cls = " & _
            "case when " & _
                "(select isnull(sum(qty), 0) from part_receipt where receipt_cls = 'R' and po_no = purchaseorder_detail.po_no and item_code = purchaseorder_detail.item_code) - " & _
                "(select isnull(sum(qty), 0) from part_receipt where receipt_cls = 'R1' and po_no = purchaseorder_detail.po_no and item_code = purchaseorder_detail.item_code) >= qty " & _
            "then 1 else 0 end " & _
            "where po_no='" & Trim$(CboPart(1)) & "' and item_code='" & Trim$(lblItemCode) & "' "
        
        If err.number <> 0 Then
            StrWDel = StrWDel & kode & ","
            nErr = nErr + 1
            err.clear
        Else
            If txtRec_Status.Text = "01" Then
                 ProsesStock 3, Trim$(lblItemCode), Trim$(lblWHCode), Trim$(lblWHCode), Format(LblRecDate, "YYYYMM"), CDbl(lblQty), "0"
                nOK = nOK + 1
            End If
        End If
        
        '*********** Proses hapus ke Supply **************
        strKey = strKey & "'R" & grid.TextMatrix(ix, bteColSeqNo) & "',"
        strKey2 = strKey2 & "'R" & grid.TextMatrix(ix, bteColSeqNo) & "',"
        '*************************************************
    End If
    strSQL = ""
Next ix

If Len(StrWDel) > 1 Then StrWDel = Mid(StrWDel, 1, Len(StrWDel) - 1)

Call seqNo.HapusDataSupp(Db, Left(strKey, Len(strKey) - 1), (blnFix), (thnFix))

'Get Consumption Subcon
'Herfin 20070606
'===================================================================
sql = "Select Supp.*,Item.StockControl_Cls stockItem, " & _
    "Isnull(WH.StockControl_Cls,'01') StockWH " & _
    "from Part_Supply Supp Inner join Item_Master Item on Supp.ChildItem_Code = Item.Item_Code " & _
    "   Left Join Warehouse_Master WH on supp.FromWarehouse_Code = WH.WH_Code " & _
    "where 'A'='A' "

sql = sql + " and SubConPartReceipt_SeqNo in (" & Left(strKey2, Len(strKey2) - 1) & ")"

If rsDelSubCon.State = 1 Then rsDelSubCon.Close
rsDelSubCon.Open sql, Db, adOpenKeyset, adLockOptimistic

'===================================================================

'Delete Consumption Subcon
'Herfin 20070606
'===================================================================

If uf_GetSubConStatus(Trim(CboPart(0))) = "3" Then
    Call up_SubConDeleteConsumption(Db, rsDelSubCon, Left(strKey2, Len(strKey2) - 1), blnFix, thnFix)
End If
If rsDelSubCon.State = 1 Then rsDelSubCon.Close
'===================================================================

If err.number = 0 Then
    Db.CommitTrans
Else
    Db.RollbackTrans
End If

If nErr > 0 Then StaErr = True

kode = ""
Sta = ""
strSQL = ""
display

End Sub

Sub Kosong(nK As Byte)
Select Case nK
Case 0
    CboPart(0) = ""
    CboPart(1) = ""
    Tgl1 = Now()
    Tgl2 = Now()
Case 1

    cboitem = ""
    LblItemName = ""
    cboWhCode = ""
    Lbladdress = ""
    cbounit = ""
    cbocurr = ""
    cboprice.Text = Format(0, gs_formatPrice)
    txtQty = Format(0, gs_formatQty)
    txtamount = Format(0, gs_formatAmount)
    txtremarks = ""
    lblWHCode = ""
    lblItemCode = ""
    LblRecCls = ""
    LblRecDate = ""
    txtNoSeri.Text = ""
'    txtbcno = ""
'    TxtSj = ""
'    CbotypeBC = ""
    
End Select
End Sub

Function cek() As Boolean
cek = True
If Trim$(CboPart(0)) = "" Then LblErr = DisplayMsg(1054):  CboPart(0).SetFocus: cek = False: Exit Function
CboPart(0) = CboPart(0)
If CboPart(0).MatchFound = False Then LblErr = DisplayMsg("0032"): CboPart(0).SetFocus: cek = False: Exit Function

If Trim$(CboPart(1)) = "" Then LblErr = DisplayMsg(1046):  CboPart(1).SetFocus: cek = False: Exit Function
CboPart(1) = CboPart(1)
If CboPart(1).MatchFound = False Then LblErr = DisplayMsg(4015): CboPart(1).SetFocus: cek = False: Exit Function

If Trim(TxtSj) = "" Then LblErr = DisplayMsg(1036): TxtSj.SetFocus:  cek = False: Exit Function
If Trim(CboRecCls) = "" Then LblErr = DisplayMsg(8088): CboRecCls.SetFocus: cek = False: Exit Function
CboRecCls = CboRecCls
If CboRecCls.MatchFound = False Then LblErr = DisplayMsg(8088): CboRecCls.SetFocus: cek = False: Exit Function

If Trim$(cboitem) = "" Then LblErr = DisplayMsg(1009): cboitem.SetFocus: cek = False: Exit Function

If Trim$(cboWhCode) = "" Then LblErr = DisplayMsg(1042): cboWhCode.SetFocus: cek = False: Exit Function
cboWhCode = cboWhCode
If cboWhCode.MatchFound = False Then LblErr = DisplayMsg(4023): cboWhCode.SetFocus:  cek = False: Exit Function
If CDbl(txtQty) = 0 Then LblErr = DisplayMsg(1012): txtQty.SetFocus: cek = False: Exit Function
If Trim$(cbounit) = "" Then LblErr = DisplayMsg(1030): cek = False: Exit Function 'Cbounit.SetFocus
cbounit = cbounit
If cbounit.MatchFound = False Then LblErr = DisplayMsg(1030):  cek = False: Exit Function 'Cbounit.SetFocus

If txtNoSeri.Text = "" Then LblErr = DisplayMsg("0001") & " No Seri ! ":  txtNoSeri.SetFocus: cek = False: Exit Function

If Trim$(cbocurr) = "" Then
    If bteHakPrice = 0 Then
        cbocurr = uf_GetCurrencyDescription(gs_DefaultCurrencyCode)
    Else
        LblErr = DisplayMsg(1011)
        cbocurr.SetFocus
        cek = False
        Exit Function
    End If
End If
cbocurr = cbocurr
If cbocurr.MatchFound = False Then LblErr = DisplayMsg(1028): cbocurr.SetFocus: cek = False: Exit Function

If CboPart(0).Text <> "S0028" Then 'Request Pak toha jika ada khusus untuk supplier s0028 tidak perlu validasi price 0
If CDbl(cboprice) = 0 Then
    If bteHakPrice = 0 Then
        cboprice.Text = Format(0, gs_formatPrice)
    Else
        LblErr = DisplayMsg(1029)
        cboprice.SetFocus
        cek = False
        Exit Function
    End If
End If
End If

LblErr = up_ValidateSuratJalan()
If validate <> True Then
    LblErr = DisplayMsg(9013)
    cek = False
    Exit Function
End If
'Validasi closing
LblErr = up_ValidateDateRange(Format(TglReceipt, "yyyy-MM-dd"), True)
If LblErr <> "" Then
    cek = False
    Exit Function
End If

'LblErr = uf_ValidClosingReceipt(Format(TglReceipt, "yyyy-MM-dd"))
'If LblErr <> "" Then
'    cek = False
'    Exit Function
'End If

End Function

Private Sub tgl1_Change()
Tgl2_Click
End Sub

Private Sub tgl1_Click()
Tgl2_Click
End Sub

Private Sub Tgl2_Change()
Tgl2_Click
End Sub

Private Sub Tgl2_Click()
Dim rsPO As New Recordset
Dim ls_sql As String

CboPart(0) = CboPart(0)
If CboPart(0).MatchFound = True Then
    CboPart(1).clear
    ' PO Nomber showed bas on Delivery Date
    'ls_sql = "select Po_no from purchaseorder_master where Supplier_code='" & CboPart(0) & "' and (delivery_date >='" & Format(Tgl1, "YYYY-MM-DD") & "' and delivery_date <='" & Format(Tgl2, "YYYY-MM-DD") & "')"
    
    ' PO Nomber showed bas on Po Issue Date ( For Kawai )
    ls_sql = "select Po_no from purchaseorder_master where Supplier_code='" & CboPart(0) & "' and (PO_date >='" & Format(Tgl1, "YYYY-MM-DD") & "' and PO_date <='" & Format(Tgl2, "YYYY-MM-DD") & "')"
    
    
    If gb_AllowInputWithoutFix_PartReceiptSchedule = False Then
        ls_sql = ls_sql + " and isnull(fix_cls,'0')='1' "
    End If
    
    Set rsPO = Db.Execute(ls_sql)
    While Not rsPO.EOF
        CboPart(1).AddItem Trim$(rsPO!po_no)
        rsPO.MoveNext
    Wend
    CboPart(1).SelStart = 0
End If
Call Header
End Sub

Private Sub txtQty_LostFocus()
If IsNumeric(txtQty) = True Then
    txtQty.Text = Format(txtQty.Text, gs_formatQty)
Else
    txtQty.Text = Format(0, gs_formatQty)
End If
End Sub

Sub ProsesStock(nTipe As Byte, ItemCode As String, OldWHCode As String, NewWHCode As String, RecDate As String, QtyX As String, OldQty As String)
'## nTipe=1 ----> Insert
'## nTipe=2 ----> Update
'## nTipe=3 ----> Delete
Dim Sqlc As String, RsInvControl As New ADODB.Recordset
Dim RsWHS As Recordset, RSIS As Recordset, rswh As Recordset
Dim WHStock_ctrl As Boolean
Dim ItemStock_ctrl As Boolean


ItemStock_ctrl = False
WHStock_ctrl = False

QtyX = Format(QtyX, gs_formatQty)
OldQty = Format(OldQty, gs_formatQty)

Dim FlagU As Byte
        
'Proses Cek nya adalah dari warehouse dulu baru cek item nya
Select Case nTipe
    Case 1 'insert
        
            Set RsWHS = Db.Execute("Select stockcontrol_cls,WH_name from warehouse_master where WH_code='" & Trim$(OldWHCode) & "'")
                If Not RsWHS.EOF Then
                    If RsWHS!stockcontrol_cls <> "01" Then
                        'Stock tidak boleh di update
                        Exit Sub
                    Else
                        'Stock boleh di update
                        'Cek apakah item boleh  update stock atau tidak
                        Set RSIS = Db.Execute("Select StockControl_cls from item_master where item_code='" & Trim$(ItemCode) & "'") ' and (stockcontrol_cls='01')")
                        If Not RSIS.EOF Then
                            If Trim$(RSIS!stockcontrol_cls) <> "01" Then
                                'Stock tidak boleh di update
                                Exit Sub
                            End If
                            '## Update Stock Master
                            
                            If Trim$(CboRecCls) = "R" Then
                                updateStock Trim$(NewWHCode), ItemCode, Right(RecDate, 2), Mid(RecDate, 1, 4), CDbl(QtyX), 0
                            Else
                                Set rswh = Db.Execute("select * from stock_master with (nolock) where warehouse_code = '" & NewWHCode & "' and item_code = '" & ItemCode & "'")
                                If Not rswh.EOF Then
                                    updateStock Trim$(NewWHCode), ItemCode, Right(RecDate, 2), Mid(RecDate, 1, 4), 0, CDbl(QtyX)
                                Else
                                    updateStock Trim$(NewWHCode), ItemCode, Right(RecDate, 2), Mid(RecDate, 1, 4), CDbl(-(QtyX)), 0
                                End If
                                rswh.Close
                                Set rswh = Nothing
                            End If
                            
                            ItemStock_ctrl = True
                        Else
                            ItemStock_ctrl = False
                            Exit Sub
                        End If
                        RSIS.Close
                        Set RSIS = Nothing
                        
                    End If
                    WHStock_ctrl = True
                Else
                    WHStock_ctrl = False
                End If
                RsWHS.Close
                Set RsWHS = Nothing
        
    Case 2 'Update
            'Jika Update Cek Apakah WH_code nya berubah
            'Jika iya maka cek masing2 WH Code
            'Jika kedua WH Code tsb boleh update stock maka proses kedua WH COde tsb
            If OldWHCode <> NewWHCode Then 'WH_code berubah
                
                'Cek Apakah OldWH_code yg dipilih StockControl_cls nya ='01'
                Set RsWHS = Db.Execute("Select stockcontrol_cls,WH_name from warehouse_master where WH_code='" & Trim$(OldWHCode) & "'")
                If Not RsWHS.EOF Then
                    If RsWHS!stockcontrol_cls <> "01" Then
                        'Stock tidak boleh di update
                        'Exit Sub
                        GoTo terus:
                    Else
                        'Stock boleh di update
                        'Cek apakah item boleh  update stock atau tidak
                        Set RSIS = Db.Execute("Select StockControl_cls from item_master where item_code='" & Trim$(ItemCode) & "'") ' and (stockcontrol_cls='01')")
                        If Not RSIS.EOF Then
                            If Trim$(RSIS!stockcontrol_cls) <> "01" Then
                                'Stock tidak boleh di update
                                'Exit Sub
                                GoTo terus:
                            End If
                            '## Update Stock master
                            If Trim$(LblRecCls) = "R" Then
                                If Trim$(CboRecCls) = "R" Then
                                    updateStock Trim$(OldWHCode), ItemCode, Right(RecDate, 2), Mid(RecDate, 1, 4), 0, CDbl(OldQty)
                                Else
                                    updateStock Trim$(OldWHCode), ItemCode, Right(RecDate, 2), Mid(RecDate, 1, 4), 0, CDbl(OldQty)
                                End If
                            Else
                                If Trim$(CboRecCls) = "R" Then
                                    updateStock Trim$(OldWHCode), ItemCode, Right(RecDate, 2), Mid(RecDate, 1, 4), 0, (0 - CDbl(OldQty))
                                Else
                                    updateStock Trim$(OldWHCode), ItemCode, Right(RecDate, 2), Mid(RecDate, 1, 4), 0, (0 - CDbl(OldQty))
                                End If
                            End If
                            ItemStock_ctrl = True
                        Else
                            ItemStock_ctrl = False
                            Exit Sub
                        End If
                        RSIS.Close
                        Set RSIS = Nothing
                        
                        
                    End If
                    WHStock_ctrl = True
                Else
                    WHStock_ctrl = False
                End If
                RsWHS.Close
                Set RsWHS = Nothing
            
            End If
terus:
                
                Set RsWHS = Db.Execute("Select stockcontrol_cls,WH_name from warehouse_master where WH_code='" & Trim$(NewWHCode) & "'")
                If Not RsWHS.EOF Then
                    If RsWHS!stockcontrol_cls <> "01" Then
                        'Stock tidak boleh di update
                        Exit Sub
                    Else
                        'Stock boleh di update
                        'Cek apakah item boleh  update stock atau tidak
                        Set RSIS = Db.Execute("Select StockControl_cls from item_master where item_code='" & Trim$(ItemCode) & "'") ' and (stockcontrol_cls='01')")
                        If Not RSIS.EOF Then
                            If Trim$(RSIS!stockcontrol_cls) <> "01" Then
                                'Stock tidak boleh di update
                                Exit Sub
                            End If
                            '## Update Stock Master
                            If OldWHCode <> NewWHCode Then
                                
                                If Trim$(LblRecCls) = "R" Then
                                    If Trim$(CboRecCls) = "R" Then
                                        updateStock Trim$(NewWHCode), ItemCode, Right(RecDate, 2), Mid(RecDate, 1, 4), CDbl(QtyX), 0
                                    Else
                                        updateStock Trim$(NewWHCode), ItemCode, Right(RecDate, 2), Mid(RecDate, 1, 4), (0 - CDbl(QtyX)), 0
                                    End If
                                Else
                                    If Trim$(CboRecCls) = "R" Then
                                        updateStock Trim$(NewWHCode), ItemCode, Right(RecDate, 2), Mid(RecDate, 1, 4), CDbl(QtyX), 0
                                    Else
                                        updateStock Trim$(NewWHCode), ItemCode, Right(RecDate, 2), Mid(RecDate, 1, 4), (0 - CDbl(QtyX)), 0
                                    End If
                                End If
                            Else
                                'updateStock Trim$(NewWHCode), itemCode, Right(RecDate, 2), Mid(RecDate, 1, 4), CDbl(QtyX), CDbl(OldQty)
                                If Trim$(LblRecCls) = "R" Then
                                    If Trim$(CboRecCls) = "R" Then
                                        updateStock Trim$(NewWHCode), ItemCode, Right(RecDate, 2), Mid(RecDate, 1, 4), CDbl(QtyX), CDbl(OldQty)
                                    Else
                                        updateStock Trim$(NewWHCode), ItemCode, Right(RecDate, 2), Mid(RecDate, 1, 4), (0 - CDbl(OldQty)), 0
                                        updateStock Trim$(NewWHCode), ItemCode, Right(RecDate, 2), Mid(RecDate, 1, 4), (0 - CDbl(QtyX)), 0
                                    End If
                                Else
                                    If Trim$(CboRecCls) = "R" Then
                                        updateStock Trim$(NewWHCode), ItemCode, Right(RecDate, 2), Mid(RecDate, 1, 4), CDbl(OldQty), 0
                                        updateStock Trim$(NewWHCode), ItemCode, Right(RecDate, 2), Mid(RecDate, 1, 4), CDbl(QtyX), 0
                                    Else
                                        updateStock Trim$(NewWHCode), ItemCode, Right(RecDate, 2), Mid(RecDate, 1, 4), CDbl(OldQty), CDbl(QtyX)
                                    End If
                                End If
                            End If
                            ItemStock_ctrl = True
                        Else
                            ItemStock_ctrl = False
                            Exit Sub
                        End If
                        RSIS.Close
                        Set RSIS = Nothing
                        
                    End If
                    WHStock_ctrl = True
                Else
                    
                    WHStock_ctrl = False
                End If
                RsWHS.Close
                Set RsWHS = Nothing
            
            
       Case 3 'Delete
       
            Set RsWHS = Db.Execute("Select stockcontrol_cls,WH_name from warehouse_master where WH_code='" & Trim$(OldWHCode) & "'")
                If Not RsWHS.EOF Then
                    If RsWHS!stockcontrol_cls <> "01" Then
                        'Stock tidak boleh di update
                        Exit Sub
                    Else
                        'Stock boleh di update
                        'Cek apakah item boleh  update stock atau tidak
                        Set RSIS = Db.Execute("Select StockControl_cls from item_master where item_code='" & Trim$(ItemCode) & "'") ' and (stockcontrol_cls='01')")
                        If Not RSIS.EOF Then
                            If Trim$(RSIS!stockcontrol_cls) <> "01" Then
                                'Stock tidak boleh di update
                                Exit Sub
                            End If
                            '## Update Stock Master /Delete
                            
                            If Trim$(LblRecCls) = "R" Then
                                DeleteStock Trim$(NewWHCode), Trim$(ItemCode), Right(RecDate, 2), Mid(RecDate, 1, 4), CDbl(QtyX)
                            Else
                                DeleteStock Trim$(NewWHCode), Trim$(ItemCode), Right(RecDate, 2), Mid(RecDate, 1, 4), (0 - CDbl(QtyX))
                            End If
                            
                            ItemStock_ctrl = True
                        Else
                            ItemStock_ctrl = False
                            Exit Sub
                        End If
                        RSIS.Close
                        Set RSIS = Nothing
                        
                    End If
                    WHStock_ctrl = True
                Else
                    
                    WHStock_ctrl = False
                End If
                RsWHS.Close
                Set RsWHS = Nothing
       
End Select
End Sub

Sub updateStock(WHCode As String, ItemCode As String, RecMonth As String, RecYear As String, QtyX As Double, OldQty As Double)
Dim RsI As Recordset
Dim DBi As New Connection
Dim FixMonth As String
Dim FixYear As String

Dim LM_P As Double, LM_S As Double, LM_RJ As Double, LM_I As Double
Dim TM_P As Double, TM_S As Double, TM_RJ As Double, TM_I As Double
Dim NM_P As Double, NM_S As Double, NM_RJ As Double, NM_I As Double

'DBi.ConnectionTimeout = 0
'DBi.Open Db.ConnectionString

'DBi.BeginTrans
'Db.BeginTrans
If GetLastMonthStock = "" Then LblErr = DisplayMsg(4019): Exit Sub

FixMonth = Right(GetLastMonthStock, 2)
FixYear = Mid(GetLastMonthStock, 1, 4)

Set RsI = New Recordset
RsI.Open "select * from stock_master (updlock) where " & _
      " warehouse_code='" & Trim(WHCode) & "' and " & _
      " item_code='" & Trim(ItemCode) & "'", Db, adOpenDynamic, adLockOptimistic
If RsI.EOF Then
  RsI.AddNew
  'Receipt
  RsI!Item_Code = Trim$(ItemCode)
  RsI!Warehouse_Code = Trim$(WHCode)
  RsI!lm_premonth = "0"
  RsI!tm_premonth = "0"
  RsI!nm_premonth = "0"
  
  RsI!lm_supply = "0"
  RsI!tm_supply = "0"
  RsI!nm_supply = "0"
  RsI!lm_lossreject = "0"
  RsI!tm_lossreject = "0"
  RsI!nm_lossreject = "0"
  
  RsI!lm_current = "0"
  RsI!tm_current = "0"
  RsI!nm_current = "0"
  
  RsI!lm_inventory = Null
  RsI!tm_inventory = Null
  RsI!nm_inventory = Null
  
  If Val(FixYear) = Val(RecYear) Then
    RsI!lm_receipt = IIf(Val(RecMonth) = Val(FixMonth), QtyX, 0)
    RsI!tm_receipt = IIf(Val(RecMonth) = Val(FixMonth) + 1, QtyX, 0)
    RsI!nm_receipt = IIf(Val(RecMonth) = Val(FixMonth) + 2, QtyX, 0)
    'Current
    RsI!lm_current = Val(RsI!lm_premonth) + Val(RsI!lm_receipt) - Val(RsI!lm_supply) - Val(RsI!lm_lossreject)
    RsI!tm_current = Val(RsI!tm_premonth) + Val(RsI!tm_receipt) - Val(RsI!tm_supply) - Val(RsI!tm_lossreject)
    'Next Proses
    RsI!nm_premonth = RsI!tm_current
    RsI!nm_current = Val(RsI!nm_premonth) + Val(RsI!nm_receipt) - Val(RsI!nm_supply) - Val(RsI!nm_lossreject)
  ElseIf Val(FixYear) < Val(RecYear) Then
    RsI!tm_receipt = RsI!tm_receipt + IIf((Val(RecMonth) + 12) - Val(FixMonth) = 1, QtyX, 0)
    RsI!nm_receipt = RsI!nm_receipt + IIf((Val(RecMonth) + 12) - Val(FixMonth) = 2, QtyX, 0)
    'Current
    RsI!lm_current = Val(RsI!lm_premonth) + Val(RsI!lm_receipt) - Val(RsI!lm_supply) - Val(RsI!lm_lossreject)
    RsI!tm_current = Val(RsI!tm_premonth) + Val(RsI!tm_receipt) - Val(RsI!tm_supply) - Val(RsI!tm_lossreject)
    RsI!nm_current = Val(RsI!nm_premonth) + Val(RsI!nm_receipt) - Val(RsI!nm_supply) - Val(RsI!nm_lossreject)
  End If
  RsI!Last_Update = Now
  RsI!last_user = userLogin
  RsI.update
Else
  'LM Null offset
    If IsNull(RsI!lm_premonth) Then
        LM_P = 0
    Else
        LM_P = CDbl(RsI!lm_premonth)
    End If
    If IsNull(RsI!lm_supply) Then
        LM_S = 0
    Else
        LM_S = CDbl(RsI!lm_supply)
    End If
    If IsNull(RsI!lm_lossreject) Then
        LM_RJ = 0
    Else
        LM_RJ = CDbl(RsI!lm_lossreject)
    End If
    'TM Null Offset
    If IsNull(RsI!tm_premonth) Then
        TM_P = 0
    Else
        TM_P = CDbl(RsI!tm_premonth)
    End If
    If IsNull(RsI!tm_supply) Then
        TM_S = 0
    Else
        TM_S = CDbl(RsI!tm_supply)
    End If
    If IsNull(RsI!tm_lossreject) Then
        TM_RJ = 0
    Else
        TM_RJ = CDbl(RsI!tm_lossreject)
    End If
    'NM Null Offset
    If IsNull(RsI!nm_premonth) Then
        NM_P = 0
    Else
        NM_P = CDbl(RsI!nm_premonth)
    End If
    If IsNull(RsI!nm_supply) Then
        NM_S = 0
    Else
        NM_S = CDbl(RsI!nm_supply)
    End If
    If IsNull(RsI!nm_lossreject) Then
        NM_RJ = 0
    Else
        NM_RJ = CDbl(RsI!nm_lossreject)
    End If
    

  If Val(FixYear) = Val(RecYear) Then
    'If IsNull(RsI!LM_supply) Then
        
    RsI!lm_receipt = RsI!lm_receipt - IIf(Val(RecMonth) = Val(FixMonth), (OldQty - QtyX), 0)
    RsI!tm_receipt = RsI!tm_receipt - IIf(Val(RecMonth) = Val(FixMonth) + 1, (OldQty - QtyX), 0)
    RsI!nm_receipt = RsI!nm_receipt - IIf(Val(RecMonth) = Val(FixMonth) + 2, (OldQty - QtyX), 0)
    'Current
    RsI!lm_current = (LM_P + Val(RsI!lm_receipt)) - LM_S - LM_RJ
    RsI!tm_current = (TM_P + Val(RsI!tm_receipt)) - TM_S - TM_RJ
    'Next Proses
    RsI!nm_premonth = RsI!tm_current
    NM_P = RsI!nm_premonth
    RsI!nm_current = (NM_P + Val(RsI!nm_receipt)) - NM_S - NM_RJ
  ElseIf Val(FixYear) < Val(RecYear) Then
    RsI!tm_receipt = RsI!tm_receipt - IIf((Val(RecMonth) + 12) - Val(FixMonth) = 1, (OldQty - QtyX), 0)
    RsI!nm_receipt = RsI!nm_receipt - IIf((Val(RecMonth) + 12) - Val(FixMonth) = 2, (OldQty - QtyX), 0)
    'Current
    RsI!lm_current = (LM_P + Val(RsI!lm_receipt)) - LM_S - LM_RJ
    RsI!tm_current = (TM_P + Val(RsI!tm_receipt)) - TM_S - TM_RJ
    
    RsI!nm_premonth = RsI!tm_current
    NM_P = RsI!nm_premonth
    RsI!nm_current = (NM_P + Val(RsI!nm_receipt)) - NM_S - NM_RJ
  End If
  RsI!Last_Update = Now
  RsI!last_user = userLogin
  RsI.update
End If
       
       
On Error GoTo erri
    'Db.CommitTrans
    
    
    RsI.Close
    Set RsI = Nothing
    'Db.Close
    'Set DBi = Nothing
Exit Sub
erri:
    'DBi.RollbackTrans
    
    RsI.Close
    Set RsI = Nothing
    DBi.Close
    Set DBi = Nothing
End Sub
Sub DeleteStock(WHCode As String, ItemCode As String, RecMonth As String, RecYear As String, QtyX As Double)
Dim RsI As Recordset
Dim DBi As New Connection
Dim FixMonth As String
Dim FixYear As String
Dim LM_P As Double, LM_S As Double, LM_RJ As Double, LM_I As Double
Dim TM_P As Double, TM_S As Double, TM_RJ As Double, TM_I As Double
Dim NM_P As Double, NM_S As Double, NM_RJ As Double, NM_I As Double


DBi.ConnectionTimeout = 0
DBi.Open Db.ConnectionString

DBi.BeginTrans

If GetLastMonthStock = "" Then LblErr = DisplayMsg(4019): Exit Sub

FixMonth = Right(GetLastMonthStock, 2)
FixYear = Mid(GetLastMonthStock, 1, 4)

Set RsI = New Recordset
RsI.Open "select * from stock_master (updlock) where " & _
      " warehouse_code='" & Trim(WHCode) & "' and " & _
      " item_code='" & Trim(ItemCode) & "'", Db, adOpenDynamic, adLockOptimistic
      
If Not RsI.EOF Then
    'LM Null offset
    If IsNull(RsI!lm_premonth) Then
        LM_P = 0
    Else
        LM_P = CDbl(RsI!lm_premonth)
    End If
    If IsNull(RsI!lm_supply) Then
        LM_S = 0
    Else
        LM_S = CDbl(RsI!lm_supply)
    End If
    If IsNull(RsI!lm_lossreject) Then
        LM_RJ = 0
    Else
        LM_RJ = CDbl(RsI!lm_lossreject)
    End If
    'TM Null Offset
    If IsNull(RsI!tm_premonth) Then
        TM_P = 0
    Else
        TM_P = CDbl(RsI!tm_premonth)
    End If
    If IsNull(RsI!tm_supply) Then
        TM_S = 0
    Else
        TM_S = CDbl(RsI!tm_supply)
    End If
    If IsNull(RsI!tm_lossreject) Then
        TM_RJ = 0
    Else
        TM_RJ = CDbl(RsI!tm_lossreject)
    End If
    'NM Null Offset
    If IsNull(RsI!nm_premonth) Then
        NM_P = 0
    Else
        NM_P = CDbl(RsI!nm_premonth)
    End If
    If IsNull(RsI!nm_supply) Then
        NM_S = 0
    Else
        NM_S = CDbl(RsI!nm_supply)
    End If
    If IsNull(RsI!nm_lossreject) Then
        NM_RJ = 0
    Else
        NM_RJ = CDbl(RsI!nm_lossreject)
    End If
    
  If Val(FixYear) = Val(RecYear) Then
    RsI!lm_receipt = RsI!lm_receipt - IIf(Val(RecMonth) = Val(FixMonth), QtyX, 0)
    RsI!tm_receipt = RsI!tm_receipt - IIf(Val(RecMonth) = Val(FixMonth) + 1, QtyX, 0)
    RsI!nm_receipt = RsI!nm_receipt - IIf(Val(RecMonth) = Val(FixMonth) + 2, QtyX, 0)
    'Current
    RsI!lm_current = (LM_P + Val(RsI!lm_receipt)) - LM_S - LM_RJ
    RsI!tm_current = (TM_P + Val(RsI!tm_receipt)) - TM_S - TM_RJ
    'Next Proses
    RsI!nm_premonth = RsI!tm_current
    NM_P = RsI!nm_premonth
    RsI!nm_current = (NM_P + Val(RsI!nm_receipt)) - NM_S - NM_RJ
  ElseIf Val(FixYear) < Val(RecYear) Then
    RsI!tm_receipt = RsI!tm_receipt - IIf((Val(RecMonth) + 12) - Val(FixMonth) = 1, QtyX, 0)
    RsI!nm_receipt = RsI!nm_receipt - IIf((Val(RecMonth) + 12) - Val(FixMonth) = 2, QtyX, 0)
    'Current
    RsI!lm_current = (LM_P + Val(RsI!lm_receipt)) - LM_S - LM_RJ
    RsI!tm_current = (TM_P + Val(RsI!tm_receipt)) - TM_S - TM_RJ
    'Next Proses
    RsI!nm_premonth = RsI!tm_current
    RsI!nm_current = (NM_P + Val(RsI!nm_receipt)) - NM_S - NM_RJ
  End If
  RsI!Last_Update = Now
  RsI!last_user = userLogin
  RsI.update
End If
       
On Error GoTo erri
    DBi.CommitTrans
    
    RsI.Close
    Set RsI = Nothing
    DBi.Close
    Set DBi = Nothing
Exit Sub
erri:
    DBi.RollbackTrans
    
    RsI.Close
    Set RsI = Nothing
    DBi.Close
    Set DBi = Nothing
End Sub

Private Sub txtRemarks_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub TxtSj_Change()
    txtRegisterNo.Text = ""
End Sub

Private Sub TxtSj_GotFocus()
    l_SJNo = TxtSj.Text
End Sub

Private Sub TxtSj_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Sub browseprice(itmCode As String, Tgl As String, TradeCode As String, Curr As String)
Dim sql2 As String
Dim rs2 As New Recordset
Dim i As Integer, idxShow As Byte, blnpricez As Boolean
Dim tmpprice As String

    sql2 = "select trade_code, priority_cls, currency_code, price, unit_cls from price_master where " & _
           "item_code='" & Trim$(itmCode) & "' and price_cls='01' and (trade_code='" & TradeCode & _
           "' or trade_code='000000') and start_date<='" & Format(Tgl, "yyyymmdd") & "' and end_date>='" & _
           Format(Tgl, "yyyymmdd") & "' and currency_code ='" & Curr & "' order by trade_code desc, priority_cls desc"
    Set rs2 = Db.Execute(sql2)


    With cboprice
        tmpprice = cboprice
        .Text = ""
        .clear
        .columnCount = 4
        .ColumnWidths = "90pt;70pt;0pt;0pt"
        .ListWidth = 160
        .ListRows = 4
idxShow = 0
        i = 0
        blnpricez = False
        If Not rs2.EOF Then
            Do While Not rs2.EOF
                .AddItem
                .List(i, 0) = Format(Trim(rs2("price")), gs_formatPrice)
                If rs2("trade_code") = "000000" Then
                  .List(i, 1) = "Common " & Trim(rs2("priority_cls"))
                Else
                  .List(i, 1) = "Priority " & Trim(rs2("priority_cls"))
                  If blnshow Then idxShow = i: blnshow = False: blnpricez = True
                End If
                .List(i, 2) = Trim(rs2("Currency_Code"))
                .List(i, 3) = Trim(rs2("unit_cls"))
    
                rs2.MoveNext
                i = i + 1
    
            Loop
            cboprice.Text = tmpprice
            If idxShow >= 0 And blnpricez Then cboprice.Text = cboprice.List(idxShow, 0)
            End If
    End With
    
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    LblErr.Caption = ErrMsg
End If
End Sub




Private Sub up_SendEmail()
'    'Recordset
'    Dim rsmailconfig    As New ADODB.Recordset
'    Dim rsHeader        As New ADODB.Recordset
'    Dim rsdetail        As New ADODB.Recordset
'
'    'Excel Application
'    Dim xlapp As New Excel.application
'
'    Dim ls_SupplierCodeGrid As String
'    Dim ls_TicketDateGrid   As String
'
'    'Variable
'    Dim ls_sql As String
'    Dim li_Row As Integer
'    Dim ls_smtpserver As String
'    Dim li_smtpport     As Double
'    Dim li_smtptimeout  As Double
'    Dim ls_sender       As String
'    Dim ls_username     As String
'    Dim ls_password     As String
'    Dim li_timer        As Double
'    Dim ls_subject      As String
'    Dim ls_cc           As String
'    Dim ls_bcc          As String
'    Dim ls_message      As String
'    Dim ls_recipient    As String
'    Dim ls_SupplierCode As String
'    Dim ls_TicketNo     As String
'    Dim ls_SampleNo     As String
'    Dim ls_ticketdate   As String
'    Dim ls_ItemCode     As String
'    Dim li_wet          As Double
'    Dim li_drc          As Double
'    Dim li_dry          As Double
'    Dim li_price        As Double
'    Dim li_amount       As Double
'    Dim ls_log          As String
'
'    Dim Idx                 As Integer
'    Dim li_totalwet         As Double, li_totaldry          As Double
'    Dim li_totalamout       As Double, li_totaladvanced     As Double
'    Dim li_totaltaxamount   As Double, li_totalfinalpaid    As Double
'
'    Dim strFile         As String
'
'
'    Dim strAttachFile   As String
'    Const clrGrey = 15
'
'    Dim ls_desc As String
'
'    On Local Error GoTo errHandler
'
'        'Set Timer
'    ls_sql = " select smtp_server, port_number, smtp_timeout, email_address, user_email, pass_email, timer, subject, " & vbCrLf & _
'            "       ccemail_address, bccemail_address,mail_header,Mail_Content,Mail_Footer " & vbCrLf & _
'            " from email_config "
'
'    If rsmailconfig.State <> adStateClosed Then rsmailconfig.Close
'    Set rsmailconfig = Db.Execute(ls_sql)
'
'    If Not rsmailconfig.EOF Then
'
'        ls_smtpserver = Trim(rsmailconfig!smtp_server & "")
'        li_smtpport = IIf(IsNull(rsmailconfig!port_number), 0, rsmailconfig!port_number)
'        li_smtptimeout = IIf(IsNull(rsmailconfig!smtp_timeout), 0, rsmailconfig!smtp_timeout)
'        ls_sender = Trim(rsmailconfig!Email_Address & "")
'        ls_username = Trim(rsmailconfig!user_email & "")
'        ls_password = Trim(rsmailconfig!pass_email & "")
'        li_timer = IIf(IsNull(rsmailconfig!Timer), 0, rsmailconfig!Timer)
'        ls_subject = Trim(rsmailconfig!Subject & "")
'        ls_cc = Trim(rsmailconfig!ccemail_address & "")
'        ls_bcc = Trim(rsmailconfig!bccemail_address & "")
'        ls_message = Trim(rsmailconfig!mail_header & "") & vbCrLf & vbCrLf
'        ls_message = ls_message & Trim(rsmailconfig!Mail_Content & "") & vbCrLf & vbCrLf
'        ls_message = ls_message & "Part No" & vbTab & vbTab & ":" & Trim(CboItem.Text) & vbCrLf
'        ls_message = ls_message & "Part Name" & vbTab & ":" & Trim(LblItemName.Text) & vbCrLf
'        ls_message = ls_message & "Qty" & vbTab & vbTab & ":" & Trim(txtQty.Text) & vbCrLf
'        ls_message = ls_message & "Vendor" & vbTab & vbTab & ":" & Trim(CboPart(0).Text) & " - " & Trim(Lblsupp.Text) & vbCrLf
'        ls_message = ls_message & "Surat Jalan" & vbTab & ":" & Trim(TxtSj.Text) & vbCrLf & vbCrLf
'        ls_message = ls_message & "Kepada yang berkepentingan Mohon Segera Pastikan Barang Tersebut." & vbCrLf
'        ls_message = ls_message & "Terima Kasih" & vbCrLf & vbCrLf & vbCrLf
'        ls_message = ls_message & "Note :" & vbCrLf
'        ls_message = ls_message & Trim(rsmailconfig!Mail_Footer & "")
'
'
'
'    End If
'
'    rsmailconfig.Close
'    Set rsmailconfig = Nothing
'
'
'      '========================================================================
'      'start send email function
'      '========================================================================
'
'      Call up_cdoSendEmail(ls_smtpserver, li_smtpport, li_smtptimeout, ls_username, ls_password, ls_sender, ls_recipient, ls_subject & ls_TicketNo, ls_message, ls_cc, ls_bcc, strAttachFile)
'
'      '========================================================================
'      'end send email function
'      '========================================================================
'
'
'ErrExit:
'    Exit Sub
'errHandler:
'    'Call SaveToTxt("SendEmail_ErrorLog", " Ticket No : " & ls_TicketNo & " ( " & ls_ticketdate & " ) " & err.Description, CurrDate)
'    err.clear
'    Resume ErrExit
    
End Sub



Public Sub up_cdoSendEmail(ByVal ls_smtpserver As String, ByVal li_smtpport As Double, ByVal li_smtptimeout As Double, _
                    ByVal ls_username As String, ByVal ls_password As String, ByVal ls_sender As String, _
                    ByVal ls_recipient As String, ByVal ls_subject As String, ByVal ls_body As String, _
                    Optional ByVal ls_cc As String, Optional ByVal ls_bcc As String, Optional ByVal ls_attachment As String)
    
    Dim cdoMsg As New CDO.message
    Dim cdoConf As New CDO.Configuration
    Dim Flds
    Dim attachment
   
    On Error GoTo errHandler
    
    Set Flds = cdoConf.Fields
       
    With Flds
        .Item(cdoSendUsingMethod) = cdoSendUsingPort
        .Item(cdoSMTPServer) = ls_smtpserver
        '.Item(cdoSMTPServerPort) = IIf(li_smtpport = 0, 25, li_smtpport)
        .Item(cdoSMTPServerPort) = li_smtpport
        .Item(cdoSMTPConnectionTimeout) = IIf(li_smtptimeout = 0, 10, li_smtptimeout)
        .Item(cdoSMTPAuthenticate) = 1 'default
        .Item(cdoSendUserName) = ls_username
        .Item(cdoSendPassword) = ls_password
        .Item(cdoSMTPUseSSL) = 1
        
        
        
        
        
        
    .update
    End With
   
   
    ' Apply the settings to the message.
    With cdoMsg
        Set .Configuration = cdoConf
        .To = ls_cc
        .From = ls_sender
        .Subject = ls_subject
        .TextBody = ls_body
        
'        If Not colAttachments Is Nothing Then
'            For Each attachment In colAttachments
'                .AddAttachment attachment
'            Next
'        End If
        'If ls_cc <> "" Then .CC = ls_cc
        If ls_bcc <> "" Then .BCC = ls_bcc
        .Send
    End With
    LblErr.Caption = "Send Successful"
    
   
ErrExit:
    Set cdoMsg = Nothing
    Set cdoConf = Nothing
    Set Flds = Nothing
    
    Exit Sub

errHandler:
LblErr.Caption = err.Description
'Call SaveToTxt("SendEmail_ErrorLog", "SEND MAIL - " & err.number & " : " & err.Description, CurrDate)
GoTo ErrExit
    
End Sub


Public Function up_Validasi_Item(Item_Code As String) As Boolean
Dim sqlcek As String

Dim RsCekItem As New ADODB.Recordset

    sqlcek = "select tOP 1 Item_Code From Stock_Master WHERE Item_Code='" & Item_Code & "' " & vbCrLf & _
            " union select tOP 1 Item_Code From Stock_History WHERE Item_Code='" & Item_Code & "' "
    Set RsCekItem = Db.Execute(sqlcek)
    
    If Not RsCekItem.EOF Then
        up_Validasi_Item = True
    Else
        up_Validasi_Item = False
    End If

End Function

Private Sub GetBCType()
Dim RS As New ADODB.Recordset
    RS.Open "select Type_BC from Trade_Master where Trade_Code = '" & CboPart(0).Text & "' ", Db, adOpenForwardOnly, adLockReadOnly, adCmdText
    If RS.EOF = False Then
        CbotypeBC.Text = IIf(IsNull(Trim(RS!Type_BC)), "", Trim(RS!Type_BC))
    End If
    RS.Close
End Sub

Public Function up_ValidateSuratJalan() As Boolean
Dim sqlcek As String
Dim rsCek As New ADODB.Recordset
    
    DateActual = Format(TglReceipt, "MM/DD/YYYY")
    
    sqlcek = " SELECT DISTINCT * FROM Part_Receipt WHERE PO_No='" & CboPart(1).Text & "' " & vbCrLf & _
            " AND SuratJalan_No='" & TxtSj.Text & "' "
    Set rsCek = Db.Execute(sqlcek)
    
    If Not rsCek.EOF Then
    
        sqlcek = " SELECT DISTINCT Receipt_Date FROM Part_Receipt WHERE PO_No='" & CboPart(1).Text & "' " & vbCrLf & _
                 " AND SuratJalan_No='" & TxtSj.Text & "' "
        Set rsCek = Db.Execute(sqlcek)
        
        If Not rsCek.EOF Then
            receiptDate = Trim(rsCek!Receipt_Date & "")
        End If
        
        If receiptDate = DateActual Then
            validate = True
        Else
            validate = False
        End If
    Else
        validate = True
    End If
    
End Function

Private Sub TxtSj_LostFocus()
    Dim rsGetNoSeri As New ADODB.Recordset

    up_GetReceiptDate
    
    If CbotypeBC.Text <> "4.0" And CbotypeBC.Text <> "2.6.2" And CbotypeBC.Text <> "2.3" Then
        If l_SJNo <> TxtSj.Text Then
            If Trim(TxtSj.Text) <> "" Then
                sql = "EXEC dbo.sp_GetNoRegister '" & TglReceipt.Value & "', 'R', '" & userLogin & "' "
                        
                If rsGetNoSeri.State <> adStateClosed Then rsGetNoSeri.Close
                rsGetNoSeri.Open sql, Db, adOpenForwardOnly, adLockReadOnly
                
                If Not rsGetNoSeri.EOF Then
                   txtRegisterNo.Text = rsGetNoSeri.Fields("No_Register")
                End If
            End If
        End If
    Else
        txtRegisterNo.Text = ""
    End If
End Sub

Private Sub up_GetReceiptDate()
Dim rsGetReciptDate As New ADODB.Recordset

    If Trim(TxtSj.Text) <> "" Then
        sql = "SELECT Receipt_Date, ISNULL(BC40_Date, GETDATE())BC40_Date, ISNULL(BC40_No,'')BC40_No, ISNULL(BC_Type,'')BC_Type FROM dbo.Part_Receipt WHERE SuratJalan_No= '" & Trim(TxtSj.Text) & "' "
                
        If rsGetReciptDate.State <> adStateClosed Then rsGetReciptDate.Close
        rsGetReciptDate.Open sql, Db, adOpenForwardOnly, adLockReadOnly
        
        If Not rsGetReciptDate.EOF Then
           TglReceipt.Value = rsGetReciptDate.Fields("Receipt_Date")
           DtBCDate.Value = rsGetReciptDate.Fields("BC40_Date")
           txtBC40.Text = rsGetReciptDate.Fields("BC40_No")
           CbotypeBC.Text = rsGetReciptDate.Fields("BC_Type")
        End If
                
    End If
    
End Sub

Private Sub up_GetNoSeri(pItemCode As String)
Dim rsGetNo As New ADODB.Recordset

    If Trim(CbotypeBC.Text) = "4.0" Then
'        sql = " SELECT ISNULL(MAX(CAST(No_Seri AS INT)), 0) + 1 No_Seri  " & vbCrLf & _
'                " FROM Part_Receipt  " & vbCrLf & _
'                " WHERE SuratJalan_No ='" & Trim(TxtSj.Text) & "' AND Receipt_Date = '" & Format(TglReceipt.Value, "yyyy-mm-dd") & "' "
            
            sql = "EXEC dbo.sp_PartReceipt_GetNoSeri @SJNo = '" & Trim(TxtSj.Text) & "', " & vbCrLf & _
                "  @ReceiptDate = '" & Format(TglReceipt.Value, "yyyy-mm-dd") & "', " & vbCrLf & _
                "  @ItemCode ='" & pItemCode & "' "
                
                If rsGetNo.State <> adStateClosed Then rsGetNo.Close
                rsGetNo.Open sql, Db, adOpenForwardOnly, adLockReadOnly
                
                If Not rsGetNo.EOF Then
                   txtNoSeri.Text = Trim(rsGetNo.Fields("No_Seri"))
                End If
    ElseIf Trim(CbotypeBC.Text) = "2.3" Then
    
        sql = "EXEC dbo.sp_PartReceipt_GetNoSeri @SJNo = '" & Trim(TxtSj.Text) & "', " & vbCrLf & _
                "  @ReceiptDate = '" & Format(TglReceipt.Value, "yyyy-mm-dd") & "', " & vbCrLf & _
                "  @ItemCode ='" & pItemCode & "' "
                        
                If rsGetNo.State <> adStateClosed Then rsGetNo.Close
                rsGetNo.Open sql, Db, adOpenForwardOnly, adLockReadOnly
                
                If Not rsGetNo.EOF Then
                   txtNoSeri.Text = Trim(rsGetNo.Fields("No_Seri"))
                End If
            CboRecCls.Text = "R"
    Else
         sql = "EXEC dbo.sp_PartReceipt_GetNoSeri @SJNo = '" & Trim(TxtSj.Text) & "', " & vbCrLf & _
                "  @ReceiptDate = '" & Format(TglReceipt.Value, "yyyy-mm-dd") & "', " & vbCrLf & _
                "  @ItemCode ='" & pItemCode & "' "
                
                If rsGetNo.State <> adStateClosed Then rsGetNo.Close
                rsGetNo.Open sql, Db, adOpenForwardOnly, adLockReadOnly
                
                If Not rsGetNo.EOF Then
                   txtNoSeri.Text = Trim(rsGetNo.Fields("No_Seri"))
                End If
    End If
    
End Sub
