VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_part_supply 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Parts (Material) Supply [Unscheduled]"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15120
   Icon            =   "frm_part_supply.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtDONo 
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
      Left            =   12600
      MaxLength       =   100
      TabIndex        =   81
      Tag             =   "TTFF*/"
      Top             =   7440
      Width           =   2325
   End
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
      Left            =   10560
      MaxLength       =   4
      TabIndex        =   79
      Tag             =   "TTFF*/"
      Top             =   7485
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
      Left            =   7800
      MaxLength       =   25
      TabIndex        =   77
      Tag             =   "TTFF*/"
      Top             =   7485
      Width           =   1725
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
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   74
      TabStop         =   0   'False
      Top             =   7440
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
      Left            =   3000
      MaxLength       =   25
      TabIndex        =   73
      TabStop         =   0   'False
      Top             =   7485
      Width           =   2175
   End
   Begin VB.CommandButton cmd_print 
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
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   9960
      Width           =   1125
   End
   Begin MSComCtl2.DTPicker DTBCdate 
      Height          =   315
      Left            =   9210
      TabIndex        =   56
      Top             =   8850
      Width           =   1455
      _ExtentX        =   2566
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
      Format          =   129171459
      CurrentDate     =   41080
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
      Left            =   6810
      MaxLength       =   25
      TabIndex        =   54
      Top             =   8880
      Width           =   1530
   End
   Begin VB.TextBox txtSJNo 
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
      Left            =   1020
      MaxLength       =   20
      TabIndex        =   9
      Top             =   8880
      Width           =   2670
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FDDFE3&
      Caption         =   "Sort By Description"
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
      Left            =   4230
      TabIndex        =   44
      Top             =   10200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmd_cancel 
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
      Left            =   11370
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   9960
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
      Left            =   12585
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   9960
      Width           =   1125
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   7320
      TabIndex        =   35
      Text            =   "Text1"
      Top             =   10080
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1950
      Left            =   300
      TabIndex        =   22
      Top             =   855
      Width           =   14625
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
         Left            =   11640
         TabIndex        =   67
         Top             =   1440
         Width           =   1755
      End
      Begin VB.TextBox txtqty 
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
         Left            =   11640
         TabIndex        =   66
         Top             =   675
         Width           =   1755
      End
      Begin VB.TextBox txtno 
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
         Left            =   11640
         TabIndex        =   65
         Top             =   240
         Width           =   1755
      End
      Begin VB.CommandButton cmd_search 
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
         Left            =   3930
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1463
         Width           =   1125
      End
      Begin VB.TextBox txt_toadd 
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
         Left            =   5760
         MaxLength       =   6
         TabIndex        =   32
         TabStop         =   0   'False
         Text            =   "addres"
         Top             =   1440
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.TextBox txt_fromadd 
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
         Left            =   5280
         MaxLength       =   6
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   "addres"
         Top             =   1440
         Visible         =   0   'False
         Width           =   1155
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   330
         Left            =   2340
         TabIndex        =   3
         Top             =   1485
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
         Format          =   129171459
         CurrentDate     =   37867
      End
      Begin MSForms.ComboBox CboDelivery 
         Height          =   315
         Left            =   11640
         TabIndex        =   68
         Top             =   1080
         Width           =   2415
         VariousPropertyBits=   746604571
         MaxLength       =   35
         DisplayStyle    =   3
         Size            =   "4260;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PO_No"
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
         Left            =   10410
         TabIndex        =   64
         Top             =   1560
         Width           =   570
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delivery To"
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
         Left            =   10410
         TabIndex        =   63
         Top             =   1200
         Width           =   990
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QTY Set"
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
         Left            =   10440
         TabIndex        =   62
         Top             =   750
         Width           =   690
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No"
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
         Left            =   10440
         TabIndex        =   61
         Top             =   308
         Width           =   225
      End
      Begin MSForms.ComboBox cbo_Replacement 
         Height          =   330
         Left            =   8760
         TabIndex        =   51
         Top             =   225
         Width           =   1020
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "1799;582"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         Value           =   "cbo_Replacement"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Warehouse Code"
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
         Left            =   7110
         TabIndex        =   50
         Top             =   1005
         Width           =   1470
      End
      Begin VB.Line Line5 
         Visible         =   0   'False
         X1              =   10410
         X2              =   13650
         Y1              =   990
         Y2              =   990
      End
      Begin VB.Label lbl_ReplacementWarehouseCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lbl_ReplacementWarehouseCode"
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
         Left            =   10410
         TabIndex        =   49
         Top             =   735
         Visible         =   0   'False
         Width           =   2820
      End
      Begin MSForms.ComboBox cbo_ReplacementWarehouseCode 
         Height          =   330
         Left            =   8760
         TabIndex        =   48
         Top             =   667
         Width           =   1500
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2646;582"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         Value           =   "cbo_ReplacementWarehouseCode"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Replacement"
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
         Left            =   7110
         TabIndex        =   33
         Top             =   735
         Width           =   1110
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Replacement"
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
         Left            =   7110
         TabIndex        =   30
         Top             =   293
         Width           =   1110
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Warehouse CD"
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
         Left            =   225
         TabIndex        =   29
         Top             =   293
         Width           =   1785
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Location CD"
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
         Left            =   225
         TabIndex        =   28
         Top             =   735
         Width           =   1305
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supply Cls"
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
         Left            =   225
         TabIndex        =   27
         Top             =   1133
         Width           =   915
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supply Date"
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
         Left            =   225
         TabIndex        =   26
         Top             =   1553
         Width           =   1050
      End
      Begin MSForms.ComboBox cbo_warehouse 
         Height          =   330
         Left            =   2340
         TabIndex        =   0
         Top             =   225
         Width           =   1500
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2646;582"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         Value           =   "cbo_warehouse"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cbo_location 
         Height          =   330
         Left            =   2340
         TabIndex        =   1
         Top             =   667
         Width           =   1500
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2646;582"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         Value           =   "cbo_location"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cbo_supply 
         Height          =   330
         Left            =   2340
         TabIndex        =   2
         Top             =   1065
         Width           =   780
         VariousPropertyBits=   746604571
         MaxLength       =   2
         DisplayStyle    =   3
         Size            =   "1376;582"
         ShowDropButtonWhen=   2
         Value           =   "cbo_supply"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lbl_warehouse 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lbl_warehouse"
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
         Left            =   3930
         TabIndex        =   25
         Top             =   315
         Width           =   3210
      End
      Begin VB.Label lbl_location 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lbl_location"
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
         Left            =   3930
         TabIndex        =   24
         Top             =   735
         Width           =   2490
      End
      Begin VB.Line Line1 
         X1              =   3930
         X2              =   6990
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Line Line2 
         X1              =   3930
         X2              =   6990
         Y1              =   990
         Y2              =   990
      End
      Begin VB.Line Line3 
         X1              =   3930
         X2              =   6990
         Y1              =   1395
         Y2              =   1395
      End
      Begin VB.Label lbl_supply 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lbl_supply"
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
         Left            =   3930
         TabIndex        =   23
         Top             =   1125
         Width           =   855
      End
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
      Left            =   330
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9960
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
      Left            =   13785
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9960
      Width           =   1125
   End
   Begin VB.TextBox txt_remarks 
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
      Left            =   11640
      TabIndex        =   10
      Top             =   8880
      Width           =   3300
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   585
      Left            =   330
      TabIndex        =   17
      Top             =   9270
      Width           =   14595
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
         Left            =   90
         TabIndex        =   18
         Top             =   240
         Width           =   14235
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FDDFE3&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   825
      Left            =   330
      TabIndex        =   15
      Top             =   7920
      Width           =   14625
      Begin VB.TextBox txtCurrentQty 
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
         Height          =   285
         Left            =   7200
         TabIndex        =   72
         Top             =   480
         Width           =   840
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
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
         Height          =   285
         Left            =   9060
         Locked          =   -1  'True
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   480
         Width           =   705
      End
      Begin VB.CommandButton cmdBrowser 
         Caption         =   "..."
         Height          =   300
         Left            =   2070
         TabIndex        =   47
         Top             =   465
         Width           =   300
      End
      Begin VB.TextBox txt_desc 
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
         Height          =   285
         Left            =   4650
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   480
         Width           =   2445
      End
      Begin VB.TextBox txt_maker 
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
         Height          =   285
         Left            =   2460
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   34
         Top             =   480
         Width           =   2100
      End
      Begin VB.TextBox txt_qty 
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
         Height          =   285
         Left            =   8130
         TabIndex        =   6
         Top             =   480
         Width           =   840
      End
      Begin VB.TextBox txt_amount 
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
         Height          =   285
         Left            =   12930
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   480
         Width           =   1665
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Qty"
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
         Left            =   7110
         TabIndex        =   71
         Top             =   90
         Width           =   1020
      End
      Begin VB.Line Line4 
         X1              =   90
         X2              =   90
         Y1              =   810
         Y2              =   930
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Cls"
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
         Left            =   9030
         TabIndex        =   59
         Top             =   90
         Width           =   660
      End
      Begin MSForms.ComboBox txt_item_code 
         Height          =   330
         Left            =   0
         TabIndex        =   46
         Top             =   450
         Width           =   1965
         VariousPropertyBits=   746604571
         MaxLength       =   15
         DisplayStyle    =   3
         Size            =   "3466;582"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label10 
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
         Index           =   6
         Left            =   12885
         TabIndex        =   43
         Top             =   90
         Width           =   660
      End
      Begin VB.Label Label10 
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
         Index           =   5
         Left            =   10860
         TabIndex        =   42
         Top             =   90
         Width           =   420
      End
      Begin VB.Label Label10 
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
         Index           =   4
         Left            =   9885
         TabIndex        =   41
         Top             =   90
         Width           =   390
      End
      Begin VB.Label Label10 
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
         Index           =   3
         Left            =   8340
         TabIndex        =   40
         Top             =   90
         Width           =   300
      End
      Begin VB.Label Label10 
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
         Index           =   2
         Left            =   4650
         TabIndex        =   39
         Top             =   90
         Width           =   960
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Part Number"
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
         Left            =   2490
         TabIndex        =   38
         Top             =   90
         Width           =   1080
      End
      Begin MSForms.ComboBox cbo_price 
         Height          =   285
         Left            =   10830
         TabIndex        =   8
         Top             =   480
         Width           =   2040
         VariousPropertyBits=   746604571
         MaxLength       =   16
         DisplayStyle    =   3
         Size            =   "3598;503"
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cbo_curr 
         Height          =   285
         Left            =   9840
         TabIndex        =   7
         Top             =   480
         Width           =   915
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "1614;503"
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label10 
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
         Index           =   0
         Left            =   90
         TabIndex        =   16
         Top             =   90
         Width           =   1155
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00A6D2FF&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   0
         Top             =   30
         Width           =   14610
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid1 
      Height          =   4455
      Left            =   360
      TabIndex        =   31
      Top             =   2880
      Width           =   14610
      _cx             =   25770
      _cy             =   7858
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
      GridColor       =   8421504
      GridColorFixed  =   8421504
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
      Rows            =   1
      Cols            =   6
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
      Height          =   405
      Left            =   13020
      TabIndex        =   60
      Top             =   120
      Width           =   1845
      _extentx        =   3254
      _extenty        =   714
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DO No."
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
      Left            =   11760
      TabIndex        =   82
      Top             =   7515
      Width           =   615
   End
   Begin VB.Label Label20 
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
      Left            =   9720
      TabIndex        =   80
      Top             =   7515
      Width           =   690
   End
   Begin VB.Label Label19 
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
      Left            =   6600
      TabIndex        =   78
      Top             =   7515
      Width           =   1050
   End
   Begin MSForms.ComboBox cboSearch 
      Height          =   315
      Left            =   1080
      TabIndex        =   76
      TabStop         =   0   'False
      Top             =   7485
      Width           =   1845
      VariousPropertyBits=   612386843
      MaxLength       =   6
      DisplayStyle    =   7
      Size            =   "3254;556"
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
      Index           =   8
      Left            =   360
      TabIndex        =   75
      Top             =   7520
      Width           =   600
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
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
      Index           =   8
      Left            =   7410
      TabIndex        =   70
      Top             =   8040
      Width           =   60
   End
   Begin VB.Label Label14 
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
      Height          =   255
      Left            =   8370
      TabIndex        =   57
      Top             =   8910
      Width           =   735
   End
   Begin VB.Label Label13 
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
      Height          =   255
      Left            =   6150
      TabIndex        =   55
      Top             =   8910
      Width           =   645
   End
   Begin VB.Label Label11 
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
      Height          =   255
      Left            =   3840
      TabIndex        =   53
      Top             =   8910
      Width           =   735
   End
   Begin MSForms.ComboBox cbobctype 
      Height          =   315
      Left            =   4650
      TabIndex        =   52
      Top             =   8880
      Width           =   1455
      VariousPropertyBits=   746604571
      MaxLength       =   15
      DisplayStyle    =   3
      Size            =   "2566;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "SJ No."
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
      Left            =   360
      TabIndex        =   45
      Top             =   8910
      Width           =   705
   End
   Begin VB.Line Line6 
      X1              =   2595
      X2              =   2610
      Y1              =   9360
      Y2              =   9375
   End
   Begin VB.Label Label2 
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
      Height          =   255
      Left            =   10770
      TabIndex        =   21
      Top             =   8910
      Width           =   855
   End
   Begin VB.Label lbl_record 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lbl_record"
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
      Left            =   1950
      TabIndex        =   20
      Top             =   10140
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Parts (Material) Supply [Unscheduled]"
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
      Index           =   0
      Left            =   360
      TabIndex        =   19
      Top             =   225
      Width           =   14505
   End
End
Attribute VB_Name = "frm_part_supply"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim db2 As New ADODB.Connection
Dim rs_part_supply As New ADODB.Recordset
Dim rs_warehouse As New ADODB.Recordset
Dim rs_trade_master As New ADODB.Recordset
Dim rs_item As New ADODB.Recordset
Dim l_update_stock As Double
Dim l_tambah_stock As Double
Dim l_item_code_update As String, l_supply_cls As String, l_stock_warehouse As String
Dim stockcontrol_cls   As String, l_stock_location As String, l_seqNo As Double
Dim Status As String, l_SJNo As String, l_Remarks As String

Dim bteColSelect As Byte
Dim bteColProdCode As Byte
Dim bteColPartNo As Byte
Dim bteColDesc As Byte
Dim bteColQty As Byte
Dim bteColCurr As Byte
Dim bteColPrice As Byte
Dim bteColAmount As Byte
Dim bteColRemark As Byte
Dim bteColSeqNo As Byte
Dim bteColItemControl As Byte
Dim bteColSJNo As Byte
Dim bteColBcNo As Byte
Dim bteColDONo As Byte
Dim bteColBctype As Byte
Dim bteColBCDate As Byte
Dim bteColNoRegister As Byte
Dim bteColNoSeri As Byte

Dim bteHakPrice As Byte
Dim ls_ReplacementWarehouseCode As String
Dim ls_FromWarehouseCode As String
Dim ls_ToWarehouseCode As String
Dim ls_SupplySeqNo As Double
Dim ls_SupplyDate As String
Dim ls_SupplyCls As String

Sub Header()

    With Grid1

        bteColSelect = 0
        bteColProdCode = 1
        bteColPartNo = 2
        bteColDesc = 3
        bteColQty = 4
        bteColCurr = 5
        bteColPrice = 6
        bteColAmount = 7
        bteColRemark = 11
        bteColSeqNo = 9
        bteColItemControl = 10
        bteColSJNo = 8
        bteColBCDate = 12
        bteColBcNo = 13
        bteColDONo = 14
        bteColBctype = 15
        bteColNoRegister = 16
        bteColNoSeri = 17

        .clear
        .ColS = 18
        .Rows = 1

        .TextMatrix(0, bteColSelect) = "S"
        .TextMatrix(0, bteColProdCode) = "Product Code"
        .TextMatrix(0, bteColPartNo) = "Part Number"
        .TextMatrix(0, bteColDesc) = "Description"
        .TextMatrix(0, bteColQty) = "Qty"
        .TextMatrix(0, bteColCurr) = "Curr"
        .TextMatrix(0, bteColPrice) = "Price"
        .TextMatrix(0, bteColAmount) = "Amount"
        .TextMatrix(0, bteColRemark) = "Remarks"
        .TextMatrix(0, bteColSJNo) = "SJ No."
        .TextMatrix(0, bteColBCDate) = "BC Date"
        .TextMatrix(0, bteColBctype) = "BC Type"
        .TextMatrix(0, bteColBcNo) = "BC No"
        .TextMatrix(0, bteColDONo) = "DO No"
        .TextMatrix(0, bteColSeqNo) = "seqno"
        .TextMatrix(0, bteColItemControl) = "itemControlcls"
        .TextMatrix(0, bteColNoRegister) = "No. Register"
        .TextMatrix(0, bteColNoSeri) = "No. Seri"

        .ColWidth(bteColSelect) = 250
        .ColWidth(bteColProdCode) = 2000
        .ColWidth(bteColPartNo) = 2000
        .ColWidth(bteColDesc) = 3500
        .ColWidth(bteColQty) = 1200
        .ColWidth(bteColCurr) = 700
        .ColWidth(bteColPrice) = 1500
        .ColWidth(bteColAmount) = 1850
        .ColWidth(bteColRemark) = 1850
        .ColWidth(bteColSJNo) = 1000
        .ColWidth(bteColBCDate) = 1000
        .ColWidth(bteColBcNo) = 1000
        .ColWidth(bteColDONo) = 1000
        .ColWidth(bteColBctype) = 1000
        .ColWidth(bteColNoRegister) = 1700
        .ColWidth(bteColNoSeri) = 900

        .ColHidden(bteColItemControl) = True
        .ColHidden(bteColBCDate) = True
        .ColHidden(bteColBcNo) = True
        .ColHidden(bteColBctype) = True

        .ColHidden(bteColCurr) = (bteHakPrice = 0)
        .ColHidden(bteColPrice) = (bteHakPrice = 0)
        .ColHidden(bteColAmount) = (bteHakPrice = 0)
        
        .ColAlignment(bteColDesc) = flexAlignLeftCenter
        .ColAlignment(bteColQty) = flexAlignRightCenter
        .ColAlignment(bteColCurr) = flexAlignLeftCenter
        .ColAlignment(bteColAmount) = flexAlignRightCenter
         .ColAlignment(bteColSJNo) = flexAlignLeftCenter
         .ColAlignment(bteColRemark) = flexAlignLeftCenter
         .ColAlignment(bteColDONo) = flexAlignLeftCenter
        .ColAlignment(bteColPrice) = flexAlignRightCenter
        .ColAlignment(bteColProdCode) = flexAlignLeftCenter
        .ColAlignment(bteColNoRegister) = flexAlignLeftCenter
        .ColAlignment(bteColNoSeri) = flexAlignRightCenter

        ' ? Override kembali header (row ke-0) supaya rata tengah
        Dim i As Integer
        For i = 0 To .ColS - 1
            .Cell(flexcpAlignment, 0, i) = flexAlignCenterCenter
        Next i

        .EditMaxLength = 1

    End With

End Sub


Private Sub setting_grid()
    
    Dim sql_join As String, l_curr As String, l_item_name As String, L_price As String, l_price2 As String
    Dim rs_join As New ADODB.Recordset, rs_Replacement As New ADODB.Recordset
    
    Dim Bo_Replacement As Boolean
    
    Bo_Replacement = False
    Me.MousePointer = vbHourglass
    
    Header
    
    With Grid1
    DoEvents
'        sql_join = "select * from (select part_supply.*,stockcontrol_cls,makeritem_code,item_name,sheetcoil_cls,width,length,thickness,unit_cls from part_supply join item_master on part_supply.childitem_code=item_master.item_code ) xxx " & vbCrLf & _
'            " where fromwarehouse_code='" & Trim(cbo_warehouse) & "' and " & vbCrLf & _
'            "towarehouse_code='" & Trim(cbo_location) & "' and " & vbCrLf & _
'            "supply_cls='" & Trim(cbo_supply) & "' and " & vbCrLf & _
'            " childsupply_date ='" & Format(DTPicker3.Value, "yyyy-MM-dd") & "' and COALESCE(supplyrec_no,'')=''"
        
' Update Query 20210915
'        sql_join = "select * from (select part_supply.*,stockcontrol_cls,makeritem_code,item_name,sheetcoil_cls,width,length,thickness,unit_cls from (Select * From part_supply where fromwarehouse_code='" & Trim(cbo_warehouse) & "' and towarehouse_code='" & Trim(cbo_location) & "' and supply_cls='" & Trim(cbo_supply) & "' and childsupply_date ='" & Format(DTPicker3.Value, "yyyy-MM-dd") & "' and COALESCE(supplyrec_no,'')='')part_supply join item_master on part_supply.childitem_code=item_master.item_code LEFT JOIN Supply_Scan_Detail SD ON part_supply.SJNo = SD.SJ_No AND SD.Item_Code = part_supply.ChildItem_Code ) xxx order by Register_Date" & vbCrLf & _
'            "  " & vbCrLf & _
'            "" & vbCrLf & _
'            "" & vbCrLf & _
'            " "
            
            sql_join = " SELECT * FROM (SELECT part_supply.*, stockcontrol_cls, makeritem_code, item_name, sheetcoil_cls, width, length, " & vbCrLf & _
                " thickness,unit_cls, Serial_No, Qty, ISNULL(Barcode_No,'')Barcode_No, ROW_NUMBER() OVER (ORDER BY part_supply.Register_Date,part_supply.ChildItem_Code ) NoSeri  " & vbCrLf & _
                " FROM (SELECT * FROM part_supply WHERE fromwarehouse_code='" & Trim(cbo_warehouse) & "' and towarehouse_code='" & Trim(cbo_location) & "' and supply_cls='" & Trim(cbo_supply) & "'  " & vbCrLf & _
                " AND childsupply_date ='" & Format(DTPicker3.Value, "yyyy-MM-dd") & "' and COALESCE(supplyrec_no,'')='')part_supply  " & vbCrLf & _
                " JOIN item_master on part_supply.childitem_code=item_master.item_code  " & vbCrLf & _
                " LEFT JOIN Supply_Scan_Detail SD ON part_supply.SJNo = SD.SJ_No AND SD.Item_Code = part_supply.ChildItem_Code) xxx order by Register_Date "
        
        
        rs_join.Open sql_join, Db, adOpenKeyset, adLockOptimistic
   DoEvents
        
        If rs_join.EOF = False Or rs_join.BOF = False Then
            While rs_join.EOF = False
                
                If cbo_curr.ListCount > 0 Then
                    i = 0
                    For i = 0 To cbo_curr.ListCount - 1
                        If Trim(rs_join!currency_code) = cbo_curr.List(i, 1) Then
                            l_curr = cbo_curr.List(i, 0): Exit For
                        End If
                    l_curr = ""
                    Next
                End If
    
            ' # Cek keberadaan Replacement di Part_Supply
'                sql_join = "Select * From Part_Supply Where SupplySeq_No=" & rs_join!Seq_no
'
'                If rs_Replacement.State = adStateOpen Then rs_Replacement.Close
'
'                rs_Replacement.Open sql_join, Db, adOpenDynamic, adLockOptimistic
'
'                If Not rs_Replacement.EOF Then Bo_Replacement = True
            ' # ---------------
            
                l_item_name = uf_GetItemDescription(Trim(rs_join!childitem_code))
                L_price = Format(Trim(rs_join!Price), gs_formatPrice)
                l_price2 = Trim(L_price)
                
                With Grid1
                    .AddItem ""
                    .TextMatrix(.Rows - 1, bteColProdCode) = Trim(rs_join!childitem_code)
                    .TextMatrix(.Rows - 1, bteColPartNo) = Trim(rs_join!MakerItem_Code)
                    .TextMatrix(.Rows - 1, bteColDesc) = l_item_name
                    If Trim(rs_join!Barcode_No) = "" Then
                         .TextMatrix(.Rows - 1, bteColQty) = Format(Trim(rs_join!ChildRequirement_qty), gs_formatQty)
                    Else
                        .TextMatrix(.Rows - 1, bteColQty) = Format(Trim(rs_join!Qty), gs_formatQty)
                    End If
                    If cbo_supply.Text <> "S" Then
                    .TextMatrix(.Rows - 1, bteColCurr) = uf_GetCurrencyDescription(Trim(rs_join!currency_code))   'Trim(l_curr)
                    End If
                    .TextMatrix(.Rows - 1, bteColPrice) = l_price2
                    .TextMatrix(.Rows - 1, bteColAmount) = Format(Trim(rs_join!Amount), gs_formatAmount)
                    .TextMatrix(.Rows - 1, bteColRemark) = Trim(rs_join!Remarks & rs_join!Serial_No)
                    .TextMatrix(.Rows - 1, bteColSeqNo) = Trim(rs_join!Seq_no)
                    .TextMatrix(.Rows - 1, bteColItemControl) = Trim(rs_join!stockcontrol_cls)
                    .TextMatrix(.Rows - 1, bteColSJNo) = Trim(rs_join!SJNo & "")
                    .TextMatrix(.Rows - 1, bteColBcNo) = Trim(rs_join!BC40_No & "")
                    .TextMatrix(.Rows - 1, bteColDONo) = Trim(rs_join!do_no & "")
                    .TextMatrix(.Rows - 1, bteColBCDate) = Trim(rs_join!BC40_Date & "")
                    .TextMatrix(.Rows - 1, bteColBctype) = Trim(rs_join!BC_Type & "")
                    .TextMatrix(.Rows - 1, bteColNoRegister) = Trim(rs_join!No_Register & "")
                    .TextMatrix(.Rows - 1, bteColNoSeri) = Trim(rs_join!NoSeri & "")
                    
                    rs_join.MoveNext
                End With
            
            Wend
        End If
        rs_join.Close
        Status = "insertnew"

        For i = 1 To .Rows - 1
            .Cell(flexcpBackColor, i, bteColSelect) = vbWhite
        Next
        Call formatPriceGrid
        
        .ColHidden(bteColSeqNo) = True
        .ColHidden(bteColPartNo) = True
        
        If Bo_Replacement = True Then
            cbo_Replacement = cbo_Replacement.Column(0, 1)
        Else
            cbo_Replacement = cbo_Replacement.Column(0, 0)
        End If
        
    End With
    
    Me.MousePointer = vbDefault
    cmd_submit.Enabled = True

End Sub

Sub formatPriceGrid()
    
    Dim desimal As Integer, spasi As String, L_price As String
    Dim j As Integer, jKoma As Integer, k As Integer
    
    desimal = 0
    With Grid1
        For i = 1 To .Rows - 1
            If InStr(1, Trim(.TextMatrix(i, bteColPrice)), ".") > 0 Then
                If desimal < (Len(Trim(.TextMatrix(i, bteColPrice))) - InStr(1, Trim(.TextMatrix(i, bteColPrice)), ".")) Then
                    desimal = Len(Trim(.TextMatrix(i, bteColPrice))) - InStr(1, Trim(.TextMatrix(i, bteColPrice)), ".")
                End If
            End If
        Next
        
        desimal = desimal - 1
        For i = 1 To .Rows - 1
            L_price = Trim(.TextMatrix(i, bteColPrice))
            If InStr(1, L_price, ".") = 0 Then
                spasi = ""
                For j = 1 To desimal + 2
                    spasi = spasi + " "
                Next
                L_price = L_price & spasi
            ElseIf Len(L_price) - InStr(1, L_price, ".") = 5 Then
                L_price = L_price
            Else
                jKoma = Len(L_price) - InStr(1, L_price, ".")
                For k = 0 To desimal - jKoma
                    L_price = L_price + " "
                Next
            End If
            .TextMatrix(i, bteColPrice) = L_price
        Next
    End With
    
End Sub

Private Sub cbo_curr_Click()
If cbo_curr.DataChanged = True Then
'browseprice
End If
End Sub

Private Sub cbo_curr_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        For i = 0 To cbo_curr.ListCount - 1
            If UCase(Trim(cbo_curr)) = UCase(Trim(cbo_curr.List(i, 0))) Then
                cbo_curr.Text = cbo_curr.List(i, 0)
                lbl_pesan.Caption = ""
                Exit For
            Else
                lbl_pesan.Caption = DisplayMsg(4005) '"Invalid currency clasification !"
            End If
        Next
    End If
End Sub

Private Sub cbo_location_Change()
    If cbo_location.Text = "" Then
        lbl_location.Caption = ""
    End If
    Call delivery
End Sub

Private Sub cbo_location_Click()
    If cbo_location.DataChanged = False Then Exit Sub
'    Call clear_framebawah
    If cbo_location.ListIndex <> -1 Then
        lbl_location.Caption = cbo_location.List(cbo_location.ListIndex, 1)
        l_stock_location = Trim(cbo_location.List(cbo_location.ListIndex, 2))
        Call browseprice
'        cbo_curr.clear
    Else
        lbl_location.Caption = ""
    End If
    lbl_pesan = validCombo
    'Call uf_ValidateComboData(cbo_location, "4023", lbl_pesan, lbl_location)
    If Trim(lbl_pesan) <> "" Then clearGrid: Exit Sub
    'Call Header
    'Call setting_grid
    lbl_pesan = ""
    Call cbo_location_Change
End Sub

Private Sub cbo_location_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        lbl_pesan = ""
        Call clear_framebawah
        cbo_location.DataChanged = False
        lbl_pesan = validCombo
        cbo_location.DataChanged = True
        If Trim(lbl_pesan) <> "" Then clearGrid: Exit Sub
        Call browseprice
        'Call Header
        Call setting_grid
    End If
End Sub

Private Sub cbo_location_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub cbo_price_KeyPress(KeyAscii As MSForms.ReturnInteger)
     If InStr(1, "0123456789.", Chr(KeyAscii)) <= 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub cbo_price_LostFocus()
    cbo_price = Format(cbo_price.Text, gs_formatPrice)
End Sub

Private Sub cbo_Replacement_Change()
    lbl_pesan = ""
    If cbo_Replacement.Text = "" Then
        cbo_ReplacementWarehouseCode.clear
    End If
    Call cbo_Replacement_Click
End Sub

Private Sub cbo_Replacement_Click()
    lbl_pesan = ""
    If UCase(Trim(cbo_Replacement.Text)) = "YES" Then
        cbo_ReplacementWarehouseCode.clear
        cbo_ReplacementWarehouseCode.columnCount = 3
        cbo_ReplacementWarehouseCode.TextColumn = 1
        
        i = 0
        If rs_warehouse.EOF = False Or rs_warehouse.BOF = False Then
            rs_warehouse.MoveFirst
            While rs_warehouse.EOF = False
                cbo_ReplacementWarehouseCode.AddItem ""
                cbo_ReplacementWarehouseCode.List(i, 0) = Trim(rs_warehouse!wh_code)
                cbo_ReplacementWarehouseCode.List(i, 1) = Trim(rs_warehouse!WH_Name)
                cbo_ReplacementWarehouseCode.List(i, 2) = Trim(rs_warehouse!stockcontrol_cls)
                rs_warehouse.MoveNext
                i = i + 1
            Wend
            cbo_ReplacementWarehouseCode.ColumnWidths = "50 pt; 175 pt; 0 pt"
            cbo_ReplacementWarehouseCode.ListWidth = 225
            cbo_ReplacementWarehouseCode = "WH-001"
        End If
    Else
        cbo_ReplacementWarehouseCode.clear
    End If
    
End Sub

Private Sub cbo_ReplacementWarehouseCode_Change()
    If cbo_ReplacementWarehouseCode.Text = "" Then
        lbl_ReplacementWarehouseCode.Caption = ""
    End If
End Sub

Private Sub cbo_ReplacementWarehouseCode_Click()
    lbl_pesan.Caption = ""
    lbl_ReplacementWarehouseCode.Caption = cbo_ReplacementWarehouseCode.List(cbo_ReplacementWarehouseCode.ListIndex, 1)
End Sub

Private Sub cbo_supply_Change()
    lbl_pesan = ""
    If cbo_supply.Text = "" Then
        lbl_supply.Caption = ""
    End If
    lbl_pesan = validCombo
    If Trim(lbl_pesan) <> "" Then clearGrid: Exit Sub
    'Call Header
    'Call setting_grid
    Call clear_framebawah
End Sub

Private Sub cbo_supply_Click()
    lbl_supply.Caption = cbo_supply.List(cbo_supply.ListIndex, 1)
    'Call Header
    Call setting_grid
End Sub

Private Sub cbo_supply_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        lbl_pesan = ""
        lbl_pesan = validCombo
        If Trim(lbl_pesan) <> "" Then clearGrid: Exit Sub
        'Call Header
        Call setting_grid
    End If
End Sub

Private Sub cbo_supply_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub cbo_warehouse_Change()
    If cbo_warehouse.Text = "" Then
        lbl_warehouse.Caption = ""
    End If
End Sub

Private Sub cbo_warehouse_Click()
    If cbo_warehouse.DataChanged = False Then Exit Sub
    lbl_warehouse.Caption = cbo_warehouse.List(cbo_warehouse.ListIndex, 1)
    l_stock_warehouse = Trim(cbo_warehouse.List(cbo_warehouse.ListIndex, 2))
'    Call set_item
    lbl_pesan = validCombo
    If Trim(lbl_pesan) <> "" Then clearGrid: Exit Sub
'    Call clear_framebawah
    lbl_pesan = ""
    'Call Header
    Call setting_grid
End Sub

Private Sub set_item(Optional ls_SortBy As String)
    
    Dim sqlitem As String
    Dim RsItem As New Recordset
    
    If Check1.Value = 1 Then
        ls_SortBy = "item_name"
    Else
        ls_SortBy = ""
    End If
    
    sqlitem = "select item_code, makeritem_code, item_name, address ,a.unit_cls,b.Description from item_master a inner join unit_cls b on a.Unit_Cls=b.Unit_Cls where use_endday > convert(char(8), getdate(), 112)  "
    If Trim(ls_SortBy) = "" Then
        sqlitem = sqlitem & " order by item_code asc "
    Else
        sqlitem = sqlitem & " order by " & ls_SortBy & " asc "
    End If
    Set RsItem = Db.Execute(sqlitem)
    
    With txt_item_code
        .clear
        .columnCount = 5
        .ColumnWidths = "110pt;110pt;240pt;0pt;0pt"
        .ListWidth = 440
        .ListRows = 15
    
        i = 0
        Do While Not RsItem.EOF
            .AddItem
            .List(i, 0) = Trim(RsItem("item_code"))
            .List(i, 1) = Trim(RsItem("makeritem_code"))
            .List(i, 2) = Trim(RsItem("item_Name"))
            .List(i, 3) = IIf(IsNull(Trim(RsItem("address"))), "", Trim(RsItem("address")))
            .List(i, 4) = IIf(IsNull(Trim(RsItem("Description"))), "", Trim(RsItem("Description")))
            
            RsItem.MoveNext
            i = i + 1
        Loop
    End With

End Sub

Private Sub cbo_warehouse_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    
    If KeyCode = 13 Then
        lbl_pesan = ""
        'Call set_item
        cbo_warehouse.DataChanged = False
        lbl_pesan = validCombo
        cbo_warehouse.DataChanged = True
        If Trim(lbl_pesan) <> "" Then clearGrid: Exit Sub
        Call setting_grid
        Call clear_framebawah
        lbl_warehouse = cbo_warehouse.List(cbo_warehouse.ListIndex, 1)
    End If
End Sub

Sub clearGrid()
    Call Header
End Sub

Private Sub cbo_warehouse_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub cbobctype_Change()
    txtSJNo_LostFocus
    
    up_GetNoSeri (txt_item_code.Text)
    
    If cboBCType.Text = "2.6.1" Then
        txtDONo.Enabled = True
    Else
        txtDONo.Enabled = False
        txtDONo.Text = ""
    End If
End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
    set_item ("item_name")
Else
    set_item
End If
End Sub

Private Sub cmd_Cancel_Click()
    'Call set_item
    Call clear_framebawah(True)
    lbl_pesan = ""
    'Call Header
    Call setting_grid
End Sub

Private Sub cmd_clear_Click()
cmd_submit.Enabled = True
    DTPicker3.Value = Format(Date, "dd MMM yyyy")
    cbo_location = ""
    cbo_supply = "S1"
    cbo_Replacement = cbo_Replacement.Column(0, 0)
    cbo_warehouse = ""
    Call clear_framebawah(True)
    Call Header
    'Call setting_grid
    lbl_pesan.Caption = ""
    lbl_warehouse.Caption = ""
    lbl_location.Caption = ""
    lbl_supply.Caption = ""
End Sub

Private Sub cmd_print_Click()
Call toExcel
End Sub
Sub toExcel()
Dim xlapp As New Excel.application
Dim rsCek As New Recordset, Row As Long, tempi As String, tempcust As String
Dim bolcust As Boolean, bolinv As Boolean
Dim rsCompany As New Recordset, sql_plus As String, sqlP As String
Dim sqlControl As String, RsInvControl As New ADODB.Recordset
Dim selisih As Double
Dim nomor As Integer
Dim sql_sum As String

Dim ls_no As String
Dim ls_nama_part As String
Dim ls_kode_part As String
Dim ls_qty_pengiriman As String
Dim ls_satuan As String
Dim ls_keterangan As String
Dim ls_c As String
Dim ls_d As String
    
ls_no = "A"
ls_nama_part = "B"
ls_kode_part = "E"
ls_qty_pengiriman = "F"
ls_satuan = "G"
ls_keterangan = "H"
ls_c = "C"
ls_d = "D"
    
Me.MousePointer = vbHourglass

'CboLocationCD = Trim(CboLocationCD)
'    sql = "select * from (select part_supply.*,stockcontrol_cls,makeritem_code,item_name,sheetcoil_cls,width,length,thickness,unit_cls from (Select * From part_supply where fromwarehouse_code='" & Trim(cbo_warehouse) & "' and towarehouse_code='" & Trim(cbo_location) & "' and supply_cls='" & Trim(cbo_supply) & "' and childsupply_date ='" & Format(DTPicker3.Value, "yyyy-MM-dd") & "' and COALESCE(supplyrec_no,'')='')part_supply join item_master on part_supply.childitem_code=item_master.item_code ) xxx " & vbCrLf & _
'            "  " & vbCrLf & _
'            "" & vbCrLf & _
'            "" & vbCrLf & _
'            " "

sql = "   select * from  " & vbCrLf & _
            "   (select part_supply.*,stockcontrol_cls,makeritem_code,item_name,sheetcoil_cls, " & vbCrLf & _
            "   width,length,thickness,item_master.Unit_Cls,uc.description " & vbCrLf & _
            "   from (Select * From part_supply  " & vbCrLf & _
            "   where fromwarehouse_code='" & Trim(cbo_warehouse) & "' and towarehouse_code='" & Trim(cbo_location) & "' and supply_cls='" & Trim(cbo_supply) & "'  " & vbCrLf & _
            "   and childsupply_date ='" & Format(DTPicker3.Value, "yyyy-MM-dd") & "' and COALESCE(supplyrec_no,'')='')part_supply  " & vbCrLf & _
            "   join item_master on part_supply.childitem_code=item_master.item_code  " & vbCrLf & _
            "   join Unit_Cls uc on uc.Unit_Cls=part_supply.childunit_cls ) xxx  " & vbCrLf & _
            "    " & vbCrLf & _
            "  "

If rsCek.State <> adStateClosed Then rsCek.Close
rsCek.CursorLocation = adUseClient
rsCek.Open sql, Db, adOpenDynamic, adLockOptimistic

If Not rsCek.EOF Then
Screen.MousePointer = vbHourglass
With xlapp

    sql = "select rtrim(company_name) company_name, rtrim(address1) Address1, rtrim(Address2) Address2, rtrim(Province) Province, rtrim(city) City, Rtrim(Postal_Code) POstal_Code, Rtrim(phone1) Phone1, Rtrim(phone2) Phone2,rtrim(fax) Fax  From company_profile "
    If rsCompany.State <> adStateClosed Then rsCompany.Close
    rsCompany.Open sql, Db, adOpenDynamic, adLockOptimistic
    If rsCompany.EOF Then Screen.MousePointer = vbDefault: Exit Sub

    .Workbooks.Add
    .Range("a4") = rsCompany!company_name '"Judul Company"
    .Range("A5") = rsCompany!address1 & " " & rsCompany!address2 & " " & rsCompany!City & " " & rsCompany!Province & " " & rsCompany!postal_code
    .Range("A6") = "Phone (0264)351323-6. Fax:(0264)351327"
    .Range("A8") = "No"
    .Range("A9") = "Cust PO No"
    .Range("A10") = "BC Type"
    .Range("A11") = "BC Number"
    .Range("A12") = "QTY"
    .Range("F8") = "Delivery To"
    .Range("G8") = ": " & CboDelivery.Text
    .Range("F11") = "Model"
    .Range("A14") = "Kami Kirimkan barang-barang tersebut dibawah ini dengan kendaraan: .......................................... No: ........................"
    .Range("H1") = "Cikampek," & " " & Format(Now, "dd mmm YYYY")
    .Range("H4") = "Surat Jalan"
    .Range("C8") = ": " & TXTNo.Text
    .Range("C9") = ": " & txtPoNo.Text
    .Range("C10") = ": " & Trim(rsCek!BC_Type)
    .Range("C11") = ": " & Trim(rsCek!BC40_No)
    
    .Range("B16:D16").Merge
    .Range("A4").Font.Size = 14
    .Range("H4").Font.Size = 18
    .Range("H4").Font.Bold = True
    '.ActiveSheet.Pictures.Insert(App.path & "\Reports\kawai1.jpg").Select
        
    
    .ActiveSheet.Cells(1, 1).columnWidth = 5
    .ActiveSheet.Cells(1, 2).columnWidth = 4.45
    .ActiveSheet.Cells(1, 3).columnWidth = 20
    .ActiveSheet.Cells(1, 4).columnWidth = 18
    .ActiveSheet.Cells(1, 5).columnWidth = 12
    .ActiveSheet.Cells(1, 6).columnWidth = 15
    .ActiveSheet.Cells(1, 7).columnWidth = 6.71
    .ActiveSheet.Cells(1, 8).columnWidth = 30
    '.Columns("H:H")ColumnWidth = 30
    .Range(ls_no & 6, ls_keterangan & 6).Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    Row = 16

Dim jumlah As Double
    Do While Not rsCek.EOF
        If Row = 16 Then
            .Range(ls_no & Row) = "No"
            .Range(ls_nama_part & Row) = "Nama Part"
            .Range(ls_kode_part & Row) = "Kode Part"
            .Range(ls_qty_pengiriman & Row) = "QTY Pengiriman"
            .Range(ls_satuan & Row) = "Satuan"
            .Range(ls_keterangan & Row) = "Keterangan"
            
            

            Row = Row + 1
        End If
        nomor = nomor + 1
        Row = Row
        jumlah = jumlah + rsCek!ChildRequirement_qty
        .Range(ls_no & Row) = nomor
        .Range(ls_nama_part & Row) = Trim(rsCek!item_name)
        .Range(ls_kode_part & Row) = Trim(rsCek!MakerItem_Code)
        .Range(ls_qty_pengiriman & Row) = Format(rsCek!ChildRequirement_qty)
        .Range(ls_satuan & Row) = (rsCek!Description)
        .Range(ls_keterangan & Row) = Format(rsCek!Remarks)
        
        Row = Row + 1
        .Range(ls_no & Row - 1).horizontalAlignment = xlCenter
        .Range(ls_kode_part & Row - 1).horizontalAlignment = xlCenter
        .Range(ls_qty_pengiriman & Row - 1).horizontalAlignment = xlCenter
        .Range(ls_satuan & Row - 1).horizontalAlignment = xlLeft
        
        rsCek.MoveNext
    Loop
    .Range("C12") = ": " & jumlah
'    'Border
    .Range(ls_no & 16, ls_keterangan & Row - 1).Borders(xlEdgeTop).LineStyle = xlContinuous
    .Range(ls_no & 16, ls_keterangan & Row - 1).Borders(xlEdgeBottom).LineStyle = xlContinuous
    .Range(ls_no & 16, ls_keterangan & Row - 1).Borders(xlEdgeLeft).LineStyle = xlContinuous
    .Range(ls_no & 16, ls_keterangan & Row - 1).Borders(xlEdgeRight).LineStyle = xlContinuous
    .Range(ls_no & 16, ls_keterangan & Row - 1).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    .Range(ls_no & 16, ls_keterangan & Row - 1).Borders(xlInsideVertical).LineStyle = xlContinuous
    
    .Range(ls_nama_part & 16, ls_c & 16 & Row - 1).Borders(xlEdgeRight).LineStyle = xlNone
    
    .Range(ls_no & Row + 2) = "* Please return this original latter to PT.Kawai Indonesia Plat-3"
    .Range(ls_no & Row + 2).Font.Italic = True
    
    .Range(ls_no & Row + 5) = "Delivered by,"
    .Range(ls_no & Row + 10, ls_nama_part & Row + 10).Borders(xlEdgeTop).LineStyle = xlContinuous
    .Range(ls_no & Row + 10) = "WH.Member"
    
    .Range(ls_d & Row + 5) = "Approved by,"
    .Range(ls_d & Row + 10).Borders(xlEdgeTop).LineStyle = xlContinuous
    .Range(ls_d & Row + 10) = "WH.Function Head"
    
    .Range(ls_qty_pengiriman & Row + 5) = "Checked by,"
    .Range(ls_qty_pengiriman & Row + 10).Borders(xlEdgeTop).LineStyle = xlContinuous
    .Range(ls_qty_pengiriman & Row + 10) = "Security"
    
    .Range(ls_keterangan & Row + 5) = "Received by,"
    .Range(ls_keterangan & Row + 10).Borders(xlEdgeTop).LineStyle = xlContinuous
    .Range(ls_keterangan & Row + 10) = "Customer"
    
    .Range("A1", "C1").Font.Bold = True
    .Range("A16:H16").Font.Bold = True
    .Range("a4").Font.Bold = True
    .Range("A1:H1").Columns.Font.Name = "Arial"
    .Range("A1:H1").Columns.Font.Size = "10"
    .Range("H1", "H4").horizontalAlignment = xlRight
    .ActiveSheet.PageSetup.Orientation = xlLandscape
    .Range("A1:H1").Columns.AutoFit
    .Range("A1").Select
    .WindowState = xlMaximized
    .Visible = True
End With

Else
    lbl_pesan = DisplayMsg(4006)
End If

Screen.MousePointer = vbDefault
Me.MousePointer = vbDefault
End Sub

Private Sub cmd_search_Click()
cmd_search.Enabled = False
    lbl_pesan = ""
    lbl_pesan = validCombo
    If Trim(lbl_pesan) <> "" Then clearGrid: Exit Sub
    Call browseprice
    Call clear_framebawah
    Call setting_grid
cmd_search.Enabled = True
End Sub

Private Sub cmd_sub_menu_Click()
    frmMainMenu.Show
    Unload Me
End Sub

Private Sub Cmd_Submit_Click()
    Dim s As Integer, d As Integer, j  As Integer
    Dim l_curr As String, sql_del As String, l_amount As String, l_qty As String, L_price As String, l_unit_cls As String, sql_prod As String
    Dim rs As New ADODB.Recordset, ls_sql As String
   ' On Error GoTo ErrHandler
    
    cmd_submit.Enabled = False
    If hakUpdate(Me.Name) = 0 Then lbl_pesan = DisplayMsg(3008): cmd_submit.Enabled = True: Exit Sub
    
    lbl_pesan = up_ValidateDateRange(DTPicker3.Value, True)
    If lbl_pesan.Caption <> "" Then cmd_submit.Enabled = True: cmd_submit.Enabled = True: Exit Sub
    
    s = 0
    d = 0
    
    '#Get Last Closing Info
    Dim ls_ClosingMonth As String
    Dim ls_ClosingYear As String
    ls_ClosingMonth = uf_GetLastClosing("month")
    ls_ClosingYear = uf_GetLastClosing("year")
    
    '#Validate date Range
    lbl_pesan = up_ValidateDateRange(DTPicker3.Value, True)
    If lbl_pesan <> "" Then cmd_submit.Enabled = True: cmd_submit.Enabled = True:   Exit Sub
    
    '#Validate Replacement Warehouse Code
    If UCase(Trim(cbo_Replacement.Text)) = "YES" Then
        If cbo_ReplacementWarehouseCode.Text = "" Then
            cbo_ReplacementWarehouseCode.SetFocus
            lbl_pesan.Caption = DisplayMsg("9002")  'Please Select Replacement Warehouse Code !
            cmd_submit.Enabled = True
            Exit Sub
        End If
        
        If MsgBox("Process this data with Replacement from " & Trim(lbl_ReplacementWarehouseCode) & " ? ", _
                vbYesNo + vbQuestion + vbDefaultButton2, "Confirmation") = vbNo Then cmd_submit.Enabled = True: Exit Sub
    End If
    
    
        
    'Check Status (Update,Insert, or Delete)
    For i = 1 To Grid1.Rows - 1
        If Grid1.TextMatrix(i, bteColSelect) = "S" Then s = 1
        If Grid1.TextMatrix(i, bteColSelect) = "D" Then d = 1
    Next
    
    l_stock_location = cbo_location.Column(2)
    
    If d = 1 Then
        If (MsgBox("Are you sure want to delete?", vbQuestion + vbDefaultButton2 + vbYesNo, "Confirmation") = vbYes) Then
            Status = "delete"
            GoTo delete
        Else
            cmd_submit.Enabled = True
        End If
    Else
        If s = 1 Then
            Status = "update": GoTo update
        Else
            Status = "insertdetail":    GoTo inserto
        End If
    End If
    
    Exit Sub

delete:
    'Dim rs As New ADODB.Recordset
    db2.BeginTrans
    Me.MousePointer = vbHourglass
    
    
    For i = 1 To Grid1.Rows - 1
        If Grid1.TextMatrix(i, bteColSelect) = "D" Then
            l_tambah_stock = Grid1.TextMatrix(i, bteColQty)
            
            '#Delete data in part Supply berdasarkan Seq_No
            ls_sql = " DELETE FROM   Part_Supply WITH (updlock) " & vbCrLf & _
                    " WHERE Seq_No = '" & Val(Trim(Grid1.TextMatrix(i, bteColSeqNo))) & "' "
            db2.Execute ls_sql
                      
            
            '#Init Control Cls
            'FromControlCls = l_stock_warehouse
            FromControlCls = cbo_warehouse.Column(2)
            ItemControlCls = Grid1.TextMatrix(i, bteColItemControl) 'stockcontrol_cls
    
            '#Check if item influence the stock or not
            If ItemControlCls = "01" Then
                '# Delete data from stock Master
                Call up_UpdateStockMaster(Format(DTPicker3.Value, "yyyy-MM-dd"), ls_ClosingMonth, ls_ClosingYear, Trim(cbo_warehouse), Trim(cbo_location), Trim(Grid1.TextMatrix(i, bteColProdCode)), 0 - CDbl(l_tambah_stock), Trim(cbo_supply), Trim(l_stock_location), "", "D", "", "", False, False, True, db2)
    
                '# Erase data from stockMaster ( base on From WareHouse code)
                Call up_EraseBlankDataInStockMaster(Trim(cbo_warehouse), Trim(Grid1.TextMatrix(i, bteColProdCode)), Trim(Grid1.TextMatrix(i, bteColDesc)))
    
                '# Erase data from stockMaster ( base on To WareHouse code)
                Call up_EraseBlankDataInStockMaster(Trim(cbo_location), Trim(Grid1.TextMatrix(i, bteColProdCode)), Trim(Grid1.TextMatrix(i, bteColDesc)))
            End If
        
            '=== Update Replacement ===
            If UCase(Trim(cbo_Replacement.Text)) = "YES" Then

                '#Check jika SupplySeq_No nya sesuai dengan Seq_No berdasarkan replacement
                ls_sql = " SELECT * FROM Part_Supply " & vbCrLf & _
                        " WHERE SupplySeq_No = " & Val(Trim(Grid1.TextMatrix(i, bteColSeqNo))) & " "
                If rs.State = adStateOpen Then rs.Close
                Set rs = db2.Execute(ls_sql)
                
                If Not rs.EOF Then
                    ls_FromWarehouseCode = Trim(rs!FromWarehouse_Code)
                    ls_ToWarehouseCode = Trim(rs!towarehouse_code)
                End If
                
                rs.Close
                
                '#Delete data in part Supply berdasarkan SupplySeq_No
                ls_sql = " DELETE FROM   Part_Supply WITH (updlock) " & vbCrLf & _
                        " WHERE SupplySeq_No = '" & Val(Trim(Grid1.TextMatrix(i, bteColSeqNo))) & "' "
                db2.Execute ls_sql
                          
                
                '#Init Control Cls
                'FromControlCls = l_stock_warehouse
    '            FromControlCls = cbo_warehouse.Column(2)
                ItemControlCls = Grid1.TextMatrix(i, bteColItemControl) 'stockcontrol_cls
        
                '#Check if item influence the stock or not
                If ItemControlCls = "01" Then
                    '# Delete data from stock Master
                    Call up_UpdateStockMaster(Format(DTPicker3.Value, "yyyy-MM-dd"), ls_ClosingMonth, ls_ClosingYear, ls_FromWarehouseCode, ls_ToWarehouseCode, Trim(Grid1.TextMatrix(i, bteColProdCode)), 0 - CDbl(l_tambah_stock), Trim(cbo_supply), Trim(l_stock_location), "", "D", "", "", False, False, True, db2)
        
                    '# Erase data from stockMaster ( base on From WareHouse code)
                    Call up_EraseBlankDataInStockMaster(ls_FromWarehouseCode, Trim(Grid1.TextMatrix(i, bteColProdCode)), Trim(Grid1.TextMatrix(i, bteColDesc)))
        
                    '# Erase data from stockMaster ( base on To WareHouse code)
                    Call up_EraseBlankDataInStockMaster(ls_ToWarehouseCode, Trim(Grid1.TextMatrix(i, bteColProdCode)), Trim(Grid1.TextMatrix(i, bteColDesc)))
                End If
            End If
        End If
    Next
    
    db2.CommitTrans
    Call setting_grid
    
    lbl_pesan.Caption = DisplayMsg(1201) '"Delete data success !"
    Me.MousePointer = vbDefault
    cmd_submit.Enabled = True
    Exit Sub
    
update:
    Dim rs_check As New ADODB.Recordset
    
    If validasi = False Then cmd_submit.Enabled = True: Exit Sub
    Me.MousePointer = vbHourglass
    db2.BeginTrans
    
    If cbo_curr.ListCount > 0 Then
        i = 0
        For i = 0 To cbo_curr.ListCount - 1
            If Trim(cbo_curr) = cbo_curr.List(i, 0) Then
                l_curr = cbo_curr.List(i, 1): Exit For
            End If
        Next
    End If
    
    txt_item_code.Enabled = False
    
    '#Update data in part Supply
    sql_del = "update part_supply with (updlock) " & _
        "set childrequirement_qty='" & CDbl(Trim(txt_qty)) & "', " & _
        "currency_code='" & Trim(l_curr) & "', " & _
        "remarks='" & Trim(txt_remarks) & "', " & _
        "SJNO='" & Trim(txtsjno) & "', " & _
        "BC40_No='" & Trim(txtBCNo) & "', " & _
        "BC_type='" & Trim(cboBCType.Text) & "', " & _
        "BC40_date='" & Format(DtBCDate.Value, "yyyy-mm-dd") & "'," & _
        "from_address='" & Trim(txt_fromadd) & "', " & _
        "price='" & CDbl(Trim(cbo_price)) & "', " & _
        "amount='" & CDbl(Trim(txt_amount)) & "', " & _
        "Last_Update = getdate(), Last_User = '" & userLogin & "', " & _
        "No_Register='" & Trim(txtRegisterNo) & "', " & _
        "No_Seri='" & Trim(txtNoSeri) & "', " & _
        "DO_No='" & Trim(txtDONo) & "' " & _
        "where seq_no=" & l_seqNo & " "
    db2.Execute (sql_del)
    
    '#Init Update Qty
    l_update_stock = l_update_stock - CDbl(Trim(txt_qty))
    
    '#Init Control Cls
    'FromControlCls = l_stock_warehouse
    FromControlCls = cbo_warehouse.Column(2)
    'ItemControlCls = stockcontrol_cls <=== dihandle di grid after edit
    
    '#Check if item influence the stock or not
    If ItemControlCls = "01" Then
        '# Update data in stock Master
        Call up_UpdateStockMaster(Format(DTPicker3.Value, "yyyy-MM-dd"), ls_ClosingMonth, ls_ClosingYear, Trim(cbo_warehouse), Trim(cbo_location), Trim(l_item_code_update), 0 - CDbl(l_update_stock), Trim(cbo_supply), Trim(l_stock_location), "", "U", "", "", False, False, True, db2)
    End If
    
    '=== Update Replacement ===
    
    If UCase(Trim(cbo_Replacement.Text)) = "YES" Then
    
        '#Check jika SupplySeq_No nya sesuai dengan Seq_No berdasarkan replacement
        ls_sql = " SELECT * FROM Part_Supply " & vbCrLf & _
                " WHERE SupplySeq_No = " & l_seqNo & " "
        If rs.State = adStateOpen Then rs.Close
        Set rs = db2.Execute(ls_sql)
        
        If Not rs.EOF Then
            ls_FromWarehouseCode = Trim(rs!FromWarehouse_Code)
            ls_ToWarehouseCode = Trim(rs!towarehouse_code)
        End If
        
        '#Update data in part Supply
        sql_del = "update part_supply with (updlock) " & _
            "set childrequirement_qty='" & CDbl(Trim(txt_qty)) & "', " & _
            "currency_code='" & Trim(l_curr) & "', " & _
            "remarks='" & Trim(txt_remarks) & "', " & _
            "SJNO='" & Trim(txtsjno) & "', " & _
            "BC40_No='" & Trim(txtBCNo) & "', " & _
            "BC_type='" & Trim(cboBCType.Text) & "', " & _
            "BC40_date='" & Format(DtBCDate.Value, "yyyy-mm-dd") & "', " & _
            "from_address='" & Trim(txt_fromadd) & "', " & _
            "price='" & CDbl(Trim(cbo_price)) & "', " & _
            "amount='" & CDbl(Trim(txt_amount)) & "', " & _
            "Last_Update = getdate(), Last_User = '" & userLogin & "' " & _
            "No_Register='" & CDbl(Trim(txtRegisterNo)) & "', " & _
            "No_Seri='" & Trim(txtNoSeri) & "', " & _
            "DO_No='" & Trim(txtDONo) & "' " & _
            "where SupplySeq_No=" & l_seqNo & " "
        db2.Execute (sql_del)
        
        '#Init Update Qty
    '    l_update_stock = l_update_stock - CDbl(Trim(txt_qty))
        
        '#Init Control Cls
        'FromControlCls = l_stock_warehouse
        FromControlCls = cbo_warehouse.Column(2)
        'ItemControlCls = stockcontrol_cls <=== dihandle di grid after edit
        
        '#Check if item influence the stock or not
        If ItemControlCls = "01" Then
            '# Update data in stock Master
            Call up_UpdateStockMaster(Format(DTPicker3.Value, "yyyy-MM-dd"), ls_ClosingMonth, ls_ClosingYear, ls_FromWarehouseCode, ls_ToWarehouseCode, Trim(l_item_code_update), 0 - CDbl(l_update_stock), Trim(cbo_supply), Trim(l_stock_location), "", "U", "", "", False, False, True, db2)
        End If
    End If
    
    db2.CommitTrans
    txt_item_code.Enabled = True
    Call clear_framebawah
    Call setting_grid
    lbl_pesan.Caption = DisplayMsg(1101) ' "Update data success !"
    Me.MousePointer = vbDefault
    Exit Sub
    
inserto:
    If validasi = False Then cmd_submit.Enabled = True: Exit Sub
    
    l_curr = Trim(cbo_curr)
    l_amount = Trim(txt_amount)
    l_qty = Trim(txt_qty)
    L_price = Trim(cbo_price)
    
    Select Case Trim(cbo_supply)
    Case "S1":
        If Trim(cbo_warehouse) = Trim(cbo_location) Then
            lbl_pesan.Caption = DisplayMsg(4053) '"Can't supply to same warehouse !"
            cmd_submit.Enabled = True
            Exit Sub
        End If
    Case "S":
    Case "L":
        If Trim(cbo_warehouse) <> Trim(cbo_location) Then
            lbl_pesan.Caption = DisplayMsg(4054) '"Can't input loss to different warehouse !"
            Exit Sub
        End If
    Case "RJ":
        If Trim(cbo_warehouse) <> Trim(cbo_location) Then
            lbl_pesan.Caption = DisplayMsg(4055) '"Can't input reject to different warehouse !"
            Exit Sub
        End If
    End Select
    
    'Dim SSeqNo As Long
    Me.MousePointer = vbHourglass
    db2.BeginTrans
    
    rs_part_supply.AddNew
    'rs_part_supply!Seq_No = SSeqNo
    rs_part_supply!FromWarehouse_Code = Trim(cbo_warehouse)
    rs_part_supply!from_address = Trim(txt_fromadd)
    rs_part_supply!towarehouse_code = Trim(cbo_location)
    rs_part_supply!childsupply_date = Format(Trim(DTPicker3.Value), "yyyy-MM-dd")
    rs_part_supply!childitem_code = Trim(txt_item_code)
    rs_part_supply!supply_cls = Trim(cbo_supply)
    rs_part_supply!ChildRequirement_qty = Trim(txt_qty)
    
    Dim rs_item_master As New ADODB.Recordset
    rs_item_master.Open "select * from item_master where item_code='" & Trim(txt_item_code) & "'", Db, adOpenKeyset, adLockOptimistic
    If rs_item_master.EOF = False Then
        l_unit_cls = Trim(rs_item_master!Unit_cls)
    Else
        l_unit_cls = ""
    End If
    rs_item_master.Close
        
    i = 0
    For i = 0 To cbo_curr.ListCount - 1
        If Trim(cbo_curr) = cbo_curr.List(i, 0) Then
            l_curr = cbo_curr.List(i, 1): Exit For
        End If
    Next
    
    '#Insert data into part Supply
    rs_part_supply!childunit_cls = Trim(l_unit_cls)
    rs_part_supply!currency_code = Trim(l_curr)
    rs_part_supply!Price = Trim(cbo_price)
    rs_part_supply!Amount = Trim(txt_amount)
    rs_part_supply!Remarks = Trim(txt_remarks)
    rs_part_supply!SJNo = Trim(txtsjno)
    rs_part_supply!BC40_No = Trim(txtBCNo)
    rs_part_supply!BC_Type = Trim(cboBCType.Text)
    rs_part_supply!BC40_Date = Format(DtBCDate.Value, "yyyy-mm-dd")
    rs_part_supply!parentItem_code = ""
    rs_part_supply!Lot_no = ""
    rs_part_supply!production_date = Null
    rs_part_supply!do_no = Trim(txtDONo)
    rs_part_supply!Remarks = Trim(txt_remarks)
    rs_part_supply!No_Register = Trim(txtRegisterNo)
    rs_part_supply!No_Seri = Trim(txtNoSeri)
    rs_part_supply!SupplySeq_No = IIf(Trim(ls_SupplySeqNo) = 0, Null, Trim(ls_SupplySeqNo))
    rs_part_supply!Last_Update = Now
    rs_part_supply!last_user = userLogin
    rs_part_supply.update
    
    '#Init Control Cls
    'FromControlCls = l_stock_warehouse
    FromControlCls = cbo_warehouse.Column(2)
    ItemControlCls = stockcontrol_cls
    
    '#Check if item influence the stock or not
    If ItemControlCls = "01" Then
        '# Insert data into stock Master
        Call up_UpdateStockMaster(Format(DTPicker3.Value, "yyyy-MM-dd"), ls_ClosingMonth, ls_ClosingYear, Trim(cbo_warehouse), Trim(cbo_location), Trim(txt_item_code), CDbl(txt_qty), Trim(cbo_supply), Trim(l_stock_location), "", "I", "", "", False, False, True, db2)
    End If
    
    db2.CommitTrans
    
    If UCase(Trim(cbo_Replacement.Text)) = "YES" Then
        If Trim(cbo_ReplacementWarehouseCode.Text) <> ls_ReplacementWarehouseCode Then
            
            ls_ReplacementWarehouseCode = Trim(cbo_ReplacementWarehouseCode.Text)
            ls_FromWarehouseCode = Trim(cbo_warehouse.Text)
            ls_ToWarehouseCode = Trim(cbo_location.Text)
            ls_SupplyDate = Trim(Format(DTPicker3.Value, "yyyy-MM-dd"))
            ls_SupplyCls = Trim(cbo_supply.Text)
            
            ls_sql = " SELECT Seq_No FROM Part_Supply " & vbCrLf & _
                    " WHERE FromWarehouse_Code = '" & ls_FromWarehouseCode & "' " & vbCrLf & _
                    " AND ToWarehouse_Code = '" & ls_ToWarehouseCode & "' " & vbCrLf & _
                    " AND ChildSupply_Date = '" & ls_SupplyDate & "' AND Supply_Cls = '" & ls_SupplyCls & "' "
                    
            If rs.State = adStateOpen Then rs.Close
            Set rs = db2.Execute(ls_sql)
            If Not rs.EOF Then
                ls_SupplySeqNo = Trim(rs!Seq_no)
            End If
            
            rs.Close
            Set rs = Nothing
            
            cbo_warehouse.Text = ls_ReplacementWarehouseCode
            cbo_location.Text = ls_FromWarehouseCode
            
            Call Cmd_Submit_Click
            
        Else
            cbo_location.Text = ls_ToWarehouseCode
            cbo_warehouse.Text = ls_FromWarehouseCode
        End If
        
        ls_ReplacementWarehouseCode = ""
        ls_FromWarehouseCode = ""
        ls_ToWarehouseCode = ""
    End If
    
    Call setting_grid

    Call clear_framebawah
    lbl_pesan.Caption = DisplayMsg(1000) '"Insert data success !"
    Me.MousePointer = vbDefault

ErrExit:
    
    Me.MousePointer = vbDefault
    cmd_submit.Enabled = True
    Exit Sub
    
ErrHandler:
    cmd_submit.Enabled = True
    db2.RollbackTrans
    lbl_pesan.Caption = "[" & err.number & "] " & err.Description
    err.clear
    Resume ErrExit

End Sub

Private Function SSeqNo()
Dim rsmax As New ADODB.Recordset
Dim strSQL As String

strSQL = "Select Max(seq_No) from part_supply"

Set rsmax = Db.Execute(strSQL)

SSeqNo = IIf(IsNull(rsmax(0)), 1, rsmax(0) + 1)

End Function

Private Sub cmdBrowser_Click()
 If txt_item_code.Enabled = True Then
  Me.MousePointer = vbHourglass
  frm_BrowseItem.getItemCode = txt_item_code.Text
  frm_BrowseItem.Show 1
  txt_item_code.Text = frm_BrowseItem.getItemCode
  Me.MousePointer = vbDefault
 End If
End Sub

Private Sub cmdSearch_Click()
    Dim i As Double
    
    lbl_pesan = ""
    
    If txtSearch = "" Or Grid1.Rows = 2 Then txtSearch.SetFocus: Exit Sub
    If Grid1.Row = Grid1.Rows - 1 Then i = 2 Else i = Grid1.Row + 1
    
    Do
        Select Case cboSearch.ListIndex
        Case 0
            Grid1.Col = bteColProdCode
            If UCase(Mid(Grid1.TextMatrix(i, bteColProdCode), 1, Len(txtSearch))) = UCase(txtSearch) Then
                Exit Do
            End If
        Case 1
            Grid1.Col = bteColDesc
            If InStr(UCase(Grid1.TextMatrix(i, bteColDesc)), UCase(txtSearch)) <> 0 Then
                Exit Do
            End If
        End Select
        i = i + 1
        If i = Grid1.Rows - 1 Then
            txtSearch = ""
            i = 2
            lbl_pesan = DisplayMsg(8012)
            Exit Do
        End If
    Loop
    
    Grid1.Row = i
    Grid1.TopRow = i
    Grid1.SetFocus
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
    If ErrMsg = "" Then
        Unload Me
    Else
        lbl_pesan.Caption = ErrMsg
    End If
End Sub

Private Sub DTPicker3_Change()
    lbl_pesan = ""
    lbl_pesan = validCombo
    If Trim(lbl_pesan) <> "" Then clearGrid: Exit Sub
    Call browseprice
    Call clear_framebawah
    Call Header
    'Call setting_grid
    
End Sub

Function validCombo() As String
    
    Dim j As Integer
    
    j = 0
    For i = 0 To cbo_warehouse.ListCount - 1
        If UCase(Trim(cbo_warehouse)) = UCase(Trim(cbo_warehouse.List(i, 0))) Then
            cbo_warehouse.Text = cbo_warehouse.List(i, 0)
          
                lbl_warehouse.Caption = cbo_warehouse.List(i, 1)
          
            lbl_pesan.Caption = ""
            j = 1
            Exit For
        End If
    Next
    
    If j = 0 Then
        lbl_warehouse.Caption = "": validCombo = DisplayMsg(4018) ' "Invalid warehouse code !"
        Exit Function
    End If
    
    j = 0
    For i = 0 To cbo_location.ListCount - 1
        If UCase(Trim(cbo_location)) = UCase(Trim(cbo_location.List(i, 0))) Then
            cbo_location.Text = cbo_location.List(i, 0)
            lbl_location.Caption = cbo_location.List(i, 1)
            lbl_pesan.Caption = ""
            j = 1
            Exit For
        End If
    Next
    
    If j = 0 Then
        lbl_location.Caption = "":  validCombo = DisplayMsg(4014) '"Invalid location code !"
        Exit Function
    End If
    
    j = 0
    For i = 0 To cbo_supply.ListCount - 1
        If UCase(Trim(cbo_supply)) = UCase(Trim(cbo_supply.List(i, 0))) Then
            cbo_supply.Text = cbo_supply.List(i, 0)
            lbl_pesan.Caption = ""
            j = 1
            Exit For
        End If
    Next
    
    If j = 0 Then
        validCombo = DisplayMsg(4056) '"Invalid supply clasification !"
        Exit Function
    End If

End Function

Function validasi() As Boolean
    
    Dim j As Integer
    
    j = 0
    For i = 0 To cbo_warehouse.ListCount - 1
        If UCase(Trim(cbo_warehouse)) = UCase(Trim(cbo_warehouse.List(i, 0))) Then
            cbo_warehouse.Text = cbo_warehouse.List(i, 0)
            lbl_warehouse.Caption = cbo_warehouse.List(i, 1)
            lbl_pesan.Caption = ""
            j = 1
            Exit For
        End If
    Next
    If j = 0 Then
        lbl_warehouse.Caption = "": lbl_pesan.Caption = DisplayMsg(4018) '"Invalid warehouse code !"
        'Call set_item
        validasi = False
        cbo_warehouse.SetFocus
        Exit Function
    End If
    
    j = 0
    For i = 0 To cbo_location.ListCount - 1
        If UCase(Trim(cbo_location)) = UCase(Trim(cbo_location.List(i, 0))) Then
            cbo_location.Text = cbo_location.List(i, 0)
            lbl_location.Caption = cbo_location.List(i, 1)
            lbl_pesan.Caption = ""
            j = 1
            Exit For
        End If
    Next
    If j = 0 Then
        lbl_location.Caption = "": lbl_pesan.Caption = DisplayMsg(4014) '"Invalid location code !"
        validasi = False
        cbo_location.SetFocus
        Exit Function
    End If
    
    j = 0
    For i = 0 To cbo_supply.ListCount - 1
        If UCase(Trim(cbo_supply)) = UCase(Trim(cbo_supply.List(i, 0))) Then
            cbo_supply.Text = cbo_supply.List(i, 0)
            lbl_pesan.Caption = ""
            j = 1
            Exit For
        End If
    Next
    If j = 0 Then
        lbl_pesan.Caption = DisplayMsg(4056) '"Invalid supply clasification !"
        validasi = False
        cbo_supply.SetFocus
        Exit Function
    End If
    
    If Trim(txt_item_code) = "" Then
        lbl_pesan.Caption = DisplayMsg(1009) '"Please insert Product Code !"
        validasi = False
        txt_item_code.SetFocus
        Exit Function
    End If
    j = 0
    For i = 0 To txt_item_code.ListCount - 1
        If UCase(Trim(txt_item_code)) = UCase(Trim(txt_item_code.List(i, 0))) Then
            txt_item_code.Text = txt_item_code.List(i, 0)
            lbl_pesan.Caption = ""
            j = 1
            Exit For
        End If
    Next
    If j = 0 Then
        '**** cek Data cocok / tdk dgn Database
        Dim rsDB As New ADODB.Recordset
        rsDB.Open "select Item_Code from Item_Master where Item_Code='" & Trim(txt_item_code) & "'", Db, adOpenKeyset, adLockOptimistic
        If rsDB.EOF = True Then
            lbl_pesan.Caption = DisplayMsg(4003) '"Invalid Product Code !"
            validasi = False
            txt_item_code.SetFocus
            Exit Function
        End If
        rsDB.Close
    End If
    
    '===============================================================================
    'Qty
    If Trim(txt_qty) = "" Then
        lbl_pesan.Caption = DisplayMsg(1012) '"Please insert quantity !"
        validasi = False
        txt_qty.SetFocus
        Exit Function
    End If
    
    If IsNumeric(txt_qty) = False Then
        validasi = False
        txt_qty.SetFocus
        lbl_pesan = DisplayMsg(4044) 'Please Input Valid Quantity
        Exit Function
    End If
    
    If CDbl(txt_qty) > gd_MaxQty Then
        lbl_pesan = DisplayMsg(4045) & " " & gd_MaxQty
        validasi = False
        txt_qty.SetFocus
        Exit Function
    End If
    
    If CDbl(Trim(txt_qty)) < 0 And Trim(cbo_supply) <> "S1" And Trim(cbo_supply) <> "S" Then
        lbl_pesan.Caption = DisplayMsg(4057) '"Can't input minus quantity for reject or loss !"
        validasi = False
        txt_qty.SetFocus
        Exit Function
    End If
    
     
    If cbo_warehouse.Text = "WH-001" Then
        If CDbl(txt_qty) > txtCurrentQty Then
                lbl_pesan.Caption = DisplayMsg(4044)
                  validasi = False
            txt_qty.SetFocus
            Exit Function
        ElseIf txtCurrentQty = 0 Then
            lbl_pesan.Caption = DisplayMsg(4044)
              validasi = False
            txt_qty.SetFocus
            Exit Function
        End If
    ElseIf cbo_warehouse.Text = "WH-003" Then
        If CDbl(txt_qty) > txtCurrentQty Then
                lbl_pesan.Caption = DisplayMsg(4044)
                validasi = False
                txt_qty.SetFocus
                Exit Function
        ElseIf txtCurrentQty = 0 Then
            lbl_pesan.Caption = DisplayMsg(4044)
            validasi = False
            txt_qty.SetFocus
            Exit Function
        End If
    ElseIf cbo_warehouse.Text = "WH-004" Then
        If CDbl(txt_qty) > txtCurrentQty Then
                lbl_pesan.Caption = DisplayMsg(4044)
                validasi = False
                txt_qty.SetFocus
                Exit Function
        ElseIf txtCurrentQty = 0 Then
            lbl_pesan.Caption = DisplayMsg(4044)
            validasi = False
            txt_qty.SetFocus
            Exit Function
        End If
    ElseIf cbo_warehouse.Text = "WH-017" Then
        If CDbl(txt_qty) > txtCurrentQty Then
                lbl_pesan.Caption = DisplayMsg(4044)
                validasi = False
                txt_qty.SetFocus
                Exit Function
        ElseIf txtCurrentQty = 0 Then
            lbl_pesan.Caption = DisplayMsg(4044)
            validasi = False
            txt_qty.SetFocus
            Exit Function
        End If
    End If
    '===============================================================================
    
    If Trim(cbo_curr) = "" Then
        If bteHakPrice = 0 Then
            cbo_curr = uf_GetCurrencyDescription(gs_DefaultCurrencyCode)
        Else
            lbl_pesan.Caption = DisplayMsg(1028) '"Please insert currency !"
            validasi = False
            cbo_curr.SetFocus
            Exit Function
        End If
    End If
    j = 0
    For i = 0 To cbo_curr.ListCount - 1
        If UCase(Trim(cbo_curr)) = UCase(Trim(cbo_curr.List(i, 0))) Then
            cbo_curr.Text = cbo_curr.List(i, 0)
            lbl_pesan.Caption = ""
            j = 1
            Exit For
        End If
    Next
    If j = 0 Then
        lbl_pesan.Caption = DisplayMsg(4005) '"Invalid currency clasification !"
        validasi = False
        cbo_curr.SetFocus
        Exit Function
    End If
    
    '===============================================================================
    'Price
    If Trim(cbo_price) = "" Then
        If bteHakPrice = 0 Then
            cbo_price = 0
        Else
            lbl_pesan.Caption = DisplayMsg(1029) '"Please insert price !"
            validasi = False
            cbo_price.SetFocus
            Exit Function
        End If
    End If
    
    If CDbl(cbo_price) > gd_MaxPrice Then
        validasi = False
        cbo_price.SetFocus
        lbl_pesan = DisplayMsg(4048) & " " & gd_MaxPrice
        Exit Function
    End If
    '===============================================================================
    
    If CDbl(txt_amount) > gd_MaxAmount Then
        lbl_pesan = DisplayMsg(4051) & " " & gd_MaxAmount
        validasi = False
        Exit Function
    End If
        
'    If Trim(txtSJNo) = "" Then
'        lbl_pesan = DisplayMsg(1036)  '"Please input Surat Jalan No"
'        txtSJNo.SetFocus
'        validasi = False
'        Exit Function
'    End If
        
    If txtNoSeri.Text = "" Then lbl_pesan.Caption = DisplayMsg("0001") & " No Seri ! ":  txtNoSeri.SetFocus: validasi = False: Exit Function
    
    If cboBCType = "2.6.1" Then
'        If txtDONo.Enabled = True And txtDONo.Text = "" Then lbl_pesan.Caption = DisplayMsg("0001") & " DO No. ! ":  txtDONo.SetFocus: validasi = False: Exit Function
        If Trim(txtDONo.Text) = "" Then
                txtDONo.Text = GetContractNo(txt_remarks.Text)
        End If
    End If
    
    validasi = True
    lbl_pesan.Caption = ""

End Function

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"
    
    If db2.State <> adStateClosed Then db2.Close
    db2.Open Db.ConnectionString
    Call koneksi
    DTPicker3.Value = Format(Date, "dd MMM yyyy")
    lbl_warehouse.Caption = ""
    lbl_location.Caption = ""
    lbl_pesan.Caption = ""
    lbl_supply.Caption = ""
    cbo_location = ""
    cbo_supply = ""
    cbo_warehouse = ""
    cbo_Replacement.Text = ""
    cbo_ReplacementWarehouseCode.Text = ""
    lbl_ReplacementWarehouseCode.Caption = ""
    bteHakPrice = hakPrice(Me.Name)
    cbo_curr.Visible = (bteHakPrice = 1)
    cbo_price.Visible = (bteHakPrice = 1)
    txt_amount.Visible = (bteHakPrice = 1)
    Label10(4).Visible = (bteHakPrice = 1)
    Label10(5).Visible = (bteHakPrice = 1)
    Label10(6).Visible = (bteHakPrice = 1)
    Call setting
    Call Header
    'Call setting_grid
    Call comboBCtype
    DtBCDate = Format(Date, "dd mmm yyyy")
    Call set_item
    Call clear_framebawah
    lbl_pesan.Caption = ""
    cbo_warehouse.DataChanged = True
    cbo_location.DataChanged = True
'    Call adtocombo
'    Call adtocbopono
    Call delivery
    
    With cboSearch
        .AddItem "Item Code"
        .AddItem "Description"
        .ListIndex = 0
    End With
    
    
End Sub
'Sub adtocbopono()
'Dim sqlno As String
'Dim rsno As New Recordset
'If Trim(txtName.Text) = "" Then Exit Sub
'    sqlno = " select * From Purchaseorder_master " & vbCrLf & _
'                           "Where " & vbCrLf & _
'                            IIf(cboSupp.Text = "ALL", "", "and  Supplier_Code='" & cboSupp.Text & "' ") & vbCrLf & _
'                           ""
'End Sub
Sub delivery()


Dim rs As New ADODB.Recordset
sql = " Select Trade_name, Address1 as A from trade_master where trade_cls in ('2', '3') order by trade_code" & vbCrLf
            
Set rs = Db.Execute(sql)

With CboDelivery
.clear
'.ColumnCount = 1
.AddItem ""
i = 1
Do Until rs.EOF
    .AddItem ""
    .List(i, 0) = Trim(rs!trade_name)
    i = i + 1
    rs.MoveNext
Loop

'.ColumnWidths = "60 pt; 300 pt"
.ListWidth = 300
.ListRows = 15
End With
End Sub
'Sub adtocombo()
'Dim rs As New ADODB.Recordset
'sql = "SELECT Trade_Code, Trade_Name FROM Trade_Master where trade_cls='2' or trade_cls='3'"
'Set rs = Db.Execute(sql)
'
'With cboSupp
'
'.clear
'.ColumnCount = 2
'.ColumnWidths = "80 pt;300 pt"
'.ListWidth = 380
'.ListRows = 15
'.AddItem ""
'.List(0, 0) = strAll
'.List(0, 1) = strAll
'i = 1
'Do Until rs.EOF
'    .AddItem ""
'    .List(i, 0) = Trim(rs!Trade_Code)
'    .List(i, 1) = Trim(rs!trade_name)
'    i = i + 1
'    rs.MoveNext
'Loop
'.ListIndex = 0
'End With
'End Sub
Private Sub clear_framebawah(Optional lb_forceClear As Boolean)
           
    If gb_AllowClearInputArea_PartSupplyUnschedule = True Or lb_forceClear = True Then
        txtsjno = ""
        txtBCNo = ""
        cboBCType.Text = ""
        txt_remarks = ""
        txt_maker = ""
        cbo_curr = "02"
        cbo_price = ""
        txt_desc = ""
        txt_item_code = ""
        txtCurrentQty = ""
        DtBCDate = Format(Date, "dd mmm yyyy")
        txtRegisterNo.Text = ""
    End If
    Call browseprice
    
    txt_amount = ""
    txt_item_code.Enabled = True
    txt_qty = ""
    txt_item_code.Text = ""
    txtNoSeri.Text = ""
    txtDONo.Text = ""
    cbo_curr.clear
    'cbo_Replacement.Text = cbo_Replacement.Column(0, 0)
End Sub

Private Sub set_itemInfo()
    
    Dim j  As Integer
    
    j = 0
    For j = 0 To txt_item_code.ListCount - 1
        If UCase(Trim(txt_item_code)) = UCase(Trim(txt_item_code.List(j, 0))) Then j = 1: Exit For
    Next
    
    If j = 0 Then
        lbl_pesan.Caption = DisplayMsg(4006) '"Data not found !"
        Exit Sub
    Else
        lbl_pesan.Caption = ""
    End If
    
    If rs_item.State <> adStateClosed Then rs_item.Close
    rs_item.Open "select * from item_master where item_code='" & Trim(txt_item_code) & "' ", Db, adOpenKeyset, adLockOptimistic
    If rs_item.EOF = False Or rs_item.BOF = False Then
        txt_item_code = Trim(rs_item!Item_Code)
        txt_maker = Trim(rs_item!MakerItem_Code)
        txt_desc = uf_GetItemDescription(Trim(txt_item_code))
        'Text1 = IIf(IsNull(Trim(rs_item!Address)), "", Trim(rs_item!Address))
        
        stockcontrol_cls = Trim(rs_item!stockcontrol_cls)
        lbl_pesan.Caption = ""
    Else
        txt_maker = ""
        txt_desc = ""
'       Text1 = ""
        cbo_curr.clear
        cbo_price.Text = ""
'        Call cbo_price_Change
        lbl_pesan.Caption = DisplayMsg(4006) '"Data not found !"
    End If
    txt_amount.Text = ""

End Sub

Private Sub browseprice()
    
    Dim sql2 As String
    Dim rs2 As New Recordset
    Dim currIndex As Integer
'    currIndex = cbo_curr.ListIndex

'Remark 20240114
'    sql2 = "select trade_code, priority_cls, currency_code, price from price_master where " & _
'        "item_code='" & txt_item_code.Text & "' and price_cls='03' and priority_cls='1' and currency_code='" & uf_GetCurrencyCode(IIf(cbo_curr.Text = "", "USD", cbo_curr.Text)) & "' " & _
'        " and start_date<='" & Format(DTPicker3.Value, "yyyymmdd") & "' and end_date>='" & _
'        Format(DTPicker3.Value, "yyyymmdd") & "'  and Trade_Code='000000' order by trade_code desc, priority_cls desc"
        
    sql2 = "SELECT TOP 1 " & _
            "Trade_Code = PR.Supplier_Code , " & _
            "Priority_Cls = '0' , " & _
            "Currency_Code = CC.Description , " & _
            "Price = PR.Price " & _
            "FROM    dbo.Part_Receipt PR LEFT JOIN dbo.Curr_Cls CC ON PR.Currency_Code = CC.Curr_Cls " & _
            "WHERE   Item_Code = '" & txt_item_code.Text & "'   " & _
            "AND Currency_Code = '" & uf_GetCurrencyCode(IIf(cbo_curr.Text = "", "USD", cbo_curr.Text)) & "' " & _
            "ORDER BY Receipt_Date DESC"
        
    Set rs2 = Db.Execute(sql2)
    
    If Not rs2.EOF Then
        With cbo_price
            .clear
            .columnCount = 3
            .ColumnWidths = "70pt;70pt;0pt"
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
                rs2.MoveNext
                i = i + 1
            Loop
        End With
    Else
        Call up_FillCombo(cbo_curr, "curr_cls", "description, curr_cls", "WHERE Curr_Cls = '02' ") 'USD
        cbo_curr.ColumnWidths = "60 pt;0 pt"
        If txt_item_code.matchFound = True Then
            cbo_curr.Text = cbo_curr.Column(0, 0)
        End If
    End If
    
    If cbo_price.ListCount > 0 Then
        cbo_price.ListIndex = 0
        Call cbo_price_Click
    ElseIf cbo_price.Text = "" Then
       cbo_curr.ListIndex = -1
    End If
    
    cbo_curr.DataChanged = False
    
    Call up_FillCombo(cbo_curr, "curr_cls", "description, curr_cls", "WHERE Curr_Cls = '02' ") 'USD
    cbo_curr.ColumnWidths = "60 pt;0 pt"
    If txt_item_code.matchFound = True Then
        cbo_curr.Text = cbo_curr.Column(0, 0)
    End If
    
'REMARK 20240114
'    If txt_item_code.MatchFound = True Then
'        cbo_curr.ListIndex = currIndex
'    End If
    
    cbo_curr.DataChanged = True
End Sub

Private Sub setting()
    'Setting Combo From Warehouse Code
    cbo_warehouse.clear
    cbo_warehouse.columnCount = 3
    cbo_warehouse.TextColumn = 1
    
    i = 0
    If rs_warehouse.EOF = False Or rs_warehouse.BOF = False Then
        rs_warehouse.MoveFirst
        While rs_warehouse.EOF = False
            cbo_warehouse.AddItem ""
            cbo_warehouse.List(i, 0) = Trim(rs_warehouse!wh_code)
            cbo_warehouse.List(i, 1) = Trim(rs_warehouse!WH_Name)
            cbo_warehouse.List(i, 2) = Trim(rs_warehouse!stockcontrol_cls)
            rs_warehouse.MoveNext
            i = i + 1
        Wend
        cbo_warehouse.ColumnWidths = "50 pt; 175 pt; 0 pt"
        cbo_warehouse.ListWidth = 225
    End If
    
    'Setting Combo To Warehouse Code
    cbo_location.clear
    cbo_location.columnCount = 3
    cbo_location.TextColumn = 1
    
    i = 0
    If rs_warehouse.EOF = False Or rs_warehouse.BOF = False Then
        rs_warehouse.MoveFirst
        While rs_warehouse.EOF = False
            cbo_location.AddItem ""
            cbo_location.List(i, 0) = Trim(rs_warehouse!wh_code)
            cbo_location.List(i, 1) = Trim(rs_warehouse!WH_Name)
            cbo_location.List(i, 2) = Trim(rs_warehouse!stockcontrol_cls)
            rs_warehouse.MoveNext
            i = i + 1
        Wend
        cbo_location.ColumnWidths = "50 pt; 175 pt;0 pt"
        cbo_location.ListWidth = 225
    End If
    
    'Setting Combo Supply Cls
    cbo_supply.clear
    cbo_supply.columnCount = 2
    cbo_supply.TextColumn = 1
    cbo_supply.AddItem ""
    cbo_supply.List(0, 0) = "S1"
    cbo_supply.List(0, 1) = "Supply"
    cbo_supply.AddItem ""
    cbo_supply.List(1, 0) = "S"
    cbo_supply.List(1, 1) = "Consumption"
    cbo_supply.AddItem ""
    cbo_supply.List(2, 0) = "L"
    cbo_supply.List(2, 1) = "Loss"
    cbo_supply.AddItem ""
    cbo_supply.List(3, 0) = "RJ"
    cbo_supply.List(3, 1) = "Reject"
    cbo_supply.ColumnWidths = "25 pt; 75 pt"
    cbo_supply.ListWidth = 100
    cbo_supply = "S1"
    
    cbo_Replacement.clear
    cbo_Replacement.columnCount = 2
    cbo_Replacement.TextColumn = 1
    cbo_Replacement.AddItem ""
    cbo_Replacement.List(0, 0) = "No"
    cbo_Replacement.List(0, 1) = "0"
    cbo_Replacement.AddItem ""
    cbo_Replacement.List(1, 0) = "Yes"
    cbo_Replacement.List(1, 1) = "1"
    cbo_Replacement.ColumnWidths = "50 pt; 0 pt"
    cbo_Replacement.Text = cbo_Replacement.Column(0, 0)
    cbo_Replacement.ListWidth = 50
    
    'Setting Combo Curr
    Call up_FillCombo(cbo_curr, "curr_cls", "description, curr_cls", "WHERE Curr_Cls = '02'")
    cbo_curr.ColumnWidths = "60 pt;0 pt"

End Sub

Private Sub koneksi()
    Dim SqlW As String
    rs_part_supply.Open "select Top 1 * from part_supply", Db, adOpenKeyset, adLockOptimistic
    SqlW = " select * from (select wh_code,wh_name,stockControl_cls from warehouse_master " & _
        " union all " & _
        " select distinct(manufacture_line.manufacture_code)wh_code,trade_name wh_name,stockControl_Cls='01' from manufacture_line join trade_master on manufacture_line.manufacture_code=trade_master.trade_code)tbJ order by wh_code "
    rs_warehouse.Open SqlW, Db, adOpenKeyset, adLockOptimistic
    rs_trade_master.Open "select * from trade_master where trade_cls='1'", Db, adOpenKeyset, adLockOptimistic
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'db2.Close
    If rs_part_supply.State <> adStateClosed Then rs_part_supply.Close
    If rs_warehouse.State <> adStateClosed Then rs_warehouse.Close
    If rs_trade_master.State <> adStateClosed Then rs_trade_master.Close
End Sub

Private Sub cbo_price_Change()
    If Trim(cbo_price) = "." Then cbo_price = "": Exit Sub
    If Trim(cbo_price) = "" Then txt_amount = "": Exit Sub
    If Left(Trim(cbo_price), 1) = "," Then cbo_price = Right(Trim(cbo_price), Len(Trim(cbo_price)) - 1)
    If cbo_price.Text <> "" And IsNumeric(cbo_price) = True Then
        If txt_qty <> "" And IsNumeric(txt_qty) = True Then
            txt_amount.Text = Format(CDbl(cbo_price.Text) * CDbl(txt_qty.Text), gs_formatAmount)
        Else
            txt_amount = Format(0, gs_formatAmount)
        End If
    Else
        txt_amount = Format(0, gs_formatAmount)
    End If
    If Right(Trim(txt_amount), 1) = "." Then txt_amount = Left(Trim(txt_amount), Len(Trim(txt_amount)) - 1)
End Sub

Private Sub cbo_price_Click()
    If cbo_price.ListIndex <> -1 Then
        cbo_curr.Text = uf_GetCurrencyCode(cbo_price.Column(2))
        If txt_qty <> "" Then txt_amount.Text = Format(CDbl(cbo_price.Text) * CDbl(txt_qty.Text), gs_formatAmount)
    End If
End Sub

Private Sub cbo_price_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
  If KeyCode = 13 Then cbo_price_Click
End Sub

Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    Dim TextGrid As String, L_price As String
    Dim i As Integer
    
    With Grid1
    
        TextGrid = Grid1.Text
        If TextGrid = "S" Then
            
            For i = 1 To .Rows - 1
                If .TextMatrix(i, bteColSelect) = "S" Then
                    txt_item_code = .TextMatrix(i, bteColProdCode)
                    l_item_code_update = .TextMatrix(i, bteColProdCode)
                    Call txt_item_code_Change
                    
                    txt_qty = .TextMatrix(i, bteColQty)
                    l_update_stock = .TextMatrix(i, bteColQty)
                    l_seqNo = Val(Trim(.TextMatrix(i, bteColSeqNo)))
                    
                    Call browseprice
                    L_price = Format(.TextMatrix(i, bteColPrice), gs_formatPrice)
                    If Right(Trim(L_price), 1) = "." Then _
                    L_price = Left(Trim(L_price), Len(Trim(L_price)) - 1)
                    
                    cbo_curr = .TextMatrix(i, bteColCurr)
                    cbo_price = L_price
                    txt_amount = .TextMatrix(i, bteColAmount)
    
                    ItemControlCls = Trim(.TextMatrix(i, bteColItemControl))
                    txt_remarks = .TextMatrix(i, bteColRemark)
                    txtsjno = .TextMatrix(i, bteColSJNo)
                    txtBCNo = .TextMatrix(i, bteColBcNo)
                    cboBCType = .TextMatrix(i, bteColBctype)
                    txtRegisterNo.Text = .TextMatrix(i, bteColNoRegister)
                    txtNoSeri.Text = .TextMatrix(i, bteColNoSeri)
                    txtDONo.Text = .TextMatrix(i, bteColDONo)
                    
                    If .TextMatrix(i, bteColBCDate) = "" Then
                        DtBCDate = Format(Date, "dd mmm yyyy")
                    Else
                        DtBCDate = .TextMatrix(i, bteColBCDate)
                    End If
                    
                
                    Exit For
                End If
            Next
            txt_item_code.Enabled = False
        
        Else
        
            txt_item_code.Enabled = True
            Call clear_framebawah
    
        End If
        .TextMatrix(Row, Col) = TextGrid
        
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

Private Sub Grid1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Grid1.Col <> bteColSelect Then Cancel = True: Exit Sub
End Sub

Private Sub Grid1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Grid1.Col = bteColSelect Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii <> Asc("S") And KeyAscii <> Asc("D") And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then
            KeyAscii = 0
        End If
        If KeyAscii = Asc(".") Then KeyAscii = 0
    
        If KeyAscii = Asc("S") Then
            For i = 1 To Grid1.Rows - 1
                Grid1.TextMatrix(i, bteColSelect) = ""
            Next
        End If
        If KeyAscii = Asc("D") Then
            For i = 1 To Grid1.Rows - 1
                If Grid1.TextMatrix(i, bteColSelect) <> "D" Then Grid1.TextMatrix(i, bteColSelect) = ""
            Next
        End If
    End If
End Sub

Private Sub txt_amount_GotFocus()
    SendKeys vbTab
End Sub

Private Sub txt_item_code_Change()
Dim rsDB As New ADODB.Recordset
Dim sql As String

    Call set_itemInfo
    cbo_price = ""
    txt_qty = ""
    If txt_item_code.Text <> "" Then
        cbo_price.Text = "0"
        cbo_price = Format(cbo_price.Text, gs_formatPrice)
       If txt_item_code.matchFound = True Then
        Text2.Text = txt_item_code.List(txt_item_code.ListIndex, 4)
        Call browseprice
       End If
    End If
    
    Dim fieldCurrent As String
    Dim Idx As Integer
    Idx = DateDiff("M", uf_GetLastClosing("fulldate"), DTPicker3.Value)
    
    If Idx = 0 Then
        fieldCurrent = "LM_Current"
    ElseIf Idx = 1 Then
        fieldCurrent = "TM_Current"
    ElseIf Idx = 2 Then
        fieldCurrent = "NM_Current"
    Else
        fieldCurrent = 0
    End If
    
    up_GetNoSeri (txt_item_code.Text)
     
'    sql = "SELECT fieldCurrent = '" & fieldCurrent & "' FROM dbo.Stock_Master WHERE Warehouse_Code = '" & Trim(cbo_warehouse.Text) & "' AND Item_Code = '" & Trim(txt_item_code.Text) & "'"
'    rsDB.Open sql, Db, adOpenForwardOnly, adLockReadOnly
'    'rsDB.Open "SELECT fieldCurrent = '" & fieldCurrent & "' FROM dbo.Stock_Master WHERE Warehouse_Code = '" & Trim(cbo_warehouse.Text) & "' AND Item_Code = '" & Trim(txt_item_code.Text) & "'", Db, adOpenKeyset, adLockOptimistic
'    'If Not rsDB.EOF = True Then 'IIf(IsNull(rs_PR!Price), 0, Trim(rs_PR!Price))
'    If rsDB.EOF = False Or rsDB.BOF = False Then
'        txtCurrentQty.Text = IIf(IsNull(rsDB!fieldCurrent), 0, Format(rsDB!fieldCurrent, gs_formatQty)) 'Format(rsDB!fieldCurrent, gs_formatQty)
''    Else
''        txtCurrentQty.Text = Format(0, gs_formatQty)
'    End If
'
'    rsDB.Close

    'Dim rsDB As New ADODB.Recordset
    rsDB.Open "SELECT fieldCurrent = " & fieldCurrent & " FROM dbo.Stock_Master WHERE Warehouse_Code = '" & Trim(cbo_warehouse.Text) & "' AND Item_Code = '" & Trim(txt_item_code.Text) & "'", Db, adOpenKeyset, adLockOptimistic
    If Not rsDB.EOF = True Then
        txtCurrentQty.Text = Format(rsDB!fieldCurrent, gs_formatQty)
    Else
        txtCurrentQty.Text = Format(0, gs_formatQty)
    End If
    rsDB.Close
    
    
End Sub

Private Sub txt_item_code_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
'    If KeyCode = 13 Then
'        Call set_itemInfo
'        cbo_price = ""
'        txt_qty = ""
'        Call BrowsePrice
'    End If
End Sub

Private Sub txt_item_code_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txt_maker_GotFocus()
    SendKeys vbTab
End Sub

Private Sub txt_qty_Change()
    lbl_pesan.Caption = ""
    If Trim(txt_qty) = "." Then txt_qty = "": Exit Sub
    If Trim(txt_qty) = "" Then txt_amount = "": Exit Sub
    If Trim(txt_qty) = "" Or Trim(cbo_price) = "" Then Exit Sub
    If Left(Trim(txt_qty), 1) = "," Then txt_qty = Right(Trim(txt_qty), Len(Trim(txt_qty)) - 1)
   ' If IsNumeric(Trim(txt_qty)) = False Then Exit Sub
    If IsNumeric(Trim(txt_qty)) = True And IsNumeric(cbo_price) = True Then
        txt_amount = Format(CDbl(txt_qty) * CDbl(cbo_price), gs_formatAmount)
    Else
        txt_amount = Format(0, gs_formatAmount)
    End If
    If Right(Trim(txt_amount), 1) = "." Then txt_amount = Left(Trim(txt_amount), Len(Trim(txt_amount)) - 1)
    'IIf(IsNull(rsDB!fieldCurrent), 0, Format(rsDB!fieldCurrent, gs_formatQty))
    If cbo_warehouse.Text = "WH-001" Then
        If CDbl(txt_qty) > txtCurrentQty Then  'CDbl(IIf(txtCurrentQty.Text = "", 0, txtCurrentQty.Text)) Then
                lbl_pesan.Caption = DisplayMsg(4044)
                txt_qty.SetFocus
                Exit Sub
        ElseIf txtCurrentQty = 0 Then
            lbl_pesan.Caption = DisplayMsg(4044)
            txt_qty.SetFocus
            Exit Sub
        End If
    ElseIf cbo_warehouse.Text = "WH-003" Then
        If CDbl(txt_qty) > txtCurrentQty Then
                lbl_pesan.Caption = DisplayMsg(4044)
                txt_qty.SetFocus
                Exit Sub
        ElseIf txtCurrentQty = 0 Then
            lbl_pesan.Caption = DisplayMsg(4044)
            txt_qty.SetFocus
            Exit Sub
        End If
    ElseIf cbo_warehouse.Text = "WH-004" Then
        If CDbl(txt_qty) > txtCurrentQty Then
                lbl_pesan.Caption = DisplayMsg(4044)
                txt_qty.SetFocus
                Exit Sub
        ElseIf txtCurrentQty = 0 Then
            lbl_pesan.Caption = DisplayMsg(4044)
            txt_qty.SetFocus
            Exit Sub
        End If
    End If
End Sub

Private Sub txt_qty_KeyPress(KeyAscii As Integer)
 If InStr(1, "0123456789.-", Chr(KeyAscii)) <= 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub txt_qty_LostFocus()
    txt_qty = Format(txt_qty, gs_formatQty)
End Sub

Private Sub txt_remarks_GotFocus()
    l_Remarks = txt_remarks.Text
End Sub

Private Sub txt_remarks_LostFocus()
    On Error GoTo ErrHandler
    
    If l_Remarks <> txt_remarks.Text Then
        If cboBCType = "2.6.1" Then
            If Trim(txtDONo.Text) = "" Then
                txtDONo.Text = GetContractNo(txt_remarks.Text)
            End If
        Else
            txtDONo.Text = ""
        End If
    End If
    Exit Sub
    
ErrHandler:
    MsgBox "Terjadi kesalahan saat memproses Contract No: " & err.Description, vbExclamation
End Sub

Private Function GetContractNo(ByVal Remarks As String) As String
    On Error GoTo ErrHandler
    
    Dim rs As New ADODB.Recordset
    Dim sql As String
    
    GetContractNo = "" ' default kalau tidak ada hasil
    
    sql = "EXEC dbo.sp_POContractNo_Get '" & Trim(Remarks) & "'"
    
    If rs.State <> adStateClosed Then rs.Close
    rs.Open sql, Db, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.EOF Then
        GetContractNo = rs.Fields("Contract_No").Value
    End If
    
    rs.Close
    Exit Function
    
ErrHandler:
    MsgBox "Error di GetContractNo: " & err.Description, vbExclamation
End Function


Private Sub txtNoSeri_KeyPress(KeyAscii As Integer)
    ' Hanya izinkan angka (48-57) dan tombol Backspace (8)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> 8 Then
        KeyAscii = 0 ' Batalkan input
    End If
End Sub

Private Sub txtpono_Change()
'Call adtocbopono
End Sub

Private Sub txtSJNo_Change()
    txtRegisterNo.Text = ""
End Sub

Private Sub txtSJNo_GotFocus()
    l_SJNo = txtsjno.Text
End Sub

Private Sub txtSJNo_LostFocus()
    Dim rsGetNoSeri As New ADODB.Recordset
        
    If l_SJNo <> txtsjno Then
        If Trim(txtsjno.Text) <> "" Then
            sql = "EXEC dbo.sp_GetNoRegister '" & DTPicker3.Value & "', 'S1', '" & userLogin & "' "
                    
            If rsGetNoSeri.State <> adStateClosed Then rsGetNoSeri.Close
            rsGetNoSeri.Open sql, Db, adOpenForwardOnly, adLockReadOnly
            
            If Not rsGetNoSeri.EOF Then
               txtRegisterNo.Text = rsGetNoSeri.Fields("No_Register")
            End If
        End If
    End If
End Sub

Private Sub up_GetNoSeri(pItemCode As String)
Dim rsGetNo As New ADODB.Recordset

    sql = "EXEC dbo.sp_PartSupply_GetNoSeri @SJNo = '" & Trim(txtsjno.Text) & "', " & vbCrLf & _
                    "  @Date = '" & Format(DtBCDate.Value, "yyyy-mm-dd") & "', " & vbCrLf & _
                    "  @ItemCode ='" & pItemCode & "' "
                    
                    If rsGetNo.State <> adStateClosed Then rsGetNo.Close
                    rsGetNo.Open sql, Db, adOpenForwardOnly, adLockReadOnly
                    
                    If Not rsGetNo.EOF Then
                       txtNoSeri.Text = IIf(IsNull(rsGetNo.Fields("No_Seri")), "1", Trim(rsGetNo.Fields("No_Seri")))

                    End If
        
End Sub
