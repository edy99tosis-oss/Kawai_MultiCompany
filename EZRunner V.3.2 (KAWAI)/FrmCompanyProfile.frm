VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmCompanyProfile 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FDDFE3&
   Caption         =   "Company Profile"
   ClientHeight    =   7755
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11760
   Icon            =   "FrmCompanyProfile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   11760
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Copy"
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
      Left            =   7547
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   6765
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtseqno 
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
      Left            =   3647
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   6765
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
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
      Left            =   8852
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   6765
      Visible         =   0   'False
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   390
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   1125
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   8705
      _Version        =   393216
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&General"
      TabPicture(0)   =   "FrmCompanyProfile.frx":0E42
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Staff (1)"
      TabPicture(1)   =   "FrmCompanyProfile.frx":0E5E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(1)=   "Frame5"
      Tab(1).Control(2)=   "Frame6"
      Tab(1).Control(3)=   "Frame7"
      Tab(1).Control(4)=   "Frame9"
      Tab(1).Control(5)=   "Frame10"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Staff (2)"
      TabPicture(2)   =   "FrmCompanyProfile.frx":0E7A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame17"
      Tab(2).Control(1)=   "Frame18"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Bank Account"
      TabPicture(3)   =   "FrmCompanyProfile.frx":0E96
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label26"
      Tab(3).Control(1)=   "Shape2"
      Tab(3).Control(2)=   "grid"
      Tab(3).Control(3)=   "Frame11"
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "NPWP"
      TabPicture(4)   =   "FrmCompanyProfile.frx":0EB2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame3"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "GL Codes"
      TabPicture(5)   =   "FrmCompanyProfile.frx":0ECE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame8"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "BC"
      TabPicture(6)   =   "FrmCompanyProfile.frx":0EEA
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame19"
      Tab(6).ControlCount=   1
      Begin VB.Frame Frame19 
         Caption         =   "BC"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Left            =   -74760
         TabIndex        =   118
         Top             =   600
         Width           =   5100
         Begin VB.TextBox txtBC 
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
            Left            =   1710
            MaxLength       =   25
            TabIndex        =   124
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAA"
            Top             =   1665
            Width           =   3135
         End
         Begin VB.TextBox txtBC 
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
            Left            =   1710
            MaxLength       =   25
            TabIndex        =   123
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAA"
            Top             =   1230
            Width           =   3135
         End
         Begin VB.TextBox txtBC 
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
            Left            =   1710
            MaxLength       =   25
            TabIndex        =   120
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAA"
            Top             =   800
            Width           =   3135
         End
         Begin VB.TextBox txtBC 
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
            Left            =   1710
            MaxLength       =   25
            TabIndex        =   119
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAA"
            Top             =   360
            Width           =   3135
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NIP 2"
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
            Left            =   270
            TabIndex        =   126
            Top             =   1710
            Width           =   465
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Person 2"
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
            Left            =   270
            TabIndex        =   125
            Top             =   1275
            Width           =   750
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NIP 1"
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
            Left            =   270
            TabIndex        =   122
            Top             =   840
            Width           =   465
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Person 1"
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
            Left            =   270
            TabIndex        =   121
            Top             =   405
            Width           =   750
         End
      End
      Begin VB.Frame Frame18 
         Caption         =   "Material Request"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4245
         Left            =   -69480
         TabIndex        =   111
         Top             =   480
         Width           =   5325
         Begin VB.Frame Frame16 
            Caption         =   "Supply"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   90
            TabIndex        =   115
            Top             =   1560
            Width           =   5100
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
               Height          =   285
               Index           =   28
               Left            =   1800
               MaxLength       =   25
               TabIndex        =   29
               Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAA"
               Top             =   800
               Width           =   3135
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
               Height          =   285
               Index           =   27
               Left            =   1800
               MaxLength       =   25
               TabIndex        =   28
               Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAA"
               Top             =   360
               Width           =   3135
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Person"
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
               Left            =   360
               TabIndex        =   117
               Top             =   795
               Width           =   585
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Position"
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
               Left            =   360
               TabIndex        =   116
               Top             =   360
               Width           =   660
            End
         End
         Begin VB.Frame Frame15 
            Caption         =   "Receipt"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   90
            TabIndex        =   112
            Top             =   270
            Width           =   5100
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
               Height          =   285
               Index           =   26
               Left            =   1800
               MaxLength       =   25
               TabIndex        =   27
               Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAA"
               Top             =   800
               Width           =   3135
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
               Height          =   285
               Index           =   25
               Left            =   1800
               MaxLength       =   25
               TabIndex        =   26
               Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAA"
               Top             =   360
               Width           =   3135
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Person"
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
               Left            =   360
               TabIndex        =   114
               Top             =   795
               Width           =   585
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Position"
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
               Left            =   360
               TabIndex        =   113
               Top             =   360
               Width           =   660
            End
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "Worksheet"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4245
         Left            =   -74910
         TabIndex        =   101
         Top             =   480
         Width           =   5355
         Begin VB.Frame Frame13 
            Caption         =   "PPC"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   120
            TabIndex        =   108
            Top             =   2820
            Width           =   5100
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
               Height          =   285
               Index           =   23
               Left            =   1800
               MaxLength       =   25
               TabIndex        =   24
               Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAA"
               Top             =   360
               Width           =   3135
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
               Height          =   285
               Index           =   24
               Left            =   1800
               MaxLength       =   25
               TabIndex        =   25
               Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAA"
               Top             =   800
               Width           =   3135
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Position"
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
               Left            =   360
               TabIndex        =   110
               Top             =   360
               Width           =   660
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Person"
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
               Left            =   360
               TabIndex        =   109
               Top             =   795
               Width           =   585
            End
         End
         Begin VB.Frame Frame14 
            Caption         =   "Quality Control"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   120
            TabIndex        =   105
            Top             =   1560
            Width           =   5100
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
               Height          =   285
               Index           =   21
               Left            =   1800
               MaxLength       =   25
               TabIndex        =   22
               Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAA"
               Top             =   360
               Width           =   3135
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
               Height          =   285
               Index           =   22
               Left            =   1800
               MaxLength       =   25
               TabIndex        =   23
               Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAA"
               Top             =   800
               Width           =   3135
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Position"
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
               Left            =   360
               TabIndex        =   107
               Top             =   360
               Width           =   660
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Person"
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
               Left            =   360
               TabIndex        =   106
               Top             =   795
               Width           =   585
            End
         End
         Begin VB.Frame Frame12 
            Caption         =   "Production"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   120
            TabIndex        =   102
            Top             =   270
            Width           =   5100
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
               Height          =   285
               Index           =   20
               Left            =   1800
               MaxLength       =   25
               TabIndex        =   21
               Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAA"
               Top             =   800
               Width           =   3135
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
               Height          =   285
               Index           =   19
               Left            =   1800
               MaxLength       =   25
               TabIndex        =   20
               Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAA"
               Top             =   360
               Width           =   3135
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Person"
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
               Left            =   360
               TabIndex        =   104
               Top             =   795
               Width           =   585
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Position"
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
               Left            =   360
               TabIndex        =   103
               Top             =   360
               Width           =   660
            End
         End
      End
      Begin VB.Frame Frame11 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   -74760
         TabIndex        =   91
         Top             =   630
         Width           =   10335
         Begin VB.TextBox txtpostal 
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
            Left            =   7800
            MaxLength       =   10
            TabIndex        =   35
            Text            =   "AAAAAAAAAA"
            Top             =   1102
            Width           =   1335
         End
         Begin VB.TextBox txtaddress2 
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
            Left            =   1560
            MaxLength       =   33
            TabIndex        =   33
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
            Top             =   1520
            Width           =   4335
         End
         Begin VB.TextBox txtaddress1 
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
            Left            =   1560
            MaxLength       =   33
            TabIndex        =   32
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
            Top             =   1102
            Width           =   4335
         End
         Begin VB.TextBox txtbankname 
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
            Left            =   1560
            MaxLength       =   33
            TabIndex        =   31
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
            Top             =   686
            Width           =   4335
         End
         Begin VB.TextBox txtaccount 
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
            Left            =   7800
            MaxLength       =   33
            TabIndex        =   36
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAA"
            Top             =   1520
            Width           =   2295
         End
         Begin VB.TextBox txtcity 
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
            Left            =   7800
            MaxLength       =   25
            TabIndex        =   34
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAA"
            Top             =   686
            Width           =   2295
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Postal Code :"
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
            Left            =   6480
            TabIndex        =   99
            Top             =   1146
            Width           =   1215
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Address 2    :"
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
            Left            =   240
            TabIndex        =   98
            Top             =   1550
            Width           =   1215
         End
         Begin MSForms.ComboBox cbocurr 
            Height          =   315
            Left            =   1560
            TabIndex        =   30
            Top             =   240
            Width           =   825
            VariousPropertyBits=   746604571
            MaxLength       =   2
            DisplayStyle    =   7
            Size            =   "1455;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "Verdana"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "Currency     :"
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
            Left            =   240
            TabIndex        =   97
            Top             =   280
            Width           =   1215
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Account No  :"
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
            Left            =   6480
            TabIndex        =   96
            Top             =   1550
            Width           =   1215
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Address 1    :"
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
            Left            =   240
            TabIndex        =   95
            Top             =   1146
            Width           =   1215
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Name  :"
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
            Left            =   240
            TabIndex        =   94
            Top             =   713
            Width           =   1215
         End
         Begin VB.Label lbldesc 
            Alignment       =   2  'Center
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
            Left            =   2520
            TabIndex        =   93
            Top             =   270
            Width           =   735
         End
         Begin VB.Line Line1 
            Index           =   1
            X1              =   2520
            X2              =   3240
            Y1              =   510
            Y2              =   510
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "City            :"
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
            Left            =   6480
            TabIndex        =   92
            Top             =   713
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         Height          =   3975
         Left            =   -72457
         TabIndex        =   84
         Top             =   630
         Width           =   6195
         Begin VB.TextBox Text1 
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
            Index           =   35
            Left            =   1800
            MaxLength       =   100
            TabIndex        =   41
            Text            =   "AAAAAAAAAAAAAAAAAAAA"
            Top             =   2400
            Width           =   2475
         End
         Begin VB.TextBox Text1 
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
            Height          =   855
            Index           =   34
            Left            =   1800
            MaxLength       =   200
            MultiLine       =   -1  'True
            TabIndex        =   40
            Text            =   "FrmCompanyProfile.frx":0F06
            Top             =   1346
            Width           =   3975
         End
         Begin VB.TextBox Text1 
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
            Index           =   33
            Left            =   1800
            MaxLength       =   100
            TabIndex        =   39
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
            Top             =   838
            Width           =   3975
         End
         Begin VB.TextBox Text1 
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
            Index           =   32
            Left            =   1800
            MaxLength       =   100
            TabIndex        =   38
            Text            =   "AAAAAAAAAAAAAAAAAAAA"
            Top             =   330
            Width           =   2475
         End
         Begin VB.TextBox Text1 
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
            Index           =   31
            Left            =   1800
            MaxLength       =   25
            TabIndex        =   42
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAA"
            Top             =   2932
            Width           =   3135
         End
         Begin MSComCtl2.DTPicker tglpengukuhan 
            Height          =   315
            Left            =   1800
            TabIndex        =   43
            Top             =   3440
            Width           =   1575
            _ExtentX        =   2778
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
            Format          =   136970243
            CurrentDate     =   37810
         End
         Begin VB.Label Label15 
            Caption         =   "Tgl Pengukuhan :"
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
            Left            =   165
            TabIndex        =   90
            Top             =   3460
            Width           =   1575
         End
         Begin VB.Label Label16 
            Caption         =   "Faktur Pajak No :"
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
            Left            =   165
            TabIndex        =   89
            Top             =   2918
            Width           =   1575
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NPWP City         :"
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
            Left            =   165
            TabIndex        =   88
            Top             =   2436
            Width           =   1575
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NPWP Address   :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   19
            Left            =   165
            TabIndex        =   87
            Top             =   1294
            Width           =   1575
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NPWP Name      :"
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
            Left            =   165
            TabIndex        =   86
            Top             =   812
            Width           =   1575
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NPWP No           :"
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
            Left            =   165
            TabIndex        =   85
            Top             =   330
            Width           =   1575
         End
      End
      Begin VB.Frame Frame8 
         Height          =   3975
         Left            =   -74760
         TabIndex        =   82
         Top             =   660
         Width           =   10305
         Begin VB.TextBox Text1 
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
            Index           =   30
            Left            =   5190
            MaxLength       =   4
            TabIndex        =   44
            Text            =   "AAAA"
            Top             =   1530
            Width           =   855
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sales  :"
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
            Left            =   4320
            TabIndex        =   83
            Top             =   1575
            Width           =   660
         End
      End
      Begin VB.Frame Frame10 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   -74760
         TabIndex        =   79
         Top             =   600
         Width           =   5100
         Begin VB.TextBox Text1 
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
            Index           =   29
            Left            =   1800
            MaxLength       =   25
            TabIndex        =   9
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAA"
            Top             =   480
            Width           =   3135
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "President Director"
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
            TabIndex        =   80
            Top             =   480
            Width           =   1635
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Prepared By (Export)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   -74760
         TabIndex        =   76
         Top             =   3360
         Width           =   5100
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
            Height          =   285
            Index           =   12
            Left            =   1800
            MaxLength       =   25
            TabIndex        =   13
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAA"
            Top             =   800
            Width           =   3135
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
            Height          =   285
            Index           =   11
            Left            =   1800
            MaxLength       =   25
            TabIndex        =   12
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAA"
            Top             =   360
            Width           =   3135
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Person"
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
            Left            =   360
            TabIndex        =   78
            Top             =   795
            Width           =   585
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Position"
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
            Left            =   360
            TabIndex        =   77
            Top             =   360
            Width           =   660
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Purchasing Leader"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   -69480
         TabIndex        =   72
         Top             =   3360
         Width           =   5055
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
            Height          =   285
            Index           =   18
            Left            =   1680
            MaxLength       =   25
            TabIndex        =   19
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAA"
            Top             =   800
            Width           =   3135
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
            Height          =   285
            Index           =   17
            Left            =   1680
            MaxLength       =   25
            TabIndex        =   18
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAA"
            Top             =   360
            Width           =   3135
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Person"
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
            Left            =   360
            TabIndex        =   74
            Top             =   795
            Width           =   1275
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Position"
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
            Left            =   360
            TabIndex        =   73
            Top             =   360
            Width           =   660
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Approved By"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   -69480
         TabIndex        =   69
         Top             =   1980
         Width           =   5055
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
            Height          =   285
            Index           =   16
            Left            =   1680
            MaxLength       =   25
            TabIndex        =   17
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAA"
            Top             =   800
            Width           =   3135
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
            Height          =   285
            Index           =   15
            Left            =   1680
            MaxLength       =   25
            TabIndex        =   16
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAA"
            Top             =   360
            Width           =   3135
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Person"
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
            Left            =   360
            TabIndex        =   71
            Top             =   795
            Width           =   585
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Position"
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
            Left            =   360
            TabIndex        =   70
            Top             =   360
            Width           =   660
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Checked By"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   -69480
         TabIndex        =   49
         Top             =   600
         Width           =   5055
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
            Height          =   285
            Index           =   14
            Left            =   1680
            MaxLength       =   25
            TabIndex        =   15
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAA"
            Top             =   800
            Width           =   3135
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
            Height          =   285
            Index           =   13
            Left            =   1680
            MaxLength       =   25
            TabIndex        =   14
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAA"
            Top             =   360
            Width           =   3135
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Person"
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
            Left            =   360
            TabIndex        =   68
            Top             =   795
            Width           =   585
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Position"
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
            Left            =   360
            TabIndex        =   67
            Top             =   360
            Width           =   660
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Prepared By (Local)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   -74760
         TabIndex        =   64
         Top             =   1980
         Width           =   5100
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
            Height          =   285
            Index           =   10
            Left            =   1800
            MaxLength       =   25
            TabIndex        =   11
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAA"
            Top             =   800
            Width           =   3135
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
            Height          =   285
            Index           =   9
            Left            =   1800
            MaxLength       =   25
            TabIndex        =   10
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAA"
            Top             =   360
            Width           =   3135
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Person"
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
            Left            =   360
            TabIndex        =   66
            Top             =   795
            Width           =   585
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Position"
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
            Left            =   360
            TabIndex        =   65
            Top             =   360
            Width           =   660
         End
      End
      Begin VB.Frame Frame2 
         Height          =   4335
         Left            =   240
         TabIndex        =   55
         Top             =   480
         Width           =   10305
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
            Height          =   285
            Index           =   36
            Left            =   3840
            MaxLength       =   50
            TabIndex        =   127
            Text            =   "AAAAAAAAAAAAAAA"
            Top             =   3960
            Width           =   3975
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
            Height          =   285
            Index           =   0
            Left            =   3840
            MaxLength       =   50
            TabIndex        =   0
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
            Top             =   310
            Width           =   4575
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
            Height          =   285
            Index           =   8
            Left            =   3840
            MaxLength       =   15
            TabIndex        =   8
            Text            =   "AAAAAAAAAAAAAAA"
            Top             =   3480
            Width           =   1815
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
            Height          =   285
            Index           =   7
            Left            =   5880
            MaxLength       =   15
            TabIndex        =   7
            Text            =   "AAAAAAAAAAAAAAA"
            Top             =   3022
            Width           =   1815
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
            Height          =   285
            Index           =   6
            Left            =   3840
            MaxLength       =   15
            TabIndex        =   6
            Text            =   "AAAAAAAAAAAAAAA"
            Top             =   3022
            Width           =   1815
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
            Height          =   285
            Index           =   5
            Left            =   3840
            MaxLength       =   10
            TabIndex        =   4
            Text            =   "AAAAAAAAAA"
            Top             =   2118
            Width           =   1575
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
            Height          =   285
            Index           =   4
            Left            =   3840
            MaxLength       =   100
            TabIndex        =   3
            Text            =   "AAAAAAAAAAAAAAAAAAAA"
            Top             =   1666
            Width           =   2535
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
            Height          =   285
            Index           =   3
            Left            =   3840
            MaxLength       =   100
            TabIndex        =   5
            Text            =   "AAAAAAAAAAAAAAAAAAAA"
            Top             =   2570
            Width           =   2535
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
            Height          =   285
            Index           =   2
            Left            =   3840
            MaxLength       =   100
            TabIndex        =   2
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
            Top             =   1200
            Width           =   3615
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
            Height          =   285
            Index           =   1
            Left            =   3840
            MaxLength       =   100
            TabIndex        =   1
            Text            =   "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
            Top             =   762
            Width           =   3615
         End
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            Caption         =   "Nomor Izin            :"
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
            Left            =   1920
            TabIndex        =   128
            Top             =   3960
            Width           =   1815
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Company Name    :"
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
            Left            =   1958
            TabIndex        =   75
            Top             =   330
            Width           =   1815
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "/"
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
            Left            =   5640
            TabIndex        =   63
            Top             =   3030
            Width           =   255
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Fax                       :"
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
            Left            =   1920
            TabIndex        =   62
            Top             =   3480
            Width           =   1935
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Phone 1 / Phone 2 :"
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
            Left            =   1958
            TabIndex        =   61
            Top             =   3030
            Width           =   1815
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Postal Code          :"
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
            Left            =   1958
            TabIndex        =   60
            Top             =   2130
            Width           =   1815
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "City                     :"
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
            Left            =   1958
            TabIndex        =   59
            Top             =   1680
            Width           =   1815
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Province               :"
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
            Left            =   1958
            TabIndex        =   58
            Top             =   2580
            Width           =   1815
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Address 2             :"
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
            Left            =   1958
            TabIndex        =   57
            Top             =   1230
            Width           =   1815
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Address 1             : "
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
            Left            =   1958
            TabIndex        =   56
            Top             =   780
            Width           =   1815
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid grid 
         Height          =   1695
         Left            =   -74760
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   3000
         Width           =   10335
         _cx             =   18230
         _cy             =   2999
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
         RowHeightMax    =   0
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
      Begin VB.Shape Shape2 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   -74760
         Top             =   2670
         Width           =   10335
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Company Bank Account"
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
         Left            =   -70800
         TabIndex        =   100
         Top             =   2670
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   525
      Left            =   392
      TabIndex        =   53
      Top             =   6120
      Width           =   10965
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
         Height          =   210
         Left            =   120
         TabIndex        =   54
         Top             =   195
         Width           =   10740
      End
   End
   Begin VB.CommandButton Command1 
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
      Left            =   392
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   6765
      Width           =   1155
   End
   Begin VB.CommandButton Command2 
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
      Left            =   10142
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   6750
      Width           =   1215
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   9525
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   345
      Width           =   1845
      _extentx        =   3254
      _extenty        =   714
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Company Profile"
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
      Left            =   405
      TabIndex        =   52
      Top             =   360
      Width           =   10965
   End
End
Attribute VB_Name = "FrmCompanyProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS As New ADODB.Recordset
Dim rsGrid As New ADODB.Recordset
Dim sql, acc As String
Dim first, updateGrid As Boolean

Dim bteColSelect As Byte
Dim bteColCurr As Byte
Dim bteColBankName As Byte
Dim bteColAddress1 As Byte
Dim bteColAddress2 As Byte
Dim bteColCity As Byte
Dim bteColPostCode As Byte
Dim bteColAccountNo As Byte
Dim bteColSeqNo As Byte

Sub headerGrid()
  
    bteColSelect = 0
    bteColCurr = 1
    bteColBankName = 2
    bteColAddress1 = 3
    bteColAddress2 = 4
    bteColCity = 5
    bteColPostCode = 6
    bteColAccountNo = 7
    bteColSeqNo = 8
  
    With grid
        .ColS = 9
        .Rows = 1
        
        .TextMatrix(0, bteColSelect) = ""
        .TextMatrix(0, bteColCurr) = "Currency"
        .TextMatrix(0, bteColBankName) = "Bank Name"
        .TextMatrix(0, bteColAddress1) = "Address 1"
        .TextMatrix(0, bteColAddress2) = "Address 2"
        .TextMatrix(0, bteColCity) = "City"
        .TextMatrix(0, bteColPostCode) = "Postal Code"
        .TextMatrix(0, bteColAccountNo) = "Account No"
        .TextMatrix(0, bteColSeqNo) = "SeqNo"
        
        .ColWidth(bteColSelect) = 300
        .ColWidth(bteColCurr) = 975
        .ColWidth(bteColBankName) = 3000
        .ColWidth(bteColAddress1) = 3000
        .ColWidth(bteColAddress2) = 3000
        .ColWidth(bteColCity) = 2000
        .ColWidth(bteColPostCode) = 1155
        .ColWidth(bteColAccountNo) = 2000
        
        .ColHidden(bteColSeqNo) = 0
        
        .ColAlignment(bteColSelect) = flexAlignCenterCenter
        .ColAlignment(bteColCurr) = flexAlignLeftCenter
        .ColAlignment(bteColBankName) = flexAlignLeftCenter
        .ColAlignment(bteColAddress1) = flexAlignLeftCenter
        .ColAlignment(bteColAddress2) = flexAlignLeftCenter
        .ColAlignment(bteColCity) = flexAlignLeftCenter
        .ColAlignment(bteColPostCode) = flexAlignLeftCenter
        .ColAlignment(bteColAccountNo) = flexAlignLeftCenter
        
        .Cell(flexcpAlignment, 0, 1, 0, 7) = flexAlignCenterCenter
        .EditMaxLength = 1
        
        .MergeCells = flexMergeFree
        For i = 2 To 6
            .MergeCol(i) = True
        Next i
    End With

End Sub

Sub kosongBwh()
    cbocurr.ListIndex = -1
    lbldesc.Caption = ""
    txtbankname.Text = ""
    txtaddress1.Text = ""
    txtaddress2.Text = ""
    txtcity.Text = ""
    txtpostal.Text = ""
    TxtAccount.Text = ""
    TxtSeqNo.Text = ""
End Sub

Sub kosonggrid()
  With grid
    For i = 1 To .Rows - 1
      .TextMatrix(i, bteColSelect) = ""
    Next i
  End With
  kosongBwh
  LblErrMsg.Caption = ""
End Sub

Sub Browse()
    sql = "select * from Company_Profile"
    If RS.State <> adStateClosed Then RS.Close
    RS.Open sql, Db, adOpenKeyset, adLockOptimistic
    
    If RS.BOF And RS.EOF Then
      first = True
    Else
      For i = 0 To 35
        Text1(i).Text = IIf(IsNull(Trim(RS(i + 1))), "", Trim(RS(i + 1)))
      Next i
      tglpengukuhan.Value = IIf(IsNull(RS(40)), Format(Now, "dd MMM yyyy"), Format(RS(40), "dd MMM yyyy"))
      txtBC(0).Text = Trim(RS(43) & "")
      txtBC(1).Text = Trim(RS(44) & "")
      txtBC(2).Text = Trim(RS(45) & "")
      txtBC(3).Text = Trim(RS(46) & "")
      Text1(36).Text = Trim(RS(52) & "")
      first = False
    End If

End Sub

Sub BrowseGrid()
  Dim i  As Integer

  i = 1

    sql = "select * from company_bank where company_code='00000' order by bank_name, address1, address2, city, postal_code, currency_code, account_no"
    If rsGrid.State <> adStateClosed Then rsGrid.Close
    rsGrid.Open sql, Db, adOpenKeyset, adLockOptimistic
  
  If Not (rsGrid.BOF And rsGrid.EOF) Then
  Do While Not rsGrid.EOF
    With grid

      .Rows = .Rows + 1

      .TextMatrix(i, bteColSelect) = ""
      .Cell(flexcpBackColor, i, bteColSelect) = &HFFFFFF
      .TextMatrix(i, bteColCurr) = Trim(rsGrid("Currency_Code")) & " - " & uf_GetCurrencyDescription(rsGrid("currency_code"))
      .TextMatrix(i, bteColBankName) = Trim(rsGrid("Bank_Name"))
      .TextMatrix(i, bteColAddress1) = Trim(rsGrid("Address1"))
      .TextMatrix(i, bteColAddress2) = Trim(rsGrid("Address2"))
      .TextMatrix(i, bteColCity) = Trim(rsGrid("City"))
      .TextMatrix(i, bteColPostCode) = Trim(rsGrid("Postal_code"))
      .TextMatrix(i, bteColAccountNo) = Trim(rsGrid("Account_no"))
      .TextMatrix(i, bteColSeqNo) = rsGrid("seq_no")
      rsGrid.MoveNext
      i = i + 1
    End With
  Loop
  End If

End Sub

Private Sub cboCurr_GotFocus()
    SSTab1.Tab = 3
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
    Me.CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"
    
    For i = 0 To 31
      Text1(i).Text = ""
    Next i
    tglpengukuhan = Format(Now, "dd MMM yyyy")
    txtBC(0).Text = ""
    txtBC(1).Text = ""
    txtBC(2).Text = ""
    txtBC(3).Text = ""
    
  Call up_FillCombo(cbocurr, "curr_cls")
    
    updateGrid = False
    kosonggrid
    headerGrid
    
    Browse
    BrowseGrid
    
    SSTab1.Tab = 0
End Sub

Private Sub cbocurr_Click()
    If cbocurr.ListIndex <> -1 Then
        lbldesc.Caption = cbocurr.Column(1)
    End If
End Sub

Private Sub cbocurr_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Command1_Click()
  Unload Me
  frmMainMenu.Show
End Sub

Private Sub command2_Click(Index As Integer)
Dim hapus As Boolean, ugrid As Boolean
Dim tanya

hapus = False
ugrid = False
  Select Case Index
    Case 0:
        If Text1(0).Text = "" Then
          LblErrMsg = DisplayMsg(1039)
          SSTab1.Tab = 0
          Text1(0).SetFocus
          Exit Sub
        Else
            
            If (first) Then
              RS.AddNew
              RS(0) = "00000"
            End If
            For i = 0 To 35
                RS(i + 1) = Text1(i).Text
            Next i
            RS(40) = Format(tglpengukuhan.Value, "yyyy-mm-dd")
            
            RS(43) = txtBC(0).Text
            RS(44) = txtBC(1).Text
            RS(45) = txtBC(2).Text
            RS(46) = txtBC(3).Text
            RS(52) = Text1(36).Text
            
            RS(47) = Now
            RS(48) = userLogin
            RS.update
            
            With grid
                For i = 1 To .Rows - 1
                  If .TextMatrix(i, bteColSelect) = "D" Then
                    If IsEmpty(tanya) Then tanya = MsgBox("Do You Really Want To Delete This Data ?", vbQuestion & vbYesNo, "Confirmation")
                    If tanya = vbYes Then
                        sql = "delete from Company_Bank where seq_no='" & .TextMatrix(i, bteColSeqNo) & "' "
                        Db.Execute sql
                        hapus = True
                    Else
                        Exit For
                    End If
                  End If
                Next i

                If (hapus) Then headerGrid: BrowseGrid: LblErrMsg = DisplayMsg(1201): Exit Sub
            End With
            
            If cbocurr.Text <> "" Or txtbankname.Text <> "" Or txtaddress1.Text <> "" Or txtaddress2.Text <> "" Or txtcity.Text <> "" Or txtpostal.Text <> "" Or TxtAccount.Text <> "" Then
                ugrid = True
                If cbocurr.Text = "" Then
                    LblErrMsg = DisplayMsg(1028)
                    cbocurr.SetFocus
                    Exit Sub
                ElseIf txtbankname.Text = "" Then
                    LblErrMsg = DisplayMsg(1055)
                    txtbankname.SetFocus
                    Exit Sub
                ElseIf txtaddress1.Text = "" Then
                    LblErrMsg = DisplayMsg(1056)
                    txtaddress1.SetFocus
                    Exit Sub
                ElseIf TxtAccount.Text = "" Then
                    LblErrMsg = DisplayMsg(1057)
                    TxtAccount.SetFocus
                    Exit Sub
                End If
                
                If updateGrid = False Then
                        rsGrid.AddNew
                        rsGrid("Company_Code") = RS("Company_Code")
                Else
                    rsGrid.filter = "seq_no='" & TxtSeqNo.Text & "' "
                End If
                
                rsGrid("Currency_code") = cbocurr.Text
                rsGrid("Bank_Name") = txtbankname.Text
                rsGrid("Address1") = txtaddress1.Text
                rsGrid("Address2") = txtaddress2.Text
                rsGrid("City") = txtcity.Text
                rsGrid("Postal_code") = txtpostal.Text
                rsGrid("Account_No") = TxtAccount.Text
                rsGrid("Last_Update") = Now
                rsGrid("Last_User") = userLogin
                rsGrid.update

            rsGrid.Requery
            rsGrid.filter = ""
            kosonggrid
            headerGrid
            BrowseGrid
                
            End If
            If ugrid = False Then
                LblErrMsg = DisplayMsg(IIf((first = True), 1000, 1101))
            Else
                LblErrMsg = DisplayMsg(IIf((updateGrid = False), 1000, 1101))
            End If
            updateGrid = False
            Browse
            
        End If

    Case 1: kosonggrid
            updateGrid = False
    Case 2:
            With grid
                For i = 1 To .Rows - 1
                  If .TextMatrix(i, bteColSelect) = "S" Then
                    cbocurr.ListIndex = -1
                    lbldesc.Caption = ""
                    txtbankname.Text = .TextMatrix(i, bteColBankName)
                    txtaddress1.Text = .TextMatrix(i, bteColAddress1)
                    txtaddress2.Text = .TextMatrix(i, bteColAddress2)
                    txtcity.Text = .TextMatrix(i, bteColCity)
                    txtpostal.Text = .TextMatrix(i, bteColPostCode)
                    TxtAccount.Text = ""
                    updateGrid = False
                  End If
                Next i
            End With
            
  End Select
  
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim TextGrid As String

With grid
    TextGrid = grid.Text

    If TextGrid = "S" Then
       Call kosongColGrid

       cbocurr.Text = Left(Trim(.TextMatrix(Row, bteColCurr)), 2)
       txtbankname.Text = Trim(.TextMatrix(Row, bteColBankName))
       txtaddress1.Text = Trim(.TextMatrix(Row, bteColAddress1))
       txtaddress2.Text = Trim(.TextMatrix(Row, bteColAddress2))
       txtcity.Text = Trim(.TextMatrix(Row, bteColCity))
       txtpostal.Text = Trim(.TextMatrix(Row, bteColPostCode))
       TxtAccount.Text = Trim(.TextMatrix(Row, bteColAccountNo))
       TxtSeqNo.Text = .TextMatrix(Row, bteColSeqNo)
       updateGrid = True

    Else
       Call kosongColGrid("S")
        kosongBwh
    End If

    .TextMatrix(Row, Col) = TextGrid
End With

End Sub

Private Sub kosongColGrid(Optional Kolom As String)
    Dim i As Integer

    With grid
        .Col = 0

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
  If grid.Col <> bteColSelect Then Cancel = True
End Sub

Private Sub Grid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
  If grid.Col = bteColSelect Then
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii <> Asc("S") And KeyAscii <> Asc("D") And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then
      KeyAscii = 0
    End If
    If KeyAscii = Asc(".") Then KeyAscii = 0
  End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Select Case SSTab1.Tab
    Case 0: Text1(0).SetFocus
    Case 1: Text1(29).SetFocus
    Case 2: Text1(19).SetFocus
    Case 3: cbocurr.SetFocus
    Case 4: Text1(32).SetFocus
    Case 5: Text1(30).SetFocus
    End Select
    command2(1).Visible = SSTab1.Tab = 3
    command2(2).Visible = SSTab1.Tab = 3
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    Select Case Index
    Case 8: SSTab1.Tab = 0
    Case 29: SSTab1.Tab = 1
    Case 19: SSTab1.Tab = 2
    Case 32: SSTab1.Tab = 4
    Case 30: SSTab1.Tab = 5
    End Select
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub tglpengukuhan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub txtaddress1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtaddress2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtbankname_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtcity_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtpostal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtaccount_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
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
    RS.Close
    If rsGrid.State <> adStateClosed Then rsGrid.Close
End Sub
