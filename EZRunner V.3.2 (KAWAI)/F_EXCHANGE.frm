VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form F_EXCHANGE 
   BackColor       =   &H00FDDFE3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Book Exchange Rate"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14010
   Icon            =   "F_EXCHANGE.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleMode       =   0  'User
   ScaleWidth      =   14021.79
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox BER_F 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   12
      Left            =   12480
      MaxLength       =   12
      TabIndex        =   73
      Top             =   4560
      Width           =   915
   End
   Begin VB.TextBox BER_F 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   11
      Left            =   11535
      MaxLength       =   12
      TabIndex        =   72
      Top             =   4560
      Width           =   915
   End
   Begin VB.TextBox BER_F 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   10
      Left            =   10575
      MaxLength       =   12
      TabIndex        =   71
      Top             =   4560
      Width           =   915
   End
   Begin VB.TextBox BER_F 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   9
      Left            =   9615
      MaxLength       =   12
      TabIndex        =   70
      Top             =   4560
      Width           =   915
   End
   Begin VB.TextBox BER_F 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   8
      Left            =   8670
      MaxLength       =   12
      TabIndex        =   69
      Top             =   4560
      Width           =   915
   End
   Begin VB.TextBox BER_F 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   7
      Left            =   7710
      MaxLength       =   12
      TabIndex        =   68
      Top             =   4560
      Width           =   915
   End
   Begin VB.TextBox BER_F 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   6
      Left            =   6750
      MaxLength       =   12
      TabIndex        =   67
      Top             =   4560
      Width           =   915
   End
   Begin VB.TextBox BER_F 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   5
      Left            =   5790
      MaxLength       =   12
      TabIndex        =   66
      Top             =   4560
      Width           =   915
   End
   Begin VB.TextBox BER_F 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   4
      Left            =   4830
      MaxLength       =   12
      TabIndex        =   65
      Top             =   4560
      Width           =   915
   End
   Begin VB.TextBox BER_F 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   3
      Left            =   3870
      MaxLength       =   12
      TabIndex        =   64
      Top             =   4560
      Width           =   915
   End
   Begin VB.TextBox BER_F 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   2
      Left            =   2910
      MaxLength       =   12
      TabIndex        =   63
      Top             =   4560
      Width           =   915
   End
   Begin VB.TextBox BER_F 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   1
      Left            =   1920
      MaxLength       =   12
      TabIndex        =   62
      Top             =   4560
      Width           =   915
   End
   Begin VB.TextBox code 
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
      Index           =   6
      Left            =   1215
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   111
      TabStop         =   0   'False
      Top             =   4560
      Width           =   375
   End
   Begin VB.TextBox BER_E 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   12
      Left            =   12495
      MaxLength       =   12
      TabIndex        =   61
      Top             =   4080
      Width           =   915
   End
   Begin VB.TextBox BER_E 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   11
      Left            =   11535
      MaxLength       =   12
      TabIndex        =   60
      Top             =   4080
      Width           =   915
   End
   Begin VB.TextBox BER_E 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   10
      Left            =   10575
      MaxLength       =   12
      TabIndex        =   59
      Top             =   4080
      Width           =   915
   End
   Begin VB.TextBox BER_E 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   9
      Left            =   9615
      MaxLength       =   12
      TabIndex        =   58
      Top             =   4080
      Width           =   915
   End
   Begin VB.TextBox BER_E 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   8
      Left            =   8670
      MaxLength       =   12
      TabIndex        =   57
      Top             =   4080
      Width           =   915
   End
   Begin VB.TextBox BER_E 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   7
      Left            =   7710
      MaxLength       =   12
      TabIndex        =   56
      Top             =   4080
      Width           =   915
   End
   Begin VB.TextBox BER_E 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   6
      Left            =   6750
      MaxLength       =   12
      TabIndex        =   55
      Top             =   4080
      Width           =   915
   End
   Begin VB.TextBox BER_E 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   5
      Left            =   5790
      MaxLength       =   12
      TabIndex        =   54
      Top             =   4080
      Width           =   915
   End
   Begin VB.TextBox BER_E 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   4
      Left            =   4830
      MaxLength       =   12
      TabIndex        =   53
      Top             =   4080
      Width           =   915
   End
   Begin VB.TextBox BER_E 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   3
      Left            =   3870
      MaxLength       =   12
      TabIndex        =   52
      Top             =   4080
      Width           =   915
   End
   Begin VB.TextBox BER_E 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   2
      Left            =   2910
      MaxLength       =   12
      TabIndex        =   51
      Top             =   4080
      Width           =   915
   End
   Begin VB.TextBox BER_E 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   1
      Left            =   1920
      MaxLength       =   12
      TabIndex        =   50
      Top             =   4080
      Width           =   915
   End
   Begin VB.TextBox code 
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
      Index           =   5
      Left            =   1215
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   108
      TabStop         =   0   'False
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox BER_A 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   1
      Left            =   1935
      MaxLength       =   12
      TabIndex        =   2
      Top             =   2160
      Width           =   915
   End
   Begin VB.TextBox BER_A 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   5
      Left            =   5775
      MaxLength       =   12
      TabIndex        =   6
      Top             =   2160
      Width           =   915
   End
   Begin VB.CommandButton Cmd_Clear 
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
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   5880
      Width           =   1185
   End
   Begin VB.CommandButton Cmd_Save 
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
      Left            =   12390
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   5880
      Width           =   1185
   End
   Begin VB.CommandButton CommandButton1 
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
      Left            =   375
      Style           =   1  'Graphical
      TabIndex        =   76
      Top             =   5880
      Width           =   1185
   End
   Begin VB.TextBox TxtYear 
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
      Left            =   960
      MaxLength       =   4
      TabIndex        =   0
      Top             =   1020
      Width           =   675
   End
   Begin VB.TextBox code 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Height          =   285
      Index           =   4
      Left            =   1215
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   80
      TabStop         =   0   'False
      Top             =   3600
      Width           =   375
   End
   Begin VB.TextBox code 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Height          =   285
      Index           =   3
      Left            =   1215
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   79
      TabStop         =   0   'False
      Top             =   3120
      Width           =   375
   End
   Begin VB.TextBox code 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Height          =   285
      Index           =   2
      Left            =   1215
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox code 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Height          =   285
      Index           =   1
      Left            =   1215
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   77
      TabStop         =   0   'False
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox BER_D 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   12
      Left            =   12495
      MaxLength       =   12
      TabIndex        =   49
      Top             =   3600
      Width           =   915
   End
   Begin VB.TextBox BER_D 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   11
      Left            =   11535
      MaxLength       =   12
      TabIndex        =   48
      Top             =   3600
      Width           =   915
   End
   Begin VB.TextBox BER_D 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   10
      Left            =   10575
      MaxLength       =   12
      TabIndex        =   47
      Top             =   3600
      Width           =   915
   End
   Begin VB.TextBox BER_D 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   9
      Left            =   9615
      MaxLength       =   12
      TabIndex        =   46
      Top             =   3600
      Width           =   915
   End
   Begin VB.TextBox BER_D 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   8
      Left            =   8655
      MaxLength       =   12
      TabIndex        =   45
      Top             =   3600
      Width           =   915
   End
   Begin VB.TextBox BER_D 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   7
      Left            =   7695
      MaxLength       =   12
      TabIndex        =   44
      Top             =   3600
      Width           =   915
   End
   Begin VB.TextBox BER_D 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   6
      Left            =   6735
      MaxLength       =   12
      TabIndex        =   43
      Top             =   3600
      Width           =   915
   End
   Begin VB.TextBox BER_D 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   5
      Left            =   5775
      MaxLength       =   12
      TabIndex        =   42
      Top             =   3600
      Width           =   915
   End
   Begin VB.TextBox BER_D 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   4
      Left            =   4815
      MaxLength       =   12
      TabIndex        =   41
      Top             =   3600
      Width           =   915
   End
   Begin VB.TextBox BER_D 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   3
      Left            =   3855
      MaxLength       =   12
      TabIndex        =   40
      Top             =   3600
      Width           =   915
   End
   Begin VB.TextBox BER_D 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   2
      Left            =   2895
      MaxLength       =   12
      TabIndex        =   39
      Top             =   3600
      Width           =   915
   End
   Begin VB.TextBox BER_D 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   1
      Left            =   1935
      MaxLength       =   12
      TabIndex        =   38
      Top             =   3600
      Width           =   915
   End
   Begin VB.TextBox BER_C 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   12
      Left            =   12495
      MaxLength       =   12
      TabIndex        =   37
      Top             =   3120
      Width           =   915
   End
   Begin VB.TextBox BER_C 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   11
      Left            =   11535
      MaxLength       =   12
      TabIndex        =   36
      Top             =   3120
      Width           =   915
   End
   Begin VB.TextBox BER_C 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   10
      Left            =   10575
      MaxLength       =   12
      TabIndex        =   35
      Top             =   3120
      Width           =   915
   End
   Begin VB.TextBox BER_C 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   9
      Left            =   9615
      MaxLength       =   12
      TabIndex        =   34
      Top             =   3120
      Width           =   915
   End
   Begin VB.TextBox BER_C 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   8
      Left            =   8655
      MaxLength       =   12
      TabIndex        =   33
      Top             =   3120
      Width           =   915
   End
   Begin VB.TextBox BER_C 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   7
      Left            =   7695
      MaxLength       =   12
      TabIndex        =   32
      Top             =   3120
      Width           =   915
   End
   Begin VB.TextBox BER_C 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   6
      Left            =   6735
      MaxLength       =   12
      TabIndex        =   31
      Top             =   3120
      Width           =   915
   End
   Begin VB.TextBox BER_C 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   5
      Left            =   5775
      MaxLength       =   12
      TabIndex        =   30
      Top             =   3120
      Width           =   915
   End
   Begin VB.TextBox BER_C 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   4
      Left            =   4815
      MaxLength       =   12
      TabIndex        =   29
      Top             =   3120
      Width           =   915
   End
   Begin VB.TextBox BER_C 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   3
      Left            =   3855
      MaxLength       =   12
      TabIndex        =   28
      Top             =   3120
      Width           =   915
   End
   Begin VB.TextBox BER_C 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   2
      Left            =   2895
      MaxLength       =   12
      TabIndex        =   27
      Top             =   3120
      Width           =   915
   End
   Begin VB.TextBox BER_C 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   1
      Left            =   1935
      MaxLength       =   12
      TabIndex        =   26
      Top             =   3120
      Width           =   915
   End
   Begin VB.TextBox BER_B 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   12
      Left            =   12495
      MaxLength       =   12
      TabIndex        =   25
      Top             =   2640
      Width           =   915
   End
   Begin VB.TextBox BER_B 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   11
      Left            =   11535
      MaxLength       =   12
      TabIndex        =   24
      Top             =   2640
      Width           =   915
   End
   Begin VB.TextBox BER_B 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   10
      Left            =   10575
      MaxLength       =   12
      TabIndex        =   23
      Top             =   2640
      Width           =   915
   End
   Begin VB.TextBox BER_B 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   9
      Left            =   9615
      MaxLength       =   12
      TabIndex        =   22
      Top             =   2640
      Width           =   915
   End
   Begin VB.TextBox BER_B 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   8
      Left            =   8655
      MaxLength       =   12
      TabIndex        =   21
      Top             =   2640
      Width           =   915
   End
   Begin VB.TextBox BER_B 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   7
      Left            =   7695
      MaxLength       =   12
      TabIndex        =   20
      Top             =   2640
      Width           =   915
   End
   Begin VB.TextBox BER_B 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   6
      Left            =   6735
      MaxLength       =   12
      TabIndex        =   19
      Top             =   2640
      Width           =   915
   End
   Begin VB.TextBox BER_B 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   5
      Left            =   5775
      MaxLength       =   12
      TabIndex        =   18
      Top             =   2640
      Width           =   915
   End
   Begin VB.TextBox BER_B 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   4
      Left            =   4815
      MaxLength       =   12
      TabIndex        =   17
      Top             =   2640
      Width           =   915
   End
   Begin VB.TextBox BER_B 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   3
      Left            =   3855
      MaxLength       =   12
      TabIndex        =   16
      Top             =   2640
      Width           =   915
   End
   Begin VB.TextBox BER_B 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   2
      Left            =   2895
      MaxLength       =   12
      TabIndex        =   15
      Top             =   2640
      Width           =   915
   End
   Begin VB.TextBox BER_B 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   1
      Left            =   1935
      MaxLength       =   12
      TabIndex        =   14
      Top             =   2640
      Width           =   915
   End
   Begin VB.TextBox BER_A 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   12
      Left            =   12495
      MaxLength       =   12
      TabIndex        =   13
      Top             =   2160
      Width           =   915
   End
   Begin VB.TextBox BER_A 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   11
      Left            =   11535
      MaxLength       =   12
      TabIndex        =   12
      Top             =   2160
      Width           =   915
   End
   Begin VB.TextBox BER_A 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   10
      Left            =   10575
      MaxLength       =   12
      TabIndex        =   11
      Top             =   2160
      Width           =   915
   End
   Begin VB.TextBox BER_A 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   9
      Left            =   9615
      MaxLength       =   12
      TabIndex        =   10
      Top             =   2160
      Width           =   915
   End
   Begin VB.TextBox BER_A 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   8
      Left            =   8655
      MaxLength       =   12
      TabIndex        =   9
      Top             =   2160
      Width           =   915
   End
   Begin VB.TextBox BER_A 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   7
      Left            =   7695
      MaxLength       =   12
      TabIndex        =   8
      Top             =   2160
      Width           =   915
   End
   Begin VB.TextBox BER_A 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   6
      Left            =   6735
      MaxLength       =   12
      TabIndex        =   7
      Top             =   2160
      Width           =   915
   End
   Begin VB.TextBox BER_A 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   4
      Left            =   4815
      MaxLength       =   12
      TabIndex        =   5
      Top             =   2160
      Width           =   915
   End
   Begin VB.TextBox BER_A 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   3
      Left            =   3855
      MaxLength       =   12
      TabIndex        =   4
      Top             =   2160
      Width           =   915
   End
   Begin VB.TextBox BER_A 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   2
      Left            =   2895
      MaxLength       =   12
      TabIndex        =   3
      Top             =   2160
      Width           =   915
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   375
      TabIndex        =   112
      Top             =   5100
      Width           =   13200
      Begin VB.Label LblPesan 
         Alignment       =   2  'Center
         BackColor       =   &H00FDDFE3&
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
         Height          =   285
         Left            =   60
         TabIndex        =   113
         Top             =   180
         Width           =   13035
      End
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   11730
      TabIndex        =   115
      TabStop         =   0   'False
      Top             =   330
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   714
   End
   Begin VB.Line Line17 
      X1              =   1080.909
      X2              =   1080.909
      Y1              =   1800
      Y2              =   4920
   End
   Begin VB.Label nama 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   6
      Left            =   495
      TabIndex        =   110
      Top             =   4560
      Width           =   495
   End
   Begin VB.Line Line16 
      X1              =   360.303
      X2              =   13571.41
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line6 
      X1              =   360.303
      X2              =   13571.41
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line4 
      X1              =   390.328
      X2              =   13571.41
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line2 
      X1              =   1681.414
      X2              =   1681.414
      Y1              =   1560
      Y2              =   4920
   End
   Begin VB.Label nama 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   5
      Left            =   495
      TabIndex        =   109
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackColor       =   &H00FDDFE3&
      Caption         =   "Book Exchange Rate "
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
      Left            =   375
      TabIndex        =   107
      Top             =   345
      Width           =   13200
   End
   Begin MSForms.ComboBox CboTerm 
      Height          =   315
      Left            =   2760
      TabIndex        =   1
      Top             =   1020
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
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "Dec"
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
      Left            =   12495
      TabIndex        =   104
      Top             =   1680
      Width           =   915
   End
   Begin VB.Label Label30 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "Jan"
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
      Left            =   1935
      TabIndex        =   99
      Top             =   1680
      Width           =   915
   End
   Begin VB.Line Line37 
      X1              =   13571.41
      X2              =   13571.41
      Y1              =   1560
      Y2              =   4920
   End
   Begin VB.Line Line15 
      X1              =   1816.527
      X2              =   13571.41
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line14 
      X1              =   1816.527
      X2              =   13571.41
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line13 
      X1              =   1816.527
      X2              =   13571.41
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line12 
      X1              =   1816.527
      X2              =   13571.41
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line11 
      X1              =   1816.527
      X2              =   13571.41
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line10 
      X1              =   1801.515
      X2              =   1801.515
      Y1              =   1560
      Y2              =   4920
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "Feb"
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
      Left            =   2895
      TabIndex        =   97
      Top             =   1680
      Width           =   915
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "March"
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
      Left            =   3855
      TabIndex        =   96
      Top             =   1680
      Width           =   915
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "April"
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
      Left            =   4800
      TabIndex        =   95
      Top             =   1680
      Width           =   915
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "May"
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
      Left            =   5775
      TabIndex        =   94
      Top             =   1680
      Width           =   915
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "June"
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
      Left            =   6735
      TabIndex        =   93
      Top             =   1680
      Width           =   915
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "July"
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
      Left            =   7695
      TabIndex        =   92
      Top             =   1680
      Width           =   915
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "Aug"
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
      Left            =   8655
      TabIndex        =   91
      Top             =   1680
      Width           =   915
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "Sept"
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
      Left            =   9615
      TabIndex        =   90
      Top             =   1680
      Width           =   915
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "Oct"
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
      Left            =   10575
      TabIndex        =   89
      Top             =   1680
      Width           =   915
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "Nov"
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
      Left            =   11535
      TabIndex        =   88
      Top             =   1680
      Width           =   915
   End
   Begin VB.Line Line31 
      X1              =   1696.426
      X2              =   1816.527
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label nama 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   4
      Left            =   495
      TabIndex        =   87
      Top             =   3600
      Width           =   495
   End
   Begin VB.Line Line21 
      X1              =   375.316
      X2              =   1696.426
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label nama 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   3
      Left            =   495
      TabIndex        =   86
      Top             =   3120
      Width           =   495
   End
   Begin VB.Line Line20 
      X1              =   375.316
      X2              =   1696.426
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label nama 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   2
      Left            =   495
      TabIndex        =   85
      Top             =   2640
      Width           =   495
   End
   Begin VB.Line Line9 
      X1              =   375.316
      X2              =   1696.426
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line8 
      X1              =   375.316
      X2              =   375.316
      Y1              =   1560
      Y2              =   2280
   End
   Begin VB.Line Line7 
      X1              =   375.316
      X2              =   1696.426
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line1 
      X1              =   375.316
      X2              =   1696.426
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line5 
      X1              =   375.316
      X2              =   1696.426
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label nama 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   1
      Left            =   495
      TabIndex        =   84
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FDDFE3&
      Caption         =   "Year"
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
      Left            =   375
      TabIndex        =   83
      Top             =   1050
      Width           =   735
   End
   Begin VB.Label Label22 
      BackColor       =   &H00FDDFE3&
      Caption         =   "Term Cls"
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
      Left            =   1800
      TabIndex        =   82
      Top             =   1050
      Width           =   825
   End
   Begin VB.Line Line3 
      X1              =   375.316
      X2              =   375.316
      Y1              =   2040
      Y2              =   4920
   End
   Begin VB.Line Line30 
      X1              =   375.316
      X2              =   375.316
      Y1              =   1560
      Y2              =   3960
   End
   Begin VB.Label Label19 
      BackColor       =   &H00C0E0FF&
      Height          =   2895
      Left            =   1800
      TabIndex        =   81
      Top             =   2040
      Width           =   11775
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080C0FF&
      Height          =   495
      Left            =   1815
      TabIndex        =   105
      Top             =   1560
      Width           =   11760
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "Name"
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
      Left            =   375
      TabIndex        =   101
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "Currency"
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
      Left            =   375
      TabIndex        =   100
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label kode 
      BackColor       =   &H00FDDFE3&
      Height          =   255
      Left            =   2055
      TabIndex        =   98
      Top             =   1080
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label Label29 
      BackColor       =   &H0080C0FF&
      Height          =   3375
      Left            =   1680
      TabIndex        =   106
      Top             =   1560
      Width           =   120
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "Code"
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
      Left            =   1185
      TabIndex        =   102
      Top             =   1800
      Width           =   465
   End
   Begin VB.Label Label18 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   495
      TabIndex        =   103
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C0E0FF&
      Height          =   2895
      Left            =   375
      TabIndex        =   114
      Top             =   2040
      Width           =   1335
   End
End
Attribute VB_Name = "F_EXCHANGE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TermCls As String, RsCurr As New ADODB.Recordset
Dim Exchange As String

Private Sub ShowCurr()
    
    Dim adoRs As New ADODB.Recordset
    
    adoRs.CursorLocation = adUseClient
    adoRs.Open "select curr_cls, description from curr_cls", Db, adOpenForwardOnly, adLockReadOnly, adCmdText
    While Not adoRs.EOF
        nama(adoRs.AbsolutePosition) = Trim(adoRs.Fields("description"))
        code(adoRs.AbsolutePosition) = Trim(adoRs.Fields("curr_cls"))
        adoRs.MoveNext
    Wend
    adoRs.Close

End Sub

Private Sub BER_A_Change(Index As Integer)
BER_A(Index).DataChanged = True
If InStr(1, BER_A(Index).Text, ",") = 1 Then BER_A(Index).Text = Right(BER_A(Index), Len(BER_A(Index)) - 1)
End Sub

Private Sub BER_A_GotFocus(Index As Integer)
BER_A(Index).SelStart = 0
BER_A(Index).SelLength = Len(BER_A(Index).Text)
End Sub

Private Sub BER_A_KeyPress(Index As Integer, KeyAscii As Integer)
If InStr("0123456789.,", Chr(KeyAscii)) = 0 Then
    If KeyAscii = 13 Then
      SendKeys vbTab
    ElseIf KeyAscii <> vbKeyBack Then
    KeyAscii = 0
    End If
End If
If KeyAscii = 39 Then KeyAscii = 0
If (BER_A(Index).Text & Chr(KeyAscii)) > 9999999.99 And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub BER_A_LostFocus(Index As Integer)
    Dim z As Double
    If BER_A(Index).Text <> "" Then
        z = CDbl(BER_A(Index).Text)
        If z > 9999999.99 Then
            BER_A(Index).Text = Left(z, 7)
        End If
    Else
        BER_A(Index) = Format(0, gs_formatExchangeRate)
    End If
    BER_A(Index).Text = Format(BER_A(Index).Text, gs_formatExchangeRate)
End Sub

Private Sub BER_B_Change(Index As Integer)
BER_B(Index).DataChanged = True
If InStr(1, BER_B(Index).Text, ",") = 1 Then BER_B(Index).Text = Right(BER_B(Index), Len(BER_B(Index)) - 1)
End Sub

Private Sub BER_B_GotFocus(Index As Integer)
BER_B(Index).SelStart = 0
BER_B(Index).SelLength = Len(BER_B(Index).Text)
End Sub

Private Sub BER_B_KeyPress(Index As Integer, KeyAscii As Integer)
If InStr("0123456789.,", Chr(KeyAscii)) = 0 Then
    If KeyAscii = 13 Then
      SendKeys vbTab
    ElseIf KeyAscii <> vbKeyBack Then
    KeyAscii = 0
    End If
End If
If KeyAscii = 39 Then KeyAscii = 0
If (BER_B(Index).Text & Chr(KeyAscii)) > 9999999.99 And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub BER_B_LostFocus(Index As Integer)
    Dim z As Double
    If BER_B(Index).Text <> "" Then
        z = CDbl(BER_B(Index).Text)
        If z > 9999999.99 Then
            BER_B(Index).Text = Left(z, 7)
        End If
    Else
        BER_B(Index) = Format(0, gs_formatExchangeRate)
    End If
    BER_B(Index).Text = Format(BER_B(Index).Text, gs_formatExchangeRate)
End Sub

Private Sub BER_C_Change(Index As Integer)
BER_C(Index).DataChanged = True
If InStr(1, BER_C(Index).Text, ",") = 1 Then BER_C(Index).Text = Right(BER_C(Index), Len(BER_C(Index)) - 1)
End Sub

Private Sub BER_C_GotFocus(Index As Integer)
BER_C(Index).SelStart = 0
BER_C(Index).SelLength = Len(BER_C(Index).Text)
End Sub

Private Sub BER_C_KeyPress(Index As Integer, KeyAscii As Integer)
If InStr("0123456789.,", Chr(KeyAscii)) = 0 Then
    If KeyAscii = 13 Then
      SendKeys vbTab
    ElseIf KeyAscii <> vbKeyBack Then
    KeyAscii = 0
    End If
End If
If KeyAscii = 39 Then KeyAscii = 0
If (BER_C(Index).Text & Chr(KeyAscii)) > 9999999.99 And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub BER_C_LostFocus(Index As Integer)
    Dim z As Double
    If BER_C(Index).Text <> "" Then
        z = CDbl(BER_C(Index).Text)
        If z > 9999999.99 Then
            BER_C(Index).Text = Left(z, 7)
        End If
    Else
        BER_C(Index) = Format(0, gs_formatExchangeRate)
    End If
    BER_C(Index).Text = Format(BER_C(Index).Text, gs_formatExchangeRate)
End Sub

Private Sub BER_D_Change(Index As Integer)
BER_D(Index).DataChanged = True
If InStr(1, BER_D(Index).Text, ",") = 1 Then BER_D(Index).Text = Right(BER_D(Index), Len(BER_D(Index)) - 1)
End Sub

Private Sub BER_D_GotFocus(Index As Integer)
BER_D(Index).SelStart = 0
BER_D(Index).SelLength = Len(BER_D(Index).Text)
End Sub

Private Sub BER_D_KeyPress(Index As Integer, KeyAscii As Integer)
If InStr("0123456789.,", Chr(KeyAscii)) = 0 Then
    If KeyAscii = 13 Then
      SendKeys vbTab
    ElseIf KeyAscii <> vbKeyBack Then
    KeyAscii = 0
    End If
End If
If KeyAscii = 39 Then KeyAscii = 0
If (BER_D(Index).Text & Chr(KeyAscii)) > 9999999.99 And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub BER_D_LostFocus(Index As Integer)
    Dim z As Double
    If BER_D(Index).Text <> "" Then
        z = CDbl(BER_D(Index).Text)
        If z > 9999999.99 Then
            BER_D(Index).Text = Left(z, 7)
        End If
    Else
        BER_D(Index) = Format(0, gs_formatExchangeRate)
    End If
    BER_D(Index).Text = Format(BER_D(Index).Text, gs_formatExchangeRate)
End Sub

Private Sub CboTerm_Change()
If CboTerm.MatchFound Then
   TermCls = CboTerm.List(CboTerm.ListIndex, 1)
   Call koneksi
   DoEvents
   Call Buka
   Call DisplayAda
Else
   DoEvents
   Call Kosong
End If
End Sub

Private Sub cmd_clear_Click()
TxtYear = ""
LblPesan = ""
CboTerm.ListIndex = 0
Call Kosong
End Sub

Private Sub Cmd_Save_Click()
Dim sql As String, RsSave As New ADODB.Recordset, RsSelek As New ADODB.Recordset
Dim i As Long, NilIndex, NilExch As String

If hakUpdate(Me.Name) = 0 Then LblPesan = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub

LblPesan = ""
If TxtYear = "" Then
   LblPesan = DisplayMsg("0007")     '"Please input year!"
   TxtYear.SetFocus
   Exit Sub
End If

sql = "select * from book_exchangerate where exch_year=" & Val(TxtYear) & " and term_cls=" & (TermCls) & ""
If RsSelek.State <> adStateClosed Then RsSelek.Close
RsSelek.Open sql, Db, adOpenKeyset, adLockOptimistic
If Not (RsSelek.BOF And RsSelek.EOF) Then
    'Update
    For i = 1 To 12
        If BER_A(i).DataChanged = True And Trim(code(1)) <> "" Then
            NilExch = "exch0" & i
            RsSelek.filter = " currency_code='" & (code(1)) & "'"
            If (RsSelek.BOF And RsSelek.EOF) Then
                RsSelek.AddNew
                RsSelek("exch_year") = Val(TxtYear)
                RsSelek("term_cls") = TermCls
                RsSelek("currency_code") = code(1)
            End If
            RsSelek(NilExch) = CDbl(BER_A(i))
            RsSelek("Last_Update") = Now
            RsSelek("Last_User") = userLogin
            RsSelek.update

            BER_A(i).DataChanged = False
            RsSelek.filter = ""
            RsSelek.Requery
        End If
        If BER_B(i).DataChanged = True And Trim(code(2)) <> "" Then
            NilExch = "exch0" & i
            RsSelek.filter = " currency_code='" & (code(2)) & "'"
            If (RsSelek.BOF And RsSelek.EOF) Then
                RsSelek.AddNew
                RsSelek("exch_year") = Val(TxtYear)
                RsSelek("term_cls") = TermCls
                RsSelek("currency_code") = code(2)
            End If
            RsSelek(NilExch) = CDbl(BER_B(i))
            RsSelek("Last_Update") = Now
            RsSelek("Last_User") = userLogin
            RsSelek.update
           
           BER_B(i).DataChanged = False
            RsSelek.filter = ""
            RsSelek.Requery
        End If
        If BER_C(i).DataChanged = True And Trim(code(3)) <> "" Then
            NilExch = "exch0" & i
            RsSelek.filter = " currency_code='" & (code(3)) & "'"
            If (RsSelek.BOF And RsSelek.EOF) Then
                RsSelek.AddNew
                RsSelek("exch_year") = Val(TxtYear)
                RsSelek("term_cls") = TermCls
                RsSelek("currency_code") = code(3)
            End If
            RsSelek(NilExch) = CDbl(BER_C(i))
            RsSelek("Last_Update") = Now
            RsSelek("Last_User") = userLogin
            RsSelek.update

           BER_C(i).DataChanged = False
            RsSelek.filter = ""
            RsSelek.Requery
        End If
        If BER_D(i).DataChanged = True And Trim(code(4)) <> "" Then
            NilExch = "exch0" & i
            RsSelek.filter = " currency_code='" & (code(4)) & "'"
            If (RsSelek.BOF And RsSelek.EOF) Then
                RsSelek.AddNew
                RsSelek("exch_year") = Val(TxtYear)
                RsSelek("term_cls") = TermCls
                RsSelek("currency_code") = code(4)
            End If
            RsSelek(NilExch) = CDbl(BER_D(i))
            RsSelek("Last_Update") = Now
            RsSelek("Last_User") = userLogin
            RsSelek.update

           BER_D(i).DataChanged = False
            RsSelek.filter = ""
            RsSelek.Requery
        End If
        If BER_E(i).DataChanged = True And Trim(code(5)) <> "" Then
            NilExch = "exch0" & i
            RsSelek.filter = " currency_code='" & (code(5)) & "'"
            If (RsSelek.BOF And RsSelek.EOF) Then
                RsSelek.AddNew
                RsSelek("exch_year") = Val(TxtYear)
                RsSelek("term_cls") = TermCls
                RsSelek("currency_code") = code(5)
            End If
            RsSelek(NilExch) = CDbl(BER_E(i))
            RsSelek("Last_Update") = Now
            RsSelek("Last_User") = userLogin
            RsSelek.update

           BER_E(i).DataChanged = False
            RsSelek.filter = ""
            RsSelek.Requery
        End If
        If BER_F(i).DataChanged = True And Trim(code(6)) <> "" Then
           NilExch = "exch0" & i
            RsSelek.filter = " currency_code='" & (code(6)) & "'"
            If (RsSelek.BOF And RsSelek.EOF) Then
                RsSelek.AddNew
                RsSelek("exch_year") = Val(TxtYear)
                RsSelek("term_cls") = TermCls
                RsSelek("currency_code") = code(6)
            End If
            RsSelek(NilExch) = CDbl(BER_F(i))
            RsSelek("Last_Update") = Now
            RsSelek("Last_User") = userLogin
            RsSelek.update

           BER_F(i).DataChanged = False
            RsSelek.filter = ""
            RsSelek.Requery
        End If
    Next i
    LblPesan = DisplayMsg(1101)
Else
    'Insert
    Dim X As Integer, NilaiMasuk As String, Dumy As String, NamaField As String
    On Error Resume Next
    
    For i = 1 To 6
      If i = 1 Then
      NilaiMasuk = ""
      NamaField = ""
        For X = 1 To 12
          If BER_A(X).DataChanged = True And Trim(code(1)) <> "" Then
            NilaiMasuk = NilaiMasuk & "," & CDbl(BER_A(X))
            NamaField = NamaField & "," & "exch0" & X
          End If
        Next X
       ElseIf i = 2 Then
       NilaiMasuk = ""
       NamaField = ""
        For X = 1 To 12
          If BER_B(X).DataChanged = True And Trim(code(2)) <> "" Then
            NilaiMasuk = NilaiMasuk & "," & CDbl(BER_B(X))
            NamaField = NamaField & "," & "exch0" & X
          End If
        Next X
       ElseIf i = 3 Then
       NilaiMasuk = ""
       NamaField = ""
        For X = 1 To 12
          If BER_C(X).DataChanged = True And Trim(code(3)) <> "" Then
            NilaiMasuk = NilaiMasuk & "," & CDbl(BER_C(X))
            NamaField = NamaField & "," & "exch0" & X
          End If
        Next X
       ElseIf i = 4 Then
       NilaiMasuk = ""
       NamaField = ""
        For X = 1 To 12
          If BER_D(X).DataChanged = True And Trim(code(4)) <> "" Then
            NilaiMasuk = NilaiMasuk & "," & CDbl(BER_D(X))
            NamaField = NamaField & "," & "exch0" & X
          End If
        Next X
       ElseIf i = 5 Then
       NilaiMasuk = ""
       NamaField = ""
        For X = 1 To 12
          If BER_E(X).DataChanged = True And Trim(code(5)) <> "" Then
            NilaiMasuk = NilaiMasuk & "," & CDbl(BER_E(X))
            NamaField = NamaField & "," & "exch0" & X
          End If
        Next X
       ElseIf i = 6 Then
       NilaiMasuk = ""
       NamaField = ""
        For X = 1 To 12
          If BER_F(X).DataChanged = True And Trim(code(6)) <> "" Then
            NilaiMasuk = NilaiMasuk & "," & CDbl(BER_F(X))
            NamaField = NamaField & "," & "exch0" & X
          End If
        Next X
        
       End If
        If NamaField <> "" Then
            sql = "insert into book_exchangerate (exch_year,term_cls,currency_code" & NamaField & ") " & _
                  "values(" & TxtYear & ",'" & TermCls & "','" & code(i) & "'" & _
                  "" & NilaiMasuk & ")"
            If err.number <> 0 Then
               If InStr(1, err.Description, "Violation of PRIMARY KEY") > 0 Then
                  LblPesan.Caption = ""
               Else
               End If
            End If
            Db.Execute (sql)
        End If
    Next i
    LblPesan = DisplayMsg(1000)
End If
End Sub

Private Sub CommandButton1_Click()
DoEvents
frmMainMenu.Show
DoEvents
Unload Me
End Sub

Private Sub TermRate()
CboTerm.clear
CboTerm.columnCount = 2
CboTerm.TextColumn = 1

CboTerm.AddItem ""
CboTerm.List(0, 0) = "Beginning"
CboTerm.List(0, 1) = "1"

CboTerm.AddItem ""
CboTerm.List(1, 0) = "Ending"
CboTerm.List(1, 1) = "2"

If CboTerm.ListCount > 0 Then CboTerm.ListIndex = 0
TermCls = CboTerm.List(CboTerm.ListIndex, 1)

CboTerm.ColumnWidths = "60 pt; 0 pt"
CboTerm.ListWidth = 60
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    LblPesan.Caption = ErrMsg
End If
End Sub

Private Sub Form_Load()
If gb_Simulation = True Then Call up_InitSimulation(F_EXCHANGE)

CtrlMenu1.FormName = Me.Name
Me.Caption = "Book Exchange Rate"
Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"
Call Kosong
Call ShowCurr
Call TermRate
End Sub

Private Sub BER_E_Change(Index As Integer)
BER_E(Index).DataChanged = True
If InStr(1, BER_E(Index).Text, ",") = 1 Then BER_E(Index).Text = Right(BER_E(Index), Len(BER_E(Index)) - 1)
End Sub

Private Sub BER_E_GotFocus(Index As Integer)
BER_E(Index).SelStart = 0
BER_E(Index).SelLength = Len(BER_E(Index).Text)
End Sub

Private Sub BER_E_KeyPress(Index As Integer, KeyAscii As Integer)
If InStr("0123456789.,", Chr(KeyAscii)) = 0 Then
    If KeyAscii = 13 Then
      SendKeys vbTab
    ElseIf KeyAscii <> vbKeyBack Then
    KeyAscii = 0
    End If
End If
If KeyAscii = 39 Then KeyAscii = 0
If (BER_E(Index).Text & Chr(KeyAscii)) > 9999999.99 And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub BER_E_LostFocus(Index As Integer)
    Dim z As Double
    If BER_E(Index).Text <> "" Then
        z = CDbl(BER_E(Index).Text)
        If z > 9999999.99 Then
            BER_E(Index).Text = Left(z, 7)
        End If
    Else
        BER_E(Index) = Format(0, gs_formatExchangeRate)
    End If
    BER_E(Index).Text = Format(BER_E(Index).Text, gs_formatExchangeRate)
End Sub

Private Sub BER_F_Change(Index As Integer)
BER_F(Index).DataChanged = True
If InStr(1, BER_F(Index).Text, ",") = 1 Then BER_F(Index).Text = Right(BER_F(Index), Len(BER_F(Index)) - 1)
End Sub

Private Sub BER_F_GotFocus(Index As Integer)
BER_F(Index).SelStart = 0
BER_F(Index).SelLength = Len(BER_F(Index).Text)
End Sub

Private Sub BER_F_KeyPress(Index As Integer, KeyAscii As Integer)
If InStr("0123456789.,", Chr(KeyAscii)) = 0 Then
    If KeyAscii = 13 Then
      SendKeys vbTab
    ElseIf KeyAscii <> vbKeyBack Then
    KeyAscii = 0
    End If
End If
If KeyAscii = 39 Then KeyAscii = 0
If (BER_F(Index).Text & Chr(KeyAscii)) > 9999999.99 And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub BER_F_LostFocus(Index As Integer)
    Dim z As Double
    If BER_F(Index).Text <> "" Then
        z = CDbl(BER_F(Index).Text)
        If z > 9999999.99 Then
            BER_F(Index).Text = Left(z, 7)
        End If
    Else
        BER_F(Index) = Format(0, gs_formatExchangeRate)
    End If
    BER_F(Index).Text = Format(BER_F(Index).Text, gs_formatExchangeRate)
End Sub

Private Sub TxtYear_Change()
Call koneksi
If Not (RsCurr.BOF And RsCurr.EOF) Then
  DoEvents
  LblPesan = ""
  Call Buka
  Call Kosong
  Call DisplayAda
Else
  DoEvents
  If Len(TxtYear) <> 4 Then
     LblPesan = DisplayMsg(4039)    '"Year is not found!"
  Else
     LblPesan = ""
  End If
  Call Kosong
End If
End Sub

Private Sub TxtYear_KeyPress(KeyAscii As Integer)
If InStr("0123456789.,", Chr(KeyAscii)) = 0 Then
    If KeyAscii = 13 Then
      SendKeys vbTab
    ElseIf KeyAscii <> vbKeyBack Then
    KeyAscii = 0
    End If
End If
If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub DisplayAda()
Dim i, X As Integer, ErCode1 As String, Beri(12)
Dim ExCode2 As String

Call Kosong
If Not (RsCurr.BOF And RsCurr.EOF) Then
Do While Not RsCurr.EOF
   If code(1) = RsCurr!currency_code Then
      For i = 1 To 12
          Exchange = "exch0" & i
          BER_A(i) = IIf(IsNull(RsCurr.Fields(Exchange)), 0, RsCurr.Fields(Exchange))
          BER_A(i) = Format(BER_A(i), gs_formatExchangeRate)
          BER_A(i).DataChanged = False
      Next i
   ElseIf code(2) = RsCurr!currency_code Then
      For i = 1 To 12
          Exchange = "exch0" & i
          BER_B(i) = IIf(IsNull(RsCurr.Fields(Exchange)), 0, RsCurr.Fields(Exchange))
          BER_B(i) = Format(BER_B(i), gs_formatExchangeRate)
          BER_B(i).DataChanged = False
      Next i
   ElseIf code(3) = RsCurr!currency_code Then
      For i = 1 To 12
          Exchange = "exch0" & i
          BER_C(i) = IIf(IsNull(RsCurr.Fields(Exchange)), 0, RsCurr.Fields(Exchange))
          BER_C(i) = Format(BER_C(i), gs_formatExchangeRate)
          BER_C(i).DataChanged = False
      Next i
   ElseIf code(4) = RsCurr!currency_code Then
      For i = 1 To 12
          Exchange = "exch0" & i
          BER_D(i) = IIf(IsNull(RsCurr.Fields(Exchange)), 0, RsCurr.Fields(Exchange))
          BER_D(i) = Format(BER_D(i), gs_formatExchangeRate)
          BER_D(i).DataChanged = False
      Next i
   ElseIf code(5) = RsCurr!currency_code Then
      For i = 1 To 12
          Exchange = "exch0" & i
          BER_E(i) = IIf(IsNull(RsCurr.Fields(Exchange)), 0, RsCurr.Fields(Exchange))
          BER_E(i) = Format(BER_E(i), gs_formatExchangeRate)
          BER_E(i).DataChanged = False
      Next i
   ElseIf code(6) = RsCurr!currency_code Then
      For i = 1 To 12
          Exchange = "exch0" & i
          BER_F(i) = IIf(IsNull(RsCurr.Fields(Exchange)), 0, RsCurr.Fields(Exchange))
          BER_F(i) = Format(BER_F(i), gs_formatExchangeRate)
          BER_F(i).DataChanged = False
      Next i
   End If
  RsCurr.MoveNext
Loop
End If
End Sub

Private Sub koneksi()
Dim sql As String
If RsCurr.State <> adStateClosed Then RsCurr.Close
sql = "select * from book_exchangerate where exch_year=" & Val(TxtYear) & " and term_cls=" & Val(TermCls) & ""
RsCurr.Open sql, Db, adOpenKeyset, adLockOptimistic
End Sub

Private Sub Buka()
    Dim X As Integer
    For X = 1 To 12
        BER_A(X).Enabled = True
        BER_B(X).Enabled = True
        BER_C(X).Enabled = True
        BER_D(X).Enabled = True
        BER_E(X).Enabled = True
        BER_F(X).Enabled = True
    Next X
    For X = 1 To 6
        code(X).Enabled = True
    Next X
    If TxtYear = "" Then
        Call Kosong
    End If
End Sub

Private Sub Kosong()
Dim i As Integer
For i = 1 To 12
    BER_A(i) = Format(0, gs_formatExchangeRate)
    BER_A(i).DataChanged = False
    BER_B(i) = Format(0, gs_formatExchangeRate)
    BER_B(i).DataChanged = False
    BER_C(i) = Format(0, gs_formatExchangeRate)
    BER_C(i).DataChanged = False
    BER_D(i) = Format(0, gs_formatExchangeRate)
    BER_D(i).DataChanged = False
    BER_E(i) = Format(0, gs_formatExchangeRate)
    BER_E(i).DataChanged = False
    BER_F(i) = Format(0, gs_formatExchangeRate)
    BER_F(i).DataChanged = False
Next i
End Sub
