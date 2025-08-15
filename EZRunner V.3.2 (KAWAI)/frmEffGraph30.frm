VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmEffGraph30 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Graph"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13260
   FillColor       =   &H00FFC0C0&
   Icon            =   "frmEffGraph30.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   13260
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   645
      Left            =   11430
      TabIndex        =   112
      Top             =   810
      Width           =   1455
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plan"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   630
         TabIndex        =   114
         Top             =   180
         Width           =   360
      End
      Begin VB.Line LineTbl 
         BorderColor     =   &H00FF8080&
         BorderStyle     =   3  'Dot
         Index           =   0
         X1              =   165
         X2              =   500
         Y1              =   270
         Y2              =   270
      End
      Begin VB.Line LineTbl 
         BorderColor     =   &H00FF00FF&
         Index           =   1
         X1              =   180
         X2              =   500
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Result"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   630
         TabIndex        =   113
         Top             =   390
         Width           =   525
      End
   End
   Begin VB.CommandButton cmdReport 
      BackColor       =   &H00C0C0C0&
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
      Left            =   1020
      Picture         =   "frmEffGraph30.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   90
      Width           =   435
   End
   Begin MSComDlg.CommonDialog cdl 
      Left            =   1710
      Top             =   30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   30
      Left            =   12120
      TabIndex        =   111
      Top             =   1920
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblN3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Left            =   1980
      TabIndex        =   110
      Top             =   1350
      Width           =   570
   End
   Begin VB.Label LblN2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Left            =   1410
      TabIndex        =   109
      Top             =   1350
      Width           =   570
   End
   Begin VB.Label lblN 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Left            =   840
      TabIndex        =   108
      Top             =   1350
      Width           =   570
   End
   Begin VB.Label lblY 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Index           =   9
      Left            =   510
      TabIndex        =   107
      Top             =   2380
      Width           =   750
   End
   Begin VB.Label lblY 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Index           =   8
      Left            =   510
      TabIndex        =   106
      Top             =   2880
      Width           =   750
   End
   Begin VB.Line LineValue 
      Index           =   9
      X1              =   1350
      X2              =   1440
      Y1              =   2500
      Y2              =   2500
   End
   Begin VB.Line LineValue 
      Index           =   8
      X1              =   1350
      X2              =   1440
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Period :"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   2
      Left            =   6375
      TabIndex        =   105
      Top             =   1260
      Width           =   645
   End
   Begin VB.Shape Shape1 
      Height          =   6105
      Left            =   510
      Top             =   1590
      Width           =   12375
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   20
      Left            =   8520
      TabIndex        =   104
      Top             =   1920
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   29
      Left            =   11760
      TabIndex        =   103
      Top             =   1920
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   28
      Left            =   11400
      TabIndex        =   102
      Top             =   1920
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   27
      Left            =   11040
      TabIndex        =   101
      Top             =   1920
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   26
      Left            =   10680
      TabIndex        =   100
      Top             =   1920
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   25
      Left            =   10320
      TabIndex        =   99
      Top             =   1920
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   24
      Left            =   9960
      TabIndex        =   98
      Top             =   1920
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   23
      Left            =   9600
      TabIndex        =   97
      Top             =   1920
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   22
      Left            =   9240
      TabIndex        =   96
      Top             =   1920
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   21
      Left            =   8880
      TabIndex        =   95
      Top             =   1920
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   19
      Left            =   8160
      TabIndex        =   94
      Top             =   1920
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   18
      Left            =   7800
      TabIndex        =   93
      Top             =   1920
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   17
      Left            =   7440
      TabIndex        =   92
      Top             =   1920
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   16
      Left            =   7080
      TabIndex        =   91
      Top             =   1920
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   15
      Left            =   6720
      TabIndex        =   90
      Top             =   1920
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   14
      Left            =   6360
      TabIndex        =   89
      Top             =   1920
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   13
      Left            =   6000
      TabIndex        =   88
      Top             =   1920
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   12
      Left            =   5640
      TabIndex        =   87
      Top             =   1920
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   11
      Left            =   5280
      TabIndex        =   86
      Top             =   1920
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   10
      Left            =   4920
      TabIndex        =   85
      Top             =   1920
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   9
      Left            =   4560
      TabIndex        =   84
      Top             =   1920
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   8
      Left            =   4200
      TabIndex        =   83
      Top             =   1920
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   7
      Left            =   3840
      TabIndex        =   82
      Top             =   1920
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   6
      Left            =   3480
      TabIndex        =   81
      Top             =   1920
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   5
      Left            =   3090
      TabIndex        =   80
      Top             =   1920
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Left            =   2760
      TabIndex        =   79
      Top             =   1920
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Left            =   2400
      TabIndex        =   78
      Top             =   1920
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Left            =   2040
      TabIndex        =   77
      Top             =   1920
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Left            =   1680
      TabIndex        =   76
      Top             =   1920
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Left            =   1320
      TabIndex        =   75
      Top             =   1920
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   20
      Left            =   8550
      TabIndex        =   74
      Top             =   1650
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "21"
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
      Height          =   255
      Index           =   20
      Left            =   8550
      TabIndex        =   73
      Top             =   7110
      Width           =   360
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   30
      Left            =   12150
      TabIndex        =   72
      Top             =   1650
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   29
      Left            =   11790
      TabIndex        =   71
      Top             =   1650
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   28
      Left            =   11430
      TabIndex        =   70
      Top             =   1650
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   27
      Left            =   11070
      TabIndex        =   69
      Top             =   1650
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   26
      Left            =   10710
      TabIndex        =   68
      Top             =   1650
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   25
      Left            =   10350
      TabIndex        =   67
      Top             =   1650
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   24
      Left            =   9990
      TabIndex        =   66
      Top             =   1650
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   23
      Left            =   9630
      TabIndex        =   65
      Top             =   1650
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   22
      Left            =   9270
      TabIndex        =   64
      Top             =   1650
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   21
      Left            =   8910
      TabIndex        =   63
      Top             =   1650
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   19
      Left            =   8190
      TabIndex        =   62
      Top             =   1650
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   18
      Left            =   7830
      TabIndex        =   61
      Top             =   1650
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   17
      Left            =   7470
      TabIndex        =   60
      Top             =   1650
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   16
      Left            =   7110
      TabIndex        =   59
      Top             =   1650
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   15
      Left            =   6750
      TabIndex        =   58
      Top             =   1650
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   14
      Left            =   6390
      TabIndex        =   57
      Top             =   1650
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   13
      Left            =   6030
      TabIndex        =   56
      Top             =   1650
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   12
      Left            =   5670
      TabIndex        =   55
      Top             =   1650
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   11
      Left            =   5310
      TabIndex        =   54
      Top             =   1650
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   10
      Left            =   4950
      TabIndex        =   53
      Top             =   1650
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   9
      Left            =   4590
      TabIndex        =   52
      Top             =   1650
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   8
      Left            =   4230
      TabIndex        =   51
      Top             =   1650
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   7
      Left            =   3870
      TabIndex        =   50
      Top             =   1650
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   6
      Left            =   3510
      TabIndex        =   49
      Top             =   1650
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   5
      Left            =   3120
      TabIndex        =   48
      Top             =   1650
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Left            =   2790
      TabIndex        =   47
      Top             =   1650
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Left            =   2430
      TabIndex        =   46
      Top             =   1650
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Left            =   2070
      TabIndex        =   45
      Top             =   1650
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Left            =   1710
      TabIndex        =   44
      Top             =   1650
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Left            =   1350
      TabIndex        =   43
      Top             =   1650
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H00FF00FF&
      Index           =   30
      X1              =   10470
      X2              =   10830
      Y1              =   7050
      Y2              =   4650
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H00FF00FF&
      Index           =   29
      X1              =   10350
      X2              =   10710
      Y1              =   7050
      Y2              =   4650
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H00FF00FF&
      Index           =   28
      X1              =   10200
      X2              =   10560
      Y1              =   7050
      Y2              =   4650
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H00FF00FF&
      Index           =   27
      X1              =   10020
      X2              =   10380
      Y1              =   7050
      Y2              =   4650
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H00FF00FF&
      Index           =   26
      X1              =   9870
      X2              =   10230
      Y1              =   7050
      Y2              =   4650
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H00FF00FF&
      Index           =   25
      X1              =   9750
      X2              =   10110
      Y1              =   7050
      Y2              =   4650
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H00FF00FF&
      Index           =   24
      X1              =   9630
      X2              =   9990
      Y1              =   7050
      Y2              =   4650
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H00FF00FF&
      Index           =   23
      X1              =   9480
      X2              =   9840
      Y1              =   7050
      Y2              =   4650
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H00FF00FF&
      Index           =   22
      X1              =   9360
      X2              =   9720
      Y1              =   7050
      Y2              =   4650
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H00FF00FF&
      Index           =   21
      X1              =   9180
      X2              =   9540
      Y1              =   7050
      Y2              =   4650
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H00FF00FF&
      Index           =   20
      X1              =   8970
      X2              =   9330
      Y1              =   7050
      Y2              =   4650
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H00FF00FF&
      Index           =   19
      X1              =   8730
      X2              =   9090
      Y1              =   7050
      Y2              =   4650
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H00FF00FF&
      Index           =   18
      X1              =   8580
      X2              =   8940
      Y1              =   7050
      Y2              =   4650
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H00FF00FF&
      Index           =   17
      X1              =   8400
      X2              =   8760
      Y1              =   7050
      Y2              =   4650
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H00FF00FF&
      Index           =   16
      X1              =   8220
      X2              =   8580
      Y1              =   7050
      Y2              =   4650
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H00FF00FF&
      Index           =   15
      X1              =   8070
      X2              =   8430
      Y1              =   7050
      Y2              =   4650
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H00FF00FF&
      Index           =   14
      X1              =   7860
      X2              =   8220
      Y1              =   7050
      Y2              =   4650
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H00FF00FF&
      Index           =   13
      X1              =   7710
      X2              =   8070
      Y1              =   7050
      Y2              =   4650
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H00FF00FF&
      Index           =   12
      X1              =   7560
      X2              =   7920
      Y1              =   7050
      Y2              =   4650
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H00FF00FF&
      Index           =   11
      X1              =   7440
      X2              =   7800
      Y1              =   7050
      Y2              =   4650
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H00FF00FF&
      Index           =   10
      X1              =   7260
      X2              =   7620
      Y1              =   7050
      Y2              =   4650
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H00FF00FF&
      Index           =   9
      X1              =   7110
      X2              =   7470
      Y1              =   7050
      Y2              =   4650
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H00FF00FF&
      Index           =   8
      X1              =   6990
      X2              =   7350
      Y1              =   7050
      Y2              =   4650
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H00FF00FF&
      Index           =   7
      X1              =   6810
      X2              =   7170
      Y1              =   7050
      Y2              =   4650
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H00FF00FF&
      Index           =   6
      X1              =   6660
      X2              =   7020
      Y1              =   7050
      Y2              =   4650
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H00FF00FF&
      Index           =   5
      X1              =   6540
      X2              =   6900
      Y1              =   7050
      Y2              =   4650
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H00FF00FF&
      Index           =   4
      X1              =   6420
      X2              =   6780
      Y1              =   7050
      Y2              =   4650
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H00FF00FF&
      Index           =   3
      X1              =   6300
      X2              =   6660
      Y1              =   7050
      Y2              =   4650
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H00FF00FF&
      Index           =   2
      X1              =   6180
      X2              =   6540
      Y1              =   7050
      Y2              =   4650
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H00FF00FF&
      Index           =   1
      X1              =   6060
      X2              =   6420
      Y1              =   7050
      Y2              =   4650
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H00FF00FF&
      Index           =   0
      X1              =   5970
      X2              =   6330
      Y1              =   7050
      Y2              =   4650
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      BorderStyle     =   3  'Dot
      Index           =   30
      X1              =   5550
      X2              =   5910
      Y1              =   7040
      Y2              =   4650
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      BorderStyle     =   3  'Dot
      Index           =   29
      X1              =   5430
      X2              =   5790
      Y1              =   7040
      Y2              =   4650
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      BorderStyle     =   3  'Dot
      Index           =   28
      X1              =   5280
      X2              =   5640
      Y1              =   7040
      Y2              =   4650
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      BorderStyle     =   3  'Dot
      Index           =   27
      X1              =   5160
      X2              =   5520
      Y1              =   7040
      Y2              =   4650
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      BorderStyle     =   3  'Dot
      Index           =   26
      X1              =   5070
      X2              =   5430
      Y1              =   7040
      Y2              =   4650
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      BorderStyle     =   3  'Dot
      Index           =   25
      X1              =   4920
      X2              =   5280
      Y1              =   7040
      Y2              =   4650
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      BorderStyle     =   3  'Dot
      Index           =   24
      X1              =   4770
      X2              =   5130
      Y1              =   7040
      Y2              =   4650
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      BorderStyle     =   3  'Dot
      Index           =   23
      X1              =   4650
      X2              =   5010
      Y1              =   7040
      Y2              =   4650
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      BorderStyle     =   3  'Dot
      Index           =   22
      X1              =   4470
      X2              =   4830
      Y1              =   7040
      Y2              =   4650
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      BorderStyle     =   3  'Dot
      Index           =   21
      X1              =   4290
      X2              =   4650
      Y1              =   7040
      Y2              =   4650
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      BorderStyle     =   3  'Dot
      Index           =   20
      X1              =   4170
      X2              =   4530
      Y1              =   7040
      Y2              =   4650
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      BorderStyle     =   3  'Dot
      Index           =   19
      X1              =   4020
      X2              =   4380
      Y1              =   7040
      Y2              =   4650
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      BorderStyle     =   3  'Dot
      Index           =   18
      X1              =   3930
      X2              =   4290
      Y1              =   7040
      Y2              =   4650
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      BorderStyle     =   3  'Dot
      Index           =   17
      X1              =   3840
      X2              =   4200
      Y1              =   7040
      Y2              =   4650
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      BorderStyle     =   3  'Dot
      Index           =   16
      X1              =   3660
      X2              =   4020
      Y1              =   7040
      Y2              =   4650
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      BorderStyle     =   3  'Dot
      Index           =   15
      X1              =   3540
      X2              =   3900
      Y1              =   7040
      Y2              =   4650
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      BorderStyle     =   3  'Dot
      Index           =   14
      X1              =   3390
      X2              =   3750
      Y1              =   7040
      Y2              =   4650
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      BorderStyle     =   3  'Dot
      Index           =   13
      X1              =   3270
      X2              =   3630
      Y1              =   7040
      Y2              =   4650
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      BorderStyle     =   3  'Dot
      Index           =   12
      X1              =   3060
      X2              =   3420
      Y1              =   7040
      Y2              =   4650
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      BorderStyle     =   3  'Dot
      Index           =   11
      X1              =   2940
      X2              =   3300
      Y1              =   7040
      Y2              =   4650
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      BorderStyle     =   3  'Dot
      Index           =   10
      X1              =   2790
      X2              =   3150
      Y1              =   7040
      Y2              =   4650
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      BorderStyle     =   3  'Dot
      Index           =   9
      X1              =   2640
      X2              =   3000
      Y1              =   7040
      Y2              =   4650
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      BorderStyle     =   3  'Dot
      Index           =   8
      X1              =   2520
      X2              =   2880
      Y1              =   7040
      Y2              =   4650
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "31"
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
      Height          =   255
      Index           =   30
      Left            =   12150
      TabIndex        =   42
      Top             =   7110
      Width           =   360
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "30"
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
      Height          =   255
      Index           =   29
      Left            =   11790
      TabIndex        =   41
      Top             =   7110
      Width           =   360
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "29"
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
      Height          =   255
      Index           =   28
      Left            =   11430
      TabIndex        =   40
      Top             =   7110
      Width           =   360
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "28"
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
      Height          =   255
      Index           =   27
      Left            =   11070
      TabIndex        =   39
      Top             =   7110
      Width           =   360
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "27"
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
      Height          =   255
      Index           =   26
      Left            =   10710
      TabIndex        =   38
      Top             =   7110
      Width           =   360
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "26"
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
      Height          =   255
      Index           =   25
      Left            =   10350
      TabIndex        =   37
      Top             =   7110
      Width           =   360
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "25"
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
      Height          =   255
      Index           =   24
      Left            =   9990
      TabIndex        =   36
      Top             =   7110
      Width           =   360
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "24"
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
      Height          =   255
      Index           =   23
      Left            =   9630
      TabIndex        =   35
      Top             =   7110
      Width           =   360
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "23"
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
      Height          =   255
      Index           =   22
      Left            =   9270
      TabIndex        =   34
      Top             =   7110
      Width           =   360
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "22"
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
      Height          =   255
      Index           =   21
      Left            =   8910
      TabIndex        =   33
      Top             =   7110
      Width           =   360
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "20"
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
      Height          =   255
      Index           =   19
      Left            =   8190
      TabIndex        =   32
      Top             =   7110
      Width           =   360
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "19"
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
      Height          =   255
      Index           =   18
      Left            =   7830
      TabIndex        =   31
      Top             =   7110
      Width           =   360
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "18"
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
      Height          =   255
      Index           =   17
      Left            =   7470
      TabIndex        =   30
      Top             =   7110
      Width           =   360
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "17"
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
      Height          =   255
      Index           =   16
      Left            =   7110
      TabIndex        =   29
      Top             =   7110
      Width           =   360
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "16"
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
      Height          =   255
      Index           =   15
      Left            =   6750
      TabIndex        =   28
      Top             =   7110
      Width           =   360
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "15"
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
      Height          =   255
      Index           =   14
      Left            =   6390
      TabIndex        =   27
      Top             =   7110
      Width           =   360
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "14"
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
      Height          =   255
      Index           =   13
      Left            =   6030
      TabIndex        =   26
      Top             =   7110
      Width           =   360
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "13"
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
      Height          =   255
      Index           =   12
      Left            =   5670
      TabIndex        =   25
      Top             =   7110
      Width           =   360
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "12"
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
      Height          =   255
      Index           =   11
      Left            =   5310
      TabIndex        =   24
      Top             =   7110
      Width           =   360
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "11"
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
      Height          =   255
      Index           =   10
      Left            =   4950
      TabIndex        =   23
      Top             =   7110
      Width           =   360
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "10"
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
      Height          =   255
      Index           =   9
      Left            =   4590
      TabIndex        =   22
      Top             =   7110
      Width           =   360
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "9"
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
      Height          =   255
      Index           =   8
      Left            =   4230
      TabIndex        =   21
      Top             =   7110
      Width           =   360
   End
   Begin VB.Line LinePrt 
      BorderWidth     =   2
      X1              =   0
      X2              =   13270
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Machine No :"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   1
      Left            =   6150
      TabIndex        =   19
      Top             =   1020
      Width           =   1095
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Production Schedule / Result Control"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   390
      Index           =   0
      Left            =   4185
      TabIndex        =   18
      Top             =   660
      Width           =   5025
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Print"
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
      Left            =   510
      TabIndex        =   17
      Top             =   180
      Width           =   390
   End
   Begin VB.Label lblValue0 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   915
      TabIndex        =   16
      Top             =   7110
      Width           =   345
   End
   Begin VB.Label lblY 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Index           =   7
      Left            =   510
      TabIndex        =   15
      Top             =   3380
      Width           =   750
   End
   Begin VB.Label lblY 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Index           =   6
      Left            =   510
      TabIndex        =   14
      Top             =   3880
      Width           =   750
   End
   Begin VB.Label lblY 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Index           =   5
      Left            =   510
      TabIndex        =   13
      Top             =   4380
      Width           =   750
   End
   Begin VB.Label lblY 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Left            =   510
      TabIndex        =   12
      Top             =   4880
      Width           =   750
   End
   Begin VB.Label lblY 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Left            =   510
      TabIndex        =   11
      Top             =   5380
      Width           =   750
   End
   Begin VB.Label lblY 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Left            =   510
      TabIndex        =   10
      Top             =   5780
      Width           =   750
   End
   Begin VB.Label lblY 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Left            =   510
      TabIndex        =   9
      Top             =   6130
      Width           =   750
   End
   Begin VB.Label lblY 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Left            =   510
      TabIndex        =   8
      Top             =   6630
      Width           =   750
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      BorderStyle     =   3  'Dot
      Index           =   7
      X1              =   2370
      X2              =   2730
      Y1              =   7040
      Y2              =   4650
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      BorderStyle     =   3  'Dot
      Index           =   6
      X1              =   2250
      X2              =   2610
      Y1              =   7040
      Y2              =   4650
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      BorderStyle     =   3  'Dot
      Index           =   5
      X1              =   2070
      X2              =   2430
      Y1              =   7040
      Y2              =   4650
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      BorderStyle     =   3  'Dot
      Index           =   4
      X1              =   1890
      X2              =   2250
      Y1              =   7040
      Y2              =   4650
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      BorderStyle     =   3  'Dot
      Index           =   3
      X1              =   1770
      X2              =   2130
      Y1              =   7040
      Y2              =   4650
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      BorderStyle     =   3  'Dot
      Index           =   2
      X1              =   1650
      X2              =   2010
      Y1              =   7040
      Y2              =   4650
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      BorderStyle     =   3  'Dot
      Index           =   1
      X1              =   1530
      X2              =   1890
      Y1              =   7040
      Y2              =   4650
   End
   Begin VB.Line LineValue 
      Index           =   7
      X1              =   1350
      X2              =   1440
      Y1              =   3500
      Y2              =   3500
   End
   Begin VB.Line LineValue 
      Index           =   6
      X1              =   1350
      X2              =   1440
      Y1              =   4000
      Y2              =   4000
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      BorderStyle     =   3  'Dot
      Index           =   0
      X1              =   1350
      X2              =   1710
      Y1              =   7115
      Y2              =   4650
   End
   Begin VB.Line LineYValue 
      X1              =   1350
      X2              =   1350
      Y1              =   2310
      Y2              =   7110
   End
   Begin VB.Line LineX 
      X1              =   1350
      X2              =   12510
      Y1              =   7110
      Y2              =   7110
   End
   Begin VB.Line LineValue 
      Index           =   3
      X1              =   1350
      X2              =   1440
      Y1              =   5500
      Y2              =   5500
   End
   Begin VB.Line LineValue 
      Index           =   5
      X1              =   1350
      X2              =   1440
      Y1              =   4500
      Y2              =   4500
   End
   Begin VB.Line LineValue 
      Index           =   1
      X1              =   1350
      X2              =   1440
      Y1              =   6250
      Y2              =   6250
   End
   Begin VB.Line LineValue 
      Index           =   2
      X1              =   1350
      X2              =   1440
      Y1              =   5900
      Y2              =   5900
   End
   Begin VB.Line LineValue 
      Index           =   0
      X1              =   1350
      X2              =   1440
      Y1              =   6750
      Y2              =   6750
   End
   Begin VB.Line LineValue 
      Index           =   4
      X1              =   1350
      X2              =   1440
      Y1              =   5000
      Y2              =   5000
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "8"
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
      Height          =   255
      Index           =   7
      Left            =   3870
      TabIndex        =   7
      Top             =   7110
      Width           =   360
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "7"
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
      Height          =   255
      Index           =   6
      Left            =   3510
      TabIndex        =   6
      Top             =   7110
      Width           =   360
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "6"
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
      Height          =   255
      Index           =   5
      Left            =   3150
      TabIndex        =   5
      Top             =   7110
      Width           =   360
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "5"
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
      Height          =   255
      Index           =   4
      Left            =   2790
      TabIndex        =   4
      Top             =   7110
      Width           =   360
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "4"
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
      Height          =   255
      Index           =   3
      Left            =   2430
      TabIndex        =   3
      Top             =   7110
      Width           =   360
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "3"
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
      Height          =   255
      Index           =   2
      Left            =   2070
      TabIndex        =   2
      Top             =   7110
      Width           =   360
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2"
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
      Height          =   255
      Index           =   1
      Left            =   1710
      TabIndex        =   1
      Top             =   7110
      Width           =   360
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
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
      Height          =   255
      Index           =   0
      Left            =   1350
      TabIndex        =   0
      Top             =   7110
      Width           =   360
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmEffGraph30"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i  As Integer

Dim XLeft As Double, XRight As Double, XTop As Double, YTop As Double
Dim RangeY As Double, MaxY As Double, totLineY As Integer, distY As Double
Dim Percentage As Double, Percentage2 As Double

Public jmlBar As Double, MaxQty As Double
Public MaxDailyDt As Integer, MaxResultDt As Integer

Public Sub viewGraph()
    XLeft = LineX.X1: XRight = LineX.x2: XTop = LineX.Y1: YTop = LineYValue.Y1
    RangeY = setRange(MaxQty)
    MaxY = RoundUp(MaxQty / RangeY) * RangeY
    totLineY = MaxY / RangeY
    distY = ((XTop - YTop) / MaxY)
   
    Percentage = 0: Percentage2 = 0
    For i = 0 To 30
        '******************* Set Y *******************
        If i < 10 Then
            If i < totLineY Then
                lblY(i).Visible = True: LineValue(i).Visible = True
                
                lblY(i) = Format((MaxY / totLineY) * (i + 1), gs_formatQty)
                lblY(i).top = XTop - (lblY(i) * distY)
                LineValue(i).Y1 = lblY(i).top: LineValue(i).Y2 = LineValue(i).Y1
            Else
                lblY(i).Visible = False: LineValue(i).Visible = False
            End If
        End If
        
        If i < jmlBar Then
            GLine1(i).Visible = True: GLine2(i).Visible = True

            Percentage = CDbl(lblBarVal1(i)) * distY
            Percentage2 = CDbl(lblBarVal2(i)) * distY

            '*********** Set Top & Height *****************
            If i = 0 Then
                GLine1(i).X1 = lblX(i).Left: GLine1(i).Y1 = XTop
                GLine2(i).X1 = lblX(i).Left: GLine2(i).Y1 = XTop
            Else
                GLine1(i).X1 = GLine1(i - 1).x2: GLine1(i).Y1 = GLine1(i - 1).Y2
                GLine2(i).X1 = GLine2(i - 1).x2: GLine2(i).Y1 = GLine2(i - 1).Y2
            End If

            If i < 30 Then
                GLine1(i).x2 = lblX(i + 1).Left
                GLine2(i).x2 = lblX(i + 1).Left
            Else
                GLine1(i).x2 = XRight
                GLine2(i).x2 = XRight
            End If
            GLine1(i).Y2 = XTop - Percentage
            GLine2(i).Y2 = XTop - Percentage2
            '****************************************************************
            
            '************** Add Total Daily & Result ************************
            If CInt(lblX(i)) = MaxDailyDt Then
                lblN = lblBarVal1(i)
                lblN.top = GLine1(i).Y2 - 100
                lblN.Left = GLine1(i).x2 + 50
            ElseIf CInt(lblX(i)) > MaxDailyDt Then
                GLine1(i).Visible = False: lblBarVal1(i) = 0
            End If
            
            If CInt(lblX(i)) = MaxResultDt Then
                LblN2 = lblBarVal2(i)
                LblN2.top = GLine2(i).Y2 - 100
                LblN2.Left = GLine2(i).x2 + 50
                    
                '*** Set Total Daily sesuai Tgl Result terakhir
                If MaxDailyDt > MaxResultDt Then
                    lblN3 = lblBarVal1(i)
                    lblN3.top = GLine1(i).Y2
                    lblN3.Left = GLine1(i).x2 + 150
                Else
                    lblN3.Visible = False
                End If
            ElseIf CInt(lblX(i)) > MaxResultDt Then
                GLine2(i).Visible = False: lblBarVal2(i) = 0
            End If
            '****************************************************************
        Else

            lblX(i) = ""
            lblBarVal1(i) = 0: lblBarVal2(i) = 0
            GLine1(i).Visible = False: GLine2(i).Visible = False
        End If
    Next i
End Sub

Private Sub cmdReport_Click()
On Error GoTo HandleErr

    Label1(0).Visible = False: cmdReport.Visible = False: LinePrt.Visible = False
    StPrint = False
    
    frmEffPrintGraph.Orient = 2
    frmEffPrintGraph.Show 1
    
    If StPrint Then Me.PrintForm
    
HandleErr:
    If err.Description = "Cancel was selected." Or err.Description = "Printer error" _
        Or err.Description = "Can't print form image to this type of printer" Then
    ElseIf err.Description <> "" Then
        MsgBox err.Description, vbCritical, "Error"
    End If
    
    Label1(0).Visible = True: cmdReport.Visible = True: LinePrt.Visible = True
End Sub

