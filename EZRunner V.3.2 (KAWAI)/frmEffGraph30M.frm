VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmEffGraph30M 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Graph"
   ClientHeight    =   10365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13725
   FillColor       =   &H00FFC0C0&
   Icon            =   "frmEffGraph30M.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   10365
   ScaleWidth      =   13725
   StartUpPosition =   2  'CenterScreen
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
      Picture         =   "frmEffGraph30M.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   90
      Width           =   435
   End
   Begin MSComDlg.CommonDialog cdl 
      Left            =   1470
      Top             =   30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblY2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   810
      TabIndex        =   87
      Top             =   8460
      Width           =   660
   End
   Begin VB.Label lblY2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   810
      TabIndex        =   86
      Top             =   8205
      Width           =   660
   End
   Begin VB.Line LineValue2 
      Index           =   9
      X1              =   1500
      X2              =   1590
      Y1              =   8460
      Y2              =   8460
   End
   Begin VB.Line LineValue2 
      Index           =   8
      X1              =   1500
      X2              =   1590
      Y1              =   8205
      Y2              =   8205
   End
   Begin VB.Label lblY 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   810
      TabIndex        =   85
      Top             =   3210
      Width           =   660
   End
   Begin VB.Label lblY 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   810
      TabIndex        =   84
      Top             =   3465
      Width           =   660
   End
   Begin VB.Line LineValue 
      Index           =   9
      X1              =   1500
      X2              =   1590
      Y1              =   3210
      Y2              =   3210
   End
   Begin VB.Line LineValue 
      Index           =   8
      X1              =   1500
      X2              =   1590
      Y1              =   3465
      Y2              =   3465
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
      Left            =   6645
      TabIndex        =   83
      Top             =   1200
      Width           =   645
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Left            =   810
      Top             =   9750
      Width           =   705
   End
   Begin VB.Line Line1 
      X1              =   810
      X2              =   810
      Y1              =   1965
      Y2              =   9750
   End
   Begin VB.Line LineValue2 
      Index           =   7
      X1              =   1500
      X2              =   1590
      Y1              =   7965
      Y2              =   7965
   End
   Begin VB.Shape Shape2 
      Height          =   8610
      Left            =   510
      Top             =   1605
      Width           =   12915
   End
   Begin VB.Line Line2 
      X1              =   810
      X2              =   1530
      Y1              =   5715
      Y2              =   5715
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   1500
      TabIndex        =   82
      Top             =   9750
      Width           =   375
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblBarVal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   30
      Left            =   11580
      TabIndex        =   81
      Top             =   1680
      Width           =   150
   End
   Begin VB.Label lblBarVal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   29
      Left            =   11250
      TabIndex        =   80
      Top             =   1680
      Width           =   150
   End
   Begin VB.Label lblBarVal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   28
      Left            =   10920
      TabIndex        =   79
      Top             =   1680
      Width           =   150
   End
   Begin VB.Label lblBarVal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   27
      Left            =   10590
      TabIndex        =   78
      Top             =   1680
      Width           =   150
   End
   Begin VB.Label lblBarVal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   26
      Left            =   10260
      TabIndex        =   77
      Top             =   1680
      Width           =   150
   End
   Begin VB.Label lblBarVal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   25
      Left            =   9930
      TabIndex        =   76
      Top             =   1680
      Width           =   150
   End
   Begin VB.Label lblBarVal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   24
      Left            =   9600
      TabIndex        =   75
      Top             =   1680
      Width           =   150
   End
   Begin VB.Label lblBarVal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   23
      Left            =   9270
      TabIndex        =   74
      Top             =   1680
      Width           =   150
   End
   Begin VB.Label lblBarVal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   22
      Left            =   8940
      TabIndex        =   73
      Top             =   1680
      Width           =   150
   End
   Begin VB.Label lblBarVal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   21
      Left            =   8610
      TabIndex        =   72
      Top             =   1680
      Width           =   150
   End
   Begin VB.Label lblBarVal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   20
      Left            =   8280
      TabIndex        =   71
      Top             =   1680
      Width           =   150
   End
   Begin VB.Label lblBarVal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   19
      Left            =   7950
      TabIndex        =   70
      Top             =   1680
      Width           =   150
   End
   Begin VB.Label lblBarVal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   18
      Left            =   7620
      TabIndex        =   69
      Top             =   1680
      Width           =   150
   End
   Begin VB.Label lblBarVal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   17
      Left            =   7290
      TabIndex        =   68
      Top             =   1680
      Width           =   150
   End
   Begin VB.Label lblBarVal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   16
      Left            =   6960
      TabIndex        =   67
      Top             =   1680
      Width           =   150
   End
   Begin VB.Label lblBarVal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   15
      Left            =   6630
      TabIndex        =   66
      Top             =   1680
      Width           =   150
   End
   Begin VB.Label lblBarVal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   14
      Left            =   6300
      TabIndex        =   65
      Top             =   1680
      Width           =   150
   End
   Begin VB.Label lblBarVal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   13
      Left            =   5970
      TabIndex        =   64
      Top             =   1680
      Width           =   150
   End
   Begin VB.Label lblBarVal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   12
      Left            =   5640
      TabIndex        =   63
      Top             =   1680
      Width           =   150
   End
   Begin VB.Label lblBarVal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   11
      Left            =   5310
      TabIndex        =   62
      Top             =   1680
      Width           =   150
   End
   Begin VB.Label lblBarVal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   10
      Left            =   4980
      TabIndex        =   61
      Top             =   1680
      Width           =   150
   End
   Begin VB.Label lblBarVal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   9
      Left            =   4650
      TabIndex        =   60
      Top             =   1680
      Width           =   150
   End
   Begin VB.Label lblBarVal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   8
      Left            =   4320
      TabIndex        =   59
      Top             =   1680
      Width           =   150
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   24
      Left            =   10500
      TabIndex        =   58
      Top             =   9750
      Width           =   375
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   25
      Left            =   10875
      TabIndex        =   57
      Top             =   9750
      Width           =   375
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   26
      Left            =   11250
      TabIndex        =   56
      Top             =   9750
      Width           =   375
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   27
      Left            =   11625
      TabIndex        =   55
      Top             =   9750
      Width           =   375
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   28
      Left            =   12000
      TabIndex        =   54
      Top             =   9750
      Width           =   375
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   29
      Left            =   12375
      TabIndex        =   53
      Top             =   9750
      Width           =   375
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   30
      Left            =   12750
      TabIndex        =   52
      Top             =   9750
      Width           =   375
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   16
      Left            =   7500
      TabIndex        =   51
      Top             =   9750
      Width           =   375
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   17
      Left            =   7875
      TabIndex        =   50
      Top             =   9750
      Width           =   375
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   18
      Left            =   8250
      TabIndex        =   49
      Top             =   9750
      Width           =   375
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   19
      Left            =   8625
      TabIndex        =   48
      Top             =   9750
      Width           =   375
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   20
      Left            =   9000
      TabIndex        =   47
      Top             =   9750
      Width           =   375
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   21
      Left            =   9375
      TabIndex        =   46
      Top             =   9750
      Width           =   375
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   22
      Left            =   9750
      TabIndex        =   45
      Top             =   9750
      Width           =   375
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   23
      Left            =   10125
      TabIndex        =   44
      Top             =   9750
      Width           =   375
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   4500
      TabIndex        =   43
      Top             =   9750
      Width           =   375
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   4875
      TabIndex        =   42
      Top             =   9750
      Width           =   375
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   5250
      TabIndex        =   41
      Top             =   9750
      Width           =   375
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   5625
      TabIndex        =   40
      Top             =   9750
      Width           =   375
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   12
      Left            =   6000
      TabIndex        =   39
      Top             =   9750
      Width           =   375
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   13
      Left            =   6375
      TabIndex        =   38
      Top             =   9750
      Width           =   375
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   14
      Left            =   6750
      TabIndex        =   37
      Top             =   9750
      Width           =   375
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   15
      Left            =   7125
      TabIndex        =   36
      Top             =   9750
      Width           =   375
   End
   Begin VB.Line LineValue2 
      Index           =   6
      X1              =   1500
      X2              =   1590
      Y1              =   7710
      Y2              =   7710
   End
   Begin VB.Line LineValue2 
      Index           =   5
      X1              =   1500
      X2              =   1590
      Y1              =   7455
      Y2              =   7455
   End
   Begin VB.Line LineValue2 
      Index           =   4
      X1              =   1500
      X2              =   1590
      Y1              =   7215
      Y2              =   7215
   End
   Begin VB.Line LineValue2 
      Index           =   3
      X1              =   1500
      X2              =   1590
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Line LineValue2 
      Index           =   2
      X1              =   1500
      X2              =   1590
      Y1              =   6705
      Y2              =   6705
   End
   Begin VB.Line LineValue2 
      Index           =   1
      X1              =   1500
      X2              =   1590
      Y1              =   6465
      Y2              =   6465
   End
   Begin VB.Line LineValue2 
      Index           =   0
      X1              =   1500
      X2              =   1590
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Label lblY2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   810
      TabIndex        =   35
      Top             =   7965
      Width           =   660
   End
   Begin VB.Label lblY2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   810
      TabIndex        =   34
      Top             =   7710
      Width           =   660
   End
   Begin VB.Label lblY2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   810
      TabIndex        =   33
      Top             =   7455
      Width           =   660
   End
   Begin VB.Label lblY2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   810
      TabIndex        =   32
      Top             =   7215
      Width           =   660
   End
   Begin VB.Label lblY2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   810
      TabIndex        =   31
      Top             =   6960
      Width           =   660
   End
   Begin VB.Label lblY2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   810
      TabIndex        =   30
      Top             =   6705
      Width           =   660
   End
   Begin VB.Label lblY2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   810
      TabIndex        =   29
      Top             =   6465
      Width           =   660
   End
   Begin VB.Label lblY2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   810
      TabIndex        =   28
      Top             =   6210
      Width           =   660
   End
   Begin VB.Line LinePrt 
      BorderWidth     =   2
      X1              =   0
      X2              =   13710
      Y1              =   540
      Y2              =   540
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
      Left            =   6420
      TabIndex        =   27
      Top             =   990
      Width           =   1095
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Production Schedule/Result Difference Control "
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
      Left            =   3750
      TabIndex        =   26
      Top             =   660
      Width           =   6435
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
      TabIndex        =   25
      Top             =   180
      Width           =   390
   End
   Begin VB.Label lblValue0 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   810
      TabIndex        =   23
      Top             =   5730
      Width           =   660
   End
   Begin VB.Label lblBarVal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   6
      Left            =   3660
      TabIndex        =   21
      Top             =   1680
      Width           =   150
   End
   Begin VB.Label lblBarVal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   5
      Left            =   3330
      TabIndex        =   20
      Top             =   1680
      Width           =   150
   End
   Begin VB.Label lblBarVal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   4
      Left            =   3000
      TabIndex        =   19
      Top             =   1680
      Width           =   150
   End
   Begin VB.Label lblBarVal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   3
      Left            =   2670
      TabIndex        =   18
      Top             =   1680
      Width           =   150
   End
   Begin VB.Label lblBarVal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   2
      Left            =   2340
      TabIndex        =   17
      Top             =   1680
      Width           =   150
   End
   Begin VB.Label lblBarVal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   1
      Left            =   2010
      TabIndex        =   16
      Top             =   1680
      Width           =   150
   End
   Begin VB.Label lblBarVal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   0
      Left            =   1680
      TabIndex        =   15
      Top             =   1680
      Width           =   150
   End
   Begin VB.Label lblY 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   810
      TabIndex        =   14
      Top             =   3705
      Width           =   660
   End
   Begin VB.Label lblY 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   810
      TabIndex        =   13
      Top             =   3960
      Width           =   660
   End
   Begin VB.Label lblY 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   810
      TabIndex        =   12
      Top             =   4215
      Width           =   660
   End
   Begin VB.Label lblY 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   810
      TabIndex        =   11
      Top             =   4455
      Width           =   660
   End
   Begin VB.Label lblY 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   810
      TabIndex        =   10
      Top             =   4710
      Width           =   660
   End
   Begin VB.Label lblY 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   810
      TabIndex        =   9
      Top             =   4965
      Width           =   660
   End
   Begin VB.Label lblY 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   810
      TabIndex        =   8
      Top             =   5205
      Width           =   660
   End
   Begin VB.Label lblY 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label 1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   810
      TabIndex        =   7
      Top             =   5460
      Width           =   660
   End
   Begin VB.Line LineValue 
      Index           =   6
      X1              =   1500
      X2              =   1590
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line LineValue 
      Index           =   3
      X1              =   1500
      X2              =   1590
      Y1              =   4710
      Y2              =   4710
   End
   Begin VB.Line LineValue 
      Index           =   5
      X1              =   1500
      X2              =   1590
      Y1              =   4215
      Y2              =   4215
   End
   Begin VB.Line LineValue 
      Index           =   1
      X1              =   1500
      X2              =   1590
      Y1              =   5205
      Y2              =   5205
   End
   Begin VB.Line LineValue 
      Index           =   2
      X1              =   1500
      X2              =   1590
      Y1              =   4965
      Y2              =   4965
   End
   Begin VB.Line LineValue 
      Index           =   0
      X1              =   1500
      X2              =   1590
      Y1              =   5460
      Y2              =   5460
   End
   Begin VB.Line LineValue 
      Index           =   4
      X1              =   1500
      X2              =   1590
      Y1              =   4455
      Y2              =   4455
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   4125
      TabIndex        =   6
      Top             =   9750
      Width           =   375
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   3750
      TabIndex        =   5
      Top             =   9750
      Width           =   375
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   3375
      TabIndex        =   4
      Top             =   9750
      Width           =   375
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   3000
      TabIndex        =   3
      Top             =   9750
      Width           =   375
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   2625
      TabIndex        =   2
      Top             =   9750
      Width           =   375
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   2250
      TabIndex        =   1
      Top             =   9750
      Width           =   375
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   1875
      TabIndex        =   0
      Top             =   9750
      Width           =   375
   End
   Begin VB.Shape sBar 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  'Solid
      Height          =   3000
      Index           =   29
      Left            =   12375
      Top             =   2715
      Width           =   345
   End
   Begin VB.Shape sBar 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  'Solid
      Height          =   3000
      Index           =   28
      Left            =   12000
      Top             =   2715
      Width           =   345
   End
   Begin VB.Shape sBar 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  'Solid
      Height          =   3000
      Index           =   27
      Left            =   11625
      Top             =   2715
      Width           =   345
   End
   Begin VB.Shape sBar 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  'Solid
      Height          =   3000
      Index           =   26
      Left            =   11250
      Top             =   2715
      Width           =   345
   End
   Begin VB.Shape sBar 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  'Solid
      Height          =   3000
      Index           =   25
      Left            =   10875
      Top             =   2715
      Width           =   345
   End
   Begin VB.Shape sBar 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  'Solid
      Height          =   3000
      Index           =   24
      Left            =   10500
      Top             =   2715
      Width           =   345
   End
   Begin VB.Shape sBar 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  'Solid
      Height          =   3000
      Index           =   23
      Left            =   10125
      Top             =   2715
      Width           =   345
   End
   Begin VB.Shape sBar 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  'Solid
      Height          =   3000
      Index           =   22
      Left            =   9750
      Top             =   2715
      Width           =   345
   End
   Begin VB.Shape sBar 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  'Solid
      Height          =   3000
      Index           =   21
      Left            =   9375
      Top             =   2715
      Width           =   345
   End
   Begin VB.Shape sBar 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  'Solid
      Height          =   3000
      Index           =   20
      Left            =   9000
      Top             =   2715
      Width           =   345
   End
   Begin VB.Shape sBar 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  'Solid
      Height          =   3000
      Index           =   19
      Left            =   8625
      Top             =   2715
      Width           =   345
   End
   Begin VB.Shape sBar 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  'Solid
      Height          =   3000
      Index           =   18
      Left            =   8250
      Top             =   2715
      Width           =   345
   End
   Begin VB.Shape sBar 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  'Solid
      Height          =   3000
      Index           =   17
      Left            =   7875
      Top             =   2715
      Width           =   345
   End
   Begin VB.Shape sBar 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  'Solid
      Height          =   3000
      Index           =   16
      Left            =   7500
      Top             =   2715
      Width           =   345
   End
   Begin VB.Shape sBar 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  'Solid
      Height          =   3000
      Index           =   15
      Left            =   7125
      Top             =   2715
      Width           =   345
   End
   Begin VB.Shape sBar 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  'Solid
      Height          =   3000
      Index           =   14
      Left            =   6750
      Top             =   2715
      Width           =   345
   End
   Begin VB.Shape sBar 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  'Solid
      Height          =   3000
      Index           =   13
      Left            =   6375
      Top             =   2715
      Width           =   345
   End
   Begin VB.Shape sBar 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  'Solid
      Height          =   3000
      Index           =   12
      Left            =   6000
      Top             =   2715
      Width           =   345
   End
   Begin VB.Shape sBar 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  'Solid
      Height          =   3000
      Index           =   11
      Left            =   5625
      Top             =   2715
      Width           =   345
   End
   Begin VB.Shape sBar 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  'Solid
      Height          =   3000
      Index           =   10
      Left            =   5250
      Top             =   2715
      Width           =   345
   End
   Begin VB.Shape sBar 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  'Solid
      Height          =   3000
      Index           =   9
      Left            =   4875
      Top             =   2715
      Width           =   345
   End
   Begin VB.Shape sBar 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  'Solid
      Height          =   3000
      Index           =   8
      Left            =   4500
      Top             =   2715
      Width           =   345
   End
   Begin VB.Line LineYValue 
      X1              =   1500
      X2              =   1500
      Y1              =   1965
      Y2              =   9750
   End
   Begin VB.Line LineX 
      X1              =   1500
      X2              =   13100
      Y1              =   5715
      Y2              =   5715
   End
   Begin VB.Shape sBar 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  'Solid
      Height          =   3000
      Index           =   1
      Left            =   1875
      Top             =   2715
      Width           =   345
   End
   Begin VB.Shape sBar 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  'Solid
      Height          =   3000
      Index           =   2
      Left            =   2250
      Top             =   2715
      Width           =   345
   End
   Begin VB.Shape sBar 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  'Solid
      Height          =   3000
      Index           =   3
      Left            =   2625
      Top             =   2715
      Width           =   345
   End
   Begin VB.Shape sBar 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  'Solid
      Height          =   3000
      Index           =   4
      Left            =   3000
      Top             =   2715
      Width           =   345
   End
   Begin VB.Shape sBar 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  'Solid
      Height          =   3000
      Index           =   5
      Left            =   3375
      Top             =   2715
      Width           =   345
   End
   Begin VB.Shape sBar 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  'Solid
      Height          =   3000
      Index           =   6
      Left            =   3750
      Top             =   2715
      Width           =   345
   End
   Begin VB.Shape sBar 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  'Solid
      Height          =   3000
      Index           =   30
      Left            =   12750
      Top             =   2715
      Width           =   345
   End
   Begin VB.Label lblBarVal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   7
      Left            =   3990
      TabIndex        =   22
      Top             =   1680
      Width           =   150
   End
   Begin VB.Shape sBar 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  'Solid
      Height          =   3000
      Index           =   7
      Left            =   4125
      Top             =   2715
      Width           =   345
   End
   Begin VB.Line LineValue 
      Index           =   7
      X1              =   1500
      X2              =   1590
      Y1              =   3705
      Y2              =   3705
   End
   Begin VB.Shape sBar 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  'Solid
      Height          =   3000
      Index           =   0
      Left            =   1500
      Top             =   2715
      Width           =   345
   End
End
Attribute VB_Name = "frmEffGraph30M"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i  As Integer

Dim XLeft As Double, XRight As Double, XTop As Double, YTop As Double
Dim RangeY As Double, MaxY As Double, totLineY As Integer, distY As Double
Dim Percentage As Double

Public jmlBar As Double, MaxQty As Double, MinQty As Double

Public Sub viewGraph()
    XLeft = LineX.X1: XRight = LineX.x2: XTop = LineX.Y1: YTop = LineYValue.Y1
    RangeY = setRange(MaxQty)
    MaxY = RoundUp(MaxQty / RangeY) * RangeY
    totLineY = MaxY / RangeY
    distY = ((XTop - YTop) / MaxY)
    
    Percentage = 0
    For i = 0 To 30
        If i < 10 Then
            If i < totLineY Then
                lblY(i).Visible = True: lblY2(i).Visible = True
                LineValue(i).Visible = True: LineValue2(i).Visible = True
                
                lblY(i) = Format((MaxY / totLineY) * (i + 1), gs_formatQty)
                lblY2(i) = Format(-(MaxY / totLineY) * (i + 1), gs_formatQty)
                lblY(i).top = XTop - (lblY(i) * distY)
                lblY2(i).top = XTop + (lblY(i) * distY)
                LineValue(i).Y1 = lblY(i).top: LineValue(i).Y2 = LineValue(i).Y1
                LineValue2(i).Y1 = lblY2(i).top: LineValue2(i).Y2 = LineValue2(i).Y1
            Else
                lblY(i).Visible = False: lblY2(i).Visible = False
                LineValue(i).Visible = False: LineValue2(i).Visible = False
            End If
        End If
        
        If i < jmlBar Then
            sBar(i).Visible = True: lblBarVal(i).Visible = True
            
            '*********** Set Top & Height *****************
            Percentage = CDbl(Abs(lblBarVal(i))) * distY
            
            If CDbl(lblBarVal(i)) >= 0 Then
                sBar(i).top = XTop - Percentage
                sBar(i).Height = XTop - sBar(i).top
                lblBarVal(i).top = sBar(i).top - 200
            Else
                sBar(i).top = XTop
                sBar(i).Height = Percentage
                lblBarVal(i).top = sBar(i).top + sBar(i).Height + 100
            End If
            sBar(i).FillColor = Split(ColGraph, ",")(CInt(lblBarVal(i).Tag))
            lblBarVal(i).Left = sBar(i).Left + (sBar(i).Width - lblBarVal(i).Width) / 2
            
        Else
            lblBarVal(i) = 0: lblX(i) = ""
            sBar(i).Visible = False: lblBarVal(i).Visible = False
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

