VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmEffGraph8 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Graph"
   ClientHeight    =   9900
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7980
   FillColor       =   &H00FFC0C0&
   Icon            =   "frmEffGraph8.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   9900
   ScaleWidth      =   7980
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSign 
      BackColor       =   &H00C0C0C0&
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   4470
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   90
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton cmdSign 
      BackColor       =   &H00C0C0C0&
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   90
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton cmdSign 
      BackColor       =   &H00C0C0C0&
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3450
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   90
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton cmdSign 
      BackColor       =   &H00C0C0C0&
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2940
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   90
      Visible         =   0   'False
      Width           =   435
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
      Left            =   960
      Picture         =   "frmEffGraph8.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   90
      Width           =   435
   End
   Begin MSComDlg.CommonDialog cdl 
      Left            =   1650
      Top             =   30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Line LinePercent 
      Index           =   9
      X1              =   6360
      X2              =   6270
      Y1              =   2610
      Y2              =   2610
   End
   Begin VB.Line LinePercent 
      Index           =   8
      X1              =   6360
      X2              =   6270
      Y1              =   3088
      Y2              =   3088
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
      Left            =   720
      TabIndex        =   73
      Top             =   3690
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
      Left            =   720
      TabIndex        =   72
      Top             =   4020
      Width           =   750
   End
   Begin VB.Line LineValue 
      Index           =   9
      X1              =   1560
      X2              =   1650
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line LineValue 
      Index           =   8
      X1              =   1560
      X2              =   1650
      Y1              =   3960
      Y2              =   3960
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
      Index           =   3
      Left            =   3698
      TabIndex        =   71
      Top             =   1590
      Width           =   645
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Index           =   31
      Left            =   6210
      TabIndex        =   70
      Top             =   7470
      Width           =   105
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Index           =   30
      Left            =   6075
      TabIndex        =   69
      Top             =   7470
      Width           =   105
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Index           =   29
      Left            =   5940
      TabIndex        =   68
      Top             =   7470
      Width           =   105
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Index           =   28
      Left            =   5820
      TabIndex        =   67
      Top             =   7470
      Width           =   105
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Index           =   27
      Left            =   5610
      TabIndex        =   66
      Top             =   7470
      Width           =   105
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Index           =   26
      Left            =   5475
      TabIndex        =   65
      Top             =   7470
      Width           =   105
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Index           =   25
      Left            =   5340
      TabIndex        =   64
      Top             =   7470
      Width           =   105
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Index           =   24
      Left            =   5220
      TabIndex        =   63
      Top             =   7470
      Width           =   105
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Index           =   23
      Left            =   5010
      TabIndex        =   62
      Top             =   7470
      Width           =   105
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Index           =   22
      Left            =   4875
      TabIndex        =   61
      Top             =   7470
      Width           =   105
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Index           =   21
      Left            =   4740
      TabIndex        =   60
      Top             =   7470
      Width           =   105
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Index           =   20
      Left            =   4620
      TabIndex        =   59
      Top             =   7470
      Width           =   105
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Index           =   19
      Left            =   4410
      TabIndex        =   58
      Top             =   7470
      Width           =   105
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Index           =   18
      Left            =   4275
      TabIndex        =   57
      Top             =   7470
      Width           =   105
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Index           =   17
      Left            =   4140
      TabIndex        =   56
      Top             =   7470
      Width           =   105
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Index           =   16
      Left            =   4020
      TabIndex        =   55
      Top             =   7470
      Width           =   105
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblRecord 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Record  of "
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
      Left            =   5070
      TabIndex        =   54
      Top             =   180
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Material"
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
      Left            =   3668
      TabIndex        =   53
      Top             =   1320
      Width           =   705
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Index           =   15
      Left            =   3810
      TabIndex        =   52
      Top             =   7470
      Width           =   105
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Index           =   14
      Left            =   3675
      TabIndex        =   51
      Top             =   7470
      Width           =   105
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Index           =   13
      Left            =   3540
      TabIndex        =   50
      Top             =   7470
      Width           =   105
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Index           =   12
      Left            =   3420
      TabIndex        =   49
      Top             =   7470
      Width           =   105
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Index           =   11
      Left            =   3210
      TabIndex        =   48
      Top             =   7470
      Width           =   105
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Index           =   10
      Left            =   3075
      TabIndex        =   47
      Top             =   7470
      Width           =   105
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Index           =   9
      Left            =   2940
      TabIndex        =   46
      Top             =   7470
      Width           =   105
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Index           =   8
      Left            =   2820
      TabIndex        =   45
      Top             =   7470
      Width           =   105
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Index           =   7
      Left            =   2610
      TabIndex        =   44
      Top             =   7470
      Width           =   105
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Index           =   6
      Left            =   2475
      TabIndex        =   43
      Top             =   7470
      Width           =   105
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Index           =   5
      Left            =   2340
      TabIndex        =   42
      Top             =   7470
      Width           =   105
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Index           =   4
      Left            =   2220
      TabIndex        =   41
      Top             =   7470
      Width           =   105
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Index           =   3
      Left            =   2010
      TabIndex        =   40
      Top             =   7470
      Width           =   105
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Index           =   2
      Left            =   1875
      TabIndex        =   39
      Top             =   7470
      Width           =   105
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Index           =   1
      Left            =   1740
      TabIndex        =   38
      Top             =   7470
      Width           =   105
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Index           =   0
      Left            =   1620
      TabIndex        =   37
      Top             =   7470
      Width           =   105
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblLineDot 
      BackStyle       =   0  'Transparent
      Caption         =   "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -"
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
      Height          =   210
      Index           =   7
      Left            =   1560
      TabIndex        =   36
      Top             =   3390
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Label lblLineDot 
      BackStyle       =   0  'Transparent
      Caption         =   "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -"
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
      Height          =   210
      Index           =   6
      Left            =   1560
      TabIndex        =   35
      Top             =   3210
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Label lblLineDot 
      BackStyle       =   0  'Transparent
      Caption         =   "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -"
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
      Height          =   210
      Index           =   5
      Left            =   1560
      TabIndex        =   34
      Top             =   3030
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Label lblLineDot 
      BackStyle       =   0  'Transparent
      Caption         =   "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -"
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
      Height          =   210
      Index           =   4
      Left            =   1560
      TabIndex        =   33
      Top             =   3660
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Label lblLineDot 
      BackStyle       =   0  'Transparent
      Caption         =   "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -"
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
      Height          =   210
      Index           =   3
      Left            =   1560
      TabIndex        =   32
      Top             =   2910
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Label lblLineDot 
      BackStyle       =   0  'Transparent
      Caption         =   "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -"
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
      Height          =   210
      Index           =   2
      Left            =   1560
      TabIndex        =   31
      Top             =   2760
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Label lblLineDot 
      BackStyle       =   0  'Transparent
      Caption         =   "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -"
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
      Height          =   210
      Index           =   1
      Left            =   1560
      TabIndex        =   30
      Top             =   2610
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Label lblLineDot 
      BackStyle       =   0  'Transparent
      Caption         =   "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -"
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
      Height          =   210
      Index           =   0
      Left            =   1560
      TabIndex        =   29
      Top             =   3540
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Shape Shape1 
      Height          =   7605
      Left            =   450
      Top             =   1980
      Width           =   7140
   End
   Begin VB.Line LinePrt 
      BorderWidth     =   2
      X1              =   0
      X2              =   8100
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Machine No"
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
      Left            =   3518
      TabIndex        =   28
      Top             =   1080
      Width           =   1005
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bad Material Quantity Control"
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
      Left            =   1943
      TabIndex        =   27
      Top             =   690
      Width           =   4155
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
      Left            =   450
      TabIndex        =   26
      Top             =   180
      Width           =   390
      WordWrap        =   -1  'True
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
      Left            =   1575
      TabIndex        =   24
      Top             =   2040
      Width           =   570
   End
   Begin VB.Label lblPercent50 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "50 %"
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
      Left            =   6450
      TabIndex        =   23
      Top             =   4920
      Width           =   450
   End
   Begin VB.Label lblPercent100 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100 %"
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
      Left            =   6450
      TabIndex        =   22
      Top             =   2520
      Width           =   555
   End
   Begin VB.Label lblPercent0 
      AutoSize        =   -1  'True
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
      Height          =   195
      Left            =   6450
      TabIndex        =   21
      Top             =   7320
      Width           =   105
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
      Left            =   1125
      TabIndex        =   20
      Top             =   7290
      Width           =   345
   End
   Begin VB.Label lblBarVal 
      Alignment       =   2  'Center
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
      Left            =   5760
      TabIndex        =   19
      Top             =   2340
      Width           =   600
   End
   Begin VB.Label lblBarVal 
      Alignment       =   2  'Center
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
      Left            =   5160
      TabIndex        =   18
      Top             =   2340
      Width           =   600
   End
   Begin VB.Label lblBarVal 
      Alignment       =   2  'Center
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
      Left            =   4560
      TabIndex        =   17
      Top             =   2340
      Width           =   600
   End
   Begin VB.Label lblBarVal 
      Alignment       =   2  'Center
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
      Left            =   3960
      TabIndex        =   16
      Top             =   2340
      Width           =   600
   End
   Begin VB.Label lblBarVal 
      Alignment       =   2  'Center
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
      Left            =   3360
      TabIndex        =   15
      Top             =   2340
      Width           =   600
   End
   Begin VB.Label lblBarVal 
      Alignment       =   2  'Center
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
      Left            =   2760
      TabIndex        =   14
      Top             =   2340
      Width           =   600
   End
   Begin VB.Label lblBarVal 
      Alignment       =   2  'Center
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
      Left            =   2160
      TabIndex        =   13
      Top             =   2340
      Width           =   600
   End
   Begin VB.Label lblBarVal 
      Alignment       =   2  'Center
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
      Left            =   1560
      TabIndex        =   12
      Top             =   2340
      Width           =   600
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
      Left            =   720
      TabIndex        =   11
      Top             =   4410
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
      Left            =   720
      TabIndex        =   10
      Top             =   4710
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
      Left            =   720
      TabIndex        =   9
      Top             =   5010
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
      Left            =   720
      TabIndex        =   8
      Top             =   5370
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
      Left            =   720
      TabIndex        =   7
      Top             =   5730
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
      Left            =   720
      TabIndex        =   6
      Top             =   6000
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
      Left            =   720
      TabIndex        =   5
      Top             =   6330
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
      Left            =   720
      TabIndex        =   4
      Top             =   6690
      Width           =   750
   End
   Begin VB.Line LinePercent 
      Index           =   0
      X1              =   6360
      X2              =   6270
      Y1              =   6916
      Y2              =   6916
   End
   Begin VB.Line LinePercent 
      Index           =   5
      X1              =   6360
      X2              =   6270
      Y1              =   4524
      Y2              =   4524
   End
   Begin VB.Line LinePercent 
      Index           =   4
      X1              =   6360
      X2              =   6270
      Y1              =   5002
      Y2              =   5002
   End
   Begin VB.Line LinePercent 
      Index           =   3
      X1              =   6360
      X2              =   6270
      Y1              =   5481
      Y2              =   5481
   End
   Begin VB.Line LinePercent 
      Index           =   2
      X1              =   6360
      X2              =   6270
      Y1              =   5960
      Y2              =   5960
   End
   Begin VB.Line LinePercent 
      Index           =   1
      X1              =   6360
      X2              =   6270
      Y1              =   6438
      Y2              =   6438
   End
   Begin VB.Line LinePercent 
      Index           =   6
      X1              =   6360
      X2              =   6270
      Y1              =   4046
      Y2              =   4046
   End
   Begin VB.Line LinePercent 
      Index           =   7
      X1              =   6360
      X2              =   6270
      Y1              =   3567
      Y2              =   3567
   End
   Begin VB.Line GLine 
      Index           =   7
      X1              =   5730
      X2              =   6330
      Y1              =   7400
      Y2              =   5010
   End
   Begin VB.Line GLine 
      Index           =   6
      X1              =   5160
      X2              =   5730
      Y1              =   7400
      Y2              =   5010
   End
   Begin VB.Line GLine 
      Index           =   5
      X1              =   4560
      X2              =   5160
      Y1              =   7400
      Y2              =   5010
   End
   Begin VB.Line GLine 
      Index           =   4
      X1              =   3960
      X2              =   4560
      Y1              =   7400
      Y2              =   5010
   End
   Begin VB.Line GLine 
      Index           =   3
      X1              =   3360
      X2              =   3960
      Y1              =   7400
      Y2              =   5010
   End
   Begin VB.Line GLine 
      Index           =   2
      X1              =   2760
      X2              =   3360
      Y1              =   7400
      Y2              =   5010
   End
   Begin VB.Line GLine 
      Index           =   1
      X1              =   2160
      X2              =   2760
      Y1              =   7400
      Y2              =   5010
   End
   Begin VB.Line LineValue 
      Index           =   7
      X1              =   1560
      X2              =   1650
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line LineValue 
      Index           =   6
      X1              =   1560
      X2              =   1650
      Y1              =   4260
      Y2              =   4260
   End
   Begin VB.Line GLine 
      Index           =   0
      X1              =   1560
      X2              =   2130
      Y1              =   7395
      Y2              =   5040
   End
   Begin VB.Line LineYValue 
      X1              =   1560
      X2              =   1560
      Y1              =   2610
      Y2              =   7410
   End
   Begin VB.Line LineX 
      X1              =   1560
      X2              =   6360
      Y1              =   7395
      Y2              =   7395
   End
   Begin VB.Shape sBar 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  'Solid
      Height          =   2400
      Index           =   1
      Left            =   2160
      Top             =   5010
      Width           =   600
   End
   Begin VB.Shape sBar 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  'Solid
      Height          =   2400
      Index           =   2
      Left            =   2760
      Top             =   5010
      Width           =   600
   End
   Begin VB.Shape sBar 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  'Solid
      Height          =   2400
      Index           =   3
      Left            =   3360
      Top             =   5010
      Width           =   600
   End
   Begin VB.Shape sBar 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  'Solid
      Height          =   2400
      Index           =   4
      Left            =   3960
      Top             =   5010
      Width           =   600
   End
   Begin VB.Shape sBar 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  'Solid
      Height          =   2400
      Index           =   5
      Left            =   4560
      Top             =   5010
      Width           =   600
   End
   Begin VB.Shape sBar 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  'Solid
      Height          =   2400
      Index           =   6
      Left            =   5160
      Top             =   5010
      Width           =   600
   End
   Begin VB.Shape sBar 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  'Solid
      Height          =   2400
      Index           =   7
      Left            =   5760
      Top             =   5010
      Width           =   600
   End
   Begin VB.Line LineYPercent 
      X1              =   6360
      X2              =   6360
      Y1              =   2610
      Y2              =   7410
   End
   Begin VB.Line LineValue 
      Index           =   3
      X1              =   1560
      X2              =   1650
      Y1              =   5010
      Y2              =   5010
   End
   Begin VB.Line LineValue 
      Index           =   5
      X1              =   1560
      X2              =   1650
      Y1              =   4500
      Y2              =   4500
   End
   Begin VB.Line LineValue 
      Index           =   1
      X1              =   1560
      X2              =   1650
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LineValue 
      Index           =   2
      X1              =   1560
      X2              =   1650
      Y1              =   5610
      Y2              =   5610
   End
   Begin VB.Line LineValue 
      Index           =   0
      X1              =   1560
      X2              =   1650
      Y1              =   6810
      Y2              =   6810
   End
   Begin VB.Line LineValue 
      Index           =   4
      X1              =   1560
      X2              =   1650
      Y1              =   4770
      Y2              =   4770
   End
   Begin VB.Shape sBar 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  'Solid
      Height          =   2400
      Index           =   0
      Left            =   1560
      Top             =   5010
      Width           =   600
   End
End
Attribute VB_Name = "frmEffGraph8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i  As Integer, j As Integer

Dim XLeft As Double, XRight As Double, XTop As Double, YTop As Double
Dim RangeY As Double, MaxY As Double, totLineY As Integer, distY As Double
Dim Percentage As Double, PercentageLine As Double, MaxLen As Double

Public jmlBar As Double, MaxQty As Double
Public visibleSign As Boolean

Sub SetVisibleSign(stVisible As Boolean)
    cmdSign(0).Visible = stVisible: cmdSign(1).Visible = stVisible
    cmdSign(2).Visible = stVisible: cmdSign(3).Visible = stVisible: LblRecord.Visible = stVisible
End Sub

Public Sub viewGraph()
Dim indx As Integer
    
    Call SetVisibleSign(visibleSign)
    
    XLeft = LineX.X1: XRight = LineX.x2: XTop = LineX.Y1: YTop = LineYValue.Y1
    
    RangeY = setRange(MaxQty)
    MaxY = Fix(MaxQty / RangeY) * RangeY
    totLineY = MaxY / RangeY
    distY = ((XTop - YTop) / MaxQty)
  
    Percentage = 0: PercentageLine = 0: MaxLen = 0
    For i = 0 To 9
        '******************* Set Y *******************
        If i < totLineY Then
            lblY(i).Visible = True: LineValue(i).Visible = True
            
            lblY(i) = Format((MaxY / totLineY) * (i + 1), gs_formatQty)
            lblY(i).top = XTop - (lblY(i) * distY)
            LineValue(i).Y1 = lblY(i).top: LineValue(i).Y2 = LineValue(i).Y1
        Else
            lblY(i).Visible = False: LineValue(i).Visible = False
            
        End If
        
        '******************* Set X *******************
        If i <= 7 Then ' Set Y
            indx = (i * 4)
            lblX(indx).Visible = False: lblX(indx + 1).Visible = False
            lblX(indx + 2).Visible = False: lblX(indx + 3).Visible = False
            
            If i < jmlBar Then
                sBar(i).Visible = True: lblBarVal(i).Visible = True
                GLine(i).Visible = True: lblLineDot(i).Visible = True
        
                '*********** Set Top & Height *****************
                Percentage = CDbl(lblBarVal(i)) * distY
                PercentageLine = CDbl(lblBarVal(i).Tag) * distY
        
                sBar(i).top = XTop - Percentage
                sBar(i).Height = XTop - sBar(i).top
                lblBarVal(i).top = sBar(i).top - 250
        
                '*********** Set Left & Width *****************
                sBar(i).Width = (XRight - XLeft) / jmlBar
        
                If i = 0 Then
                    sBar(i).Left = XLeft
                    GLine(i).Y1 = XTop
                Else
                    sBar(i).Left = sBar(i - 1).Left + sBar(i).Width
                    GLine(i).Y1 = GLine(i - 1).Y2
                End If
        
                GLine(i).Y2 = XTop - PercentageLine
                GLine(i).X1 = sBar(i).Left
                GLine(i).x2 = sBar(i).Left + sBar(i).Width
        
                lblBarVal(i).Left = sBar(i).Left
                lblBarVal(i).Width = sBar(i).Width
        
                lblLineDot(i).top = GLine(i).Y2 - 70
                lblLineDot(i).Width = 4815
        
                '************* Set Wordwrap *************************
                Dim temp(3) As String, lenLbl As Double, totLbl As Integer, distAwal As Double
        
                lenLbl = Len(lblX(indx).Tag)
                If lenLbl > MaxLen Then MaxLen = lenLbl
                temp(0) = "": temp(1) = "": temp(2) = "": temp(3) = ""
        
                For j = 1 To lenLbl
                    If j <= 13 Then
                        temp(0) = temp(0) & Mid(lblX(indx).Tag, j, 1) & vbCrLf
                    ElseIf j <= 13 * 2 Then
                        temp(1) = temp(1) & Mid(lblX(indx).Tag, j, 1) & vbCrLf
                    ElseIf j <= 13 * 3 Then
                        temp(2) = temp(2) & Mid(lblX(indx).Tag, j, 1) & vbCrLf
                    ElseIf j <= 13 * 4 Then
                        temp(3) = temp(3) & Mid(lblX(indx).Tag, j, 1) & vbCrLf
                    End If
                Next j
        
                If temp(0) <> "" Then lblX(indx) = temp(0): lblX(indx).Visible = True: totLbl = 1
                If temp(1) <> "" Then lblX(indx + 1) = temp(1): lblX(indx + 1).Visible = True: totLbl = 2
                If temp(2) <> "" Then lblX(indx + 2) = temp(2): lblX(indx + 2).Visible = True: totLbl = 3
                If temp(3) <> "" Then lblX(indx + 3) = temp(3): lblX(indx + 3).Visible = True: totLbl = 4
        
                distAwal = (sBar(i).Width - (totLbl * 100) - (totLbl * 30)) / 2
                lblX(indx).Left = sBar(i).Left + distAwal
                lblX(indx + 1).Left = lblX(indx).Left + 130
                lblX(indx + 2).Left = lblX(indx + 1).Left + 130
                lblX(indx + 3).Left = lblX(indx + 2).Left + 130
                '*****************************************************
            Else
                sBar(i).Visible = False: lblBarVal(i).Visible = False
                GLine(i).Visible = False: lblLineDot(i).Visible = False
            End If
        End If
    Next i
    
    For i = 0 To jmlBar - 1
        indx = i * 4
        If MaxLen > 9 Then
            lblX(indx).Font.Size = 6: lblX(indx + 1).Font.Size = 6
            lblX(indx + 2).Font.Size = 6: lblX(indx + 3).Font.Size = 6
        Else
            lblX(indx).Font.Size = 8: lblX(indx + 1).Font.Size = 8
            lblX(indx + 2).Font.Size = 8: lblX(indx + 3).Font.Size = 8
        End If
    Next i
End Sub

Sub setLblDot(wdt As Double, Col As String)
    For i = 0 To 7
        lblLineDot(i).ForeColor = Col
    Next i
End Sub

Sub setVisibleButton(stVisible As Boolean)
    Label1(0).Visible = stVisible: cmdReport.Visible = stVisible: LinePrt.Visible = stVisible
End Sub

Private Sub cmdReport_Click()

On Error GoTo HandleErr

    Call setVisibleButton(False)
    Call SetVisibleSign(False)
    StPrint = False
    
    frmEffPrintGraph.Orient = 1
    frmEffPrintGraph.Show 1
    Call setLblDot(4815, "&H00808080")
    
    If StPrint Then Me.PrintForm
    
HandleErr:
    If err.Description = "Cancel was selected." Or err.Description = "Printer error" _
        Or err.Description = "Can't print form image to this type of printer" Then
    ElseIf err.Description <> "" Then
        MsgBox err.Description, vbCritical, "Error"
    End If
    
    Call setLblDot(4815, "&H00000000")
    Call setVisibleButton(True)
    Call SetVisibleSign(visibleSign)
End Sub

Private Sub cmdSign_Click(Index As Integer)
Select Case Index
    Case 0: 'First
        frmEffBadByMaterialControl.rsRecord.MoveFirst
    Case 1: 'Prev
        frmEffBadByMaterialControl.rsRecord.MovePrevious
        If frmEffBadByMaterialControl.rsRecord.BOF Then frmEffBadByMaterialControl.rsRecord.MoveFirst: GoTo exSub
    Case 2: 'Next
        frmEffBadByMaterialControl.rsRecord.MoveNext
        If frmEffBadByMaterialControl.rsRecord.EOF Then frmEffBadByMaterialControl.rsRecord.MoveLast: GoTo exSub
    Case 3: 'Last
        frmEffBadByMaterialControl.rsRecord.MoveLast
End Select
Call frmEffBadByMaterialControl.graphDetail(1)

exSub:
End Sub

