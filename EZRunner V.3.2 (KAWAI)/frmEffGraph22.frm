VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmEffGraph22 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Graph"
   ClientHeight    =   9255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14505
   FillColor       =   &H00FFC0C0&
   ForeColor       =   &H80000008&
   Icon            =   "frmEffGraph22.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   14505
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
      Picture         =   "frmEffGraph22.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   60
      Width           =   435
   End
   Begin MSComDlg.CommonDialog cdl 
      Left            =   2760
      Top             =   -30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Average"
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
      Left            =   2795
      TabIndex        =   124
      Top             =   8040
      Width           =   550
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblTot 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   2795
      TabIndex        =   123
      Top             =   8550
      Width           =   550
   End
   Begin VB.Label lblTot 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   2795
      TabIndex        =   122
      Top             =   8295
      Width           =   550
   End
   Begin VB.Label lblTot 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   2795
      TabIndex        =   121
      Top             =   8805
      Width           =   550
   End
   Begin VB.Line LineValue 
      Index           =   10
      X1              =   3330
      X2              =   3240
      Y1              =   1785
      Y2              =   1785
   End
   Begin VB.Line LineValue2 
      Index           =   10
      X1              =   2565
      X2              =   2475
      Y1              =   1785
      Y2              =   1785
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
      Left            =   6750
      TabIndex        =   120
      Top             =   1110
      Width           =   645
   End
   Begin VB.Line LineTbl 
      BorderColor     =   &H00FF00FF&
      BorderStyle     =   3  'Dot
      Index           =   2
      X1              =   440
      X2              =   720
      Y1              =   8910
      Y2              =   8910
   End
   Begin VB.Line LineTbl 
      BorderColor     =   &H0000C000&
      BorderStyle     =   4  'Dash-Dot
      Index           =   1
      X1              =   440
      X2              =   720
      Y1              =   8670
      Y2              =   8670
   End
   Begin VB.Line LineTbl 
      BorderColor     =   &H00FF8080&
      Index           =   0
      X1              =   440
      X2              =   735
      Y1              =   8430
      Y2              =   8430
   End
   Begin VB.Label lblBarVal3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   13620
      TabIndex        =   119
      Top             =   8805
      Width           =   495
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   13620
      TabIndex        =   118
      Top             =   8550
      Width           =   495
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   13620
      TabIndex        =   117
      Top             =   8295
      Width           =   495
   End
   Begin VB.Shape Dot2 
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   1
      Left            =   3750
      Shape           =   3  'Circle
      Top             =   2310
      Width           =   75
   End
   Begin VB.Shape Dot2 
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   20
      Left            =   6720
      Shape           =   3  'Circle
      Top             =   2310
      Width           =   75
   End
   Begin VB.Shape Dot2 
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   19
      Left            =   6570
      Shape           =   3  'Circle
      Top             =   2310
      Width           =   75
   End
   Begin VB.Shape Dot2 
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   18
      Left            =   6390
      Shape           =   3  'Circle
      Top             =   2310
      Width           =   75
   End
   Begin VB.Shape Dot2 
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   17
      Left            =   6240
      Shape           =   3  'Circle
      Top             =   2310
      Width           =   75
   End
   Begin VB.Shape Dot2 
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   16
      Left            =   6090
      Shape           =   3  'Circle
      Top             =   2310
      Width           =   75
   End
   Begin VB.Shape Dot2 
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   15
      Left            =   5940
      Shape           =   3  'Circle
      Top             =   2310
      Width           =   75
   End
   Begin VB.Shape Dot2 
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   14
      Left            =   5790
      Shape           =   3  'Circle
      Top             =   2310
      Width           =   75
   End
   Begin VB.Shape Dot2 
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   13
      Left            =   5640
      Shape           =   3  'Circle
      Top             =   2310
      Width           =   75
   End
   Begin VB.Shape Dot2 
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   12
      Left            =   5460
      Shape           =   3  'Circle
      Top             =   2310
      Width           =   75
   End
   Begin VB.Shape Dot2 
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   11
      Left            =   5310
      Shape           =   3  'Circle
      Top             =   2310
      Width           =   75
   End
   Begin VB.Shape Dot2 
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   10
      Left            =   5160
      Shape           =   3  'Circle
      Top             =   2310
      Width           =   75
   End
   Begin VB.Shape Dot2 
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   9
      Left            =   5040
      Shape           =   3  'Circle
      Top             =   2310
      Width           =   75
   End
   Begin VB.Shape Dot2 
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   8
      Left            =   4860
      Shape           =   3  'Circle
      Top             =   2310
      Width           =   75
   End
   Begin VB.Shape Dot2 
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   7
      Left            =   4680
      Shape           =   3  'Circle
      Top             =   2310
      Width           =   75
   End
   Begin VB.Shape Dot2 
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   6
      Left            =   4530
      Shape           =   3  'Circle
      Top             =   2310
      Width           =   75
   End
   Begin VB.Shape Dot2 
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   5
      Left            =   4380
      Shape           =   3  'Circle
      Top             =   2310
      Width           =   75
   End
   Begin VB.Shape Dot2 
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   4
      Left            =   4200
      Shape           =   3  'Circle
      Top             =   2310
      Width           =   75
   End
   Begin VB.Shape Dot2 
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   3
      Left            =   4050
      Shape           =   3  'Circle
      Top             =   2310
      Width           =   75
   End
   Begin VB.Shape Dot2 
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   2
      Left            =   3900
      Shape           =   3  'Circle
      Top             =   2310
      Width           =   75
   End
   Begin VB.Shape Dot2 
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   21
      Left            =   6870
      Shape           =   2  'Oval
      Top             =   2310
      Width           =   75
   End
   Begin VB.Shape Dot2 
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   0
      Left            =   3630
      Shape           =   3  'Circle
      Top             =   2310
      Width           =   75
   End
   Begin VB.Shape Dot3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   1
      Left            =   3810
      Shape           =   3  'Circle
      Top             =   2100
      Width           =   75
   End
   Begin VB.Shape Dot3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   20
      Left            =   6720
      Shape           =   3  'Circle
      Top             =   2100
      Width           =   75
   End
   Begin VB.Shape Dot3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   19
      Left            =   6570
      Shape           =   3  'Circle
      Top             =   2100
      Width           =   75
   End
   Begin VB.Shape Dot3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   18
      Left            =   6390
      Shape           =   3  'Circle
      Top             =   2100
      Width           =   75
   End
   Begin VB.Shape Dot3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   17
      Left            =   6240
      Shape           =   3  'Circle
      Top             =   2100
      Width           =   75
   End
   Begin VB.Shape Dot3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   16
      Left            =   6090
      Shape           =   3  'Circle
      Top             =   2100
      Width           =   75
   End
   Begin VB.Shape Dot3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   15
      Left            =   5940
      Shape           =   3  'Circle
      Top             =   2100
      Width           =   75
   End
   Begin VB.Shape Dot3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   14
      Left            =   5790
      Shape           =   3  'Circle
      Top             =   2100
      Width           =   75
   End
   Begin VB.Shape Dot3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   13
      Left            =   5640
      Shape           =   3  'Circle
      Top             =   2100
      Width           =   75
   End
   Begin VB.Shape Dot3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   12
      Left            =   5460
      Shape           =   3  'Circle
      Top             =   2100
      Width           =   75
   End
   Begin VB.Shape Dot3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   11
      Left            =   5310
      Shape           =   3  'Circle
      Top             =   2100
      Width           =   75
   End
   Begin VB.Shape Dot3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   10
      Left            =   5160
      Shape           =   3  'Circle
      Top             =   2100
      Width           =   75
   End
   Begin VB.Shape Dot3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   9
      Left            =   5040
      Shape           =   3  'Circle
      Top             =   2100
      Width           =   75
   End
   Begin VB.Shape Dot3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   8
      Left            =   4860
      Shape           =   3  'Circle
      Top             =   2100
      Width           =   75
   End
   Begin VB.Shape Dot3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   7
      Left            =   4680
      Shape           =   3  'Circle
      Top             =   2100
      Width           =   75
   End
   Begin VB.Shape Dot3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   6
      Left            =   4530
      Shape           =   3  'Circle
      Top             =   2100
      Width           =   75
   End
   Begin VB.Shape Dot3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   5
      Left            =   4380
      Shape           =   3  'Circle
      Top             =   2100
      Width           =   75
   End
   Begin VB.Shape Dot3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   4
      Left            =   4200
      Shape           =   3  'Circle
      Top             =   2100
      Width           =   75
   End
   Begin VB.Shape Dot3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   3
      Left            =   4050
      Shape           =   3  'Circle
      Top             =   2100
      Width           =   75
   End
   Begin VB.Shape Dot3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   2
      Left            =   3900
      Shape           =   3  'Circle
      Top             =   2100
      Width           =   75
   End
   Begin VB.Shape Dot3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   21
      Left            =   6870
      Shape           =   2  'Oval
      Top             =   2100
      Width           =   75
   End
   Begin VB.Shape Dot3 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   0
      Left            =   3630
      Shape           =   3  'Circle
      Top             =   2100
      Width           =   75
   End
   Begin VB.Shape Dot1 
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   21
      Left            =   6630
      Shape           =   3  'Circle
      Top             =   1890
      Width           =   75
   End
   Begin VB.Shape Dot1 
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   20
      Left            =   6510
      Shape           =   3  'Circle
      Top             =   1890
      Width           =   75
   End
   Begin VB.Shape Dot1 
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   19
      Left            =   6390
      Shape           =   3  'Circle
      Top             =   1890
      Width           =   75
   End
   Begin VB.Shape Dot1 
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   18
      Left            =   6270
      Shape           =   3  'Circle
      Top             =   1890
      Width           =   75
   End
   Begin VB.Shape Dot1 
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   17
      Left            =   6120
      Shape           =   3  'Circle
      Top             =   1890
      Width           =   75
   End
   Begin VB.Shape Dot1 
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   16
      Left            =   5970
      Shape           =   3  'Circle
      Top             =   1890
      Width           =   75
   End
   Begin VB.Shape Dot1 
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   15
      Left            =   5820
      Shape           =   3  'Circle
      Top             =   1890
      Width           =   75
   End
   Begin VB.Shape Dot1 
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   14
      Left            =   5700
      Shape           =   3  'Circle
      Top             =   1890
      Width           =   75
   End
   Begin VB.Shape Dot1 
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   13
      Left            =   5580
      Shape           =   3  'Circle
      Top             =   1890
      Width           =   75
   End
   Begin VB.Shape Dot1 
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   12
      Left            =   5460
      Shape           =   3  'Circle
      Top             =   1890
      Width           =   75
   End
   Begin VB.Shape Dot1 
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   11
      Left            =   5340
      Shape           =   3  'Circle
      Top             =   1890
      Width           =   75
   End
   Begin VB.Shape Dot1 
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   10
      Left            =   5190
      Shape           =   3  'Circle
      Top             =   1890
      Width           =   75
   End
   Begin VB.Shape Dot1 
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   9
      Left            =   5040
      Shape           =   3  'Circle
      Top             =   1890
      Width           =   75
   End
   Begin VB.Shape Dot1 
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   8
      Left            =   4890
      Shape           =   3  'Circle
      Top             =   1890
      Width           =   75
   End
   Begin VB.Shape Dot1 
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   7
      Left            =   4740
      Shape           =   3  'Circle
      Top             =   1890
      Width           =   75
   End
   Begin VB.Shape Dot1 
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   6
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   1890
      Width           =   75
   End
   Begin VB.Shape Dot1 
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   5
      Left            =   4440
      Shape           =   3  'Circle
      Top             =   1890
      Width           =   75
   End
   Begin VB.Shape Dot1 
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   4
      Left            =   4260
      Shape           =   3  'Circle
      Top             =   1890
      Width           =   75
   End
   Begin VB.Shape Dot1 
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   3
      Left            =   4080
      Shape           =   3  'Circle
      Top             =   1890
      Width           =   75
   End
   Begin VB.Shape Dot1 
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   2
      Left            =   3930
      Shape           =   3  'Circle
      Top             =   1890
      Width           =   75
   End
   Begin VB.Shape Dot1 
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   0
      Left            =   3630
      Shape           =   3  'Circle
      Top             =   1890
      Width           =   75
   End
   Begin VB.Shape Dot1 
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   1
      Left            =   3780
      Shape           =   3  'Circle
      Top             =   1890
      Width           =   75
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H0000C000&
      BorderStyle     =   4  'Dash-Dot
      Index           =   0
      X1              =   3630
      X2              =   3990
      Y1              =   5100
      Y2              =   2700
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H0000C000&
      BorderStyle     =   4  'Dash-Dot
      Index           =   1
      X1              =   3720
      X2              =   4080
      Y1              =   5100
      Y2              =   2700
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H0000C000&
      BorderStyle     =   4  'Dash-Dot
      Index           =   2
      X1              =   3840
      X2              =   4200
      Y1              =   5100
      Y2              =   2700
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H0000C000&
      BorderStyle     =   4  'Dash-Dot
      Index           =   3
      X1              =   3960
      X2              =   4320
      Y1              =   5100
      Y2              =   2700
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H0000C000&
      BorderStyle     =   4  'Dash-Dot
      Index           =   4
      X1              =   4080
      X2              =   4440
      Y1              =   5100
      Y2              =   2700
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H0000C000&
      BorderStyle     =   4  'Dash-Dot
      Index           =   5
      X1              =   4200
      X2              =   4560
      Y1              =   5100
      Y2              =   2700
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H0000C000&
      BorderStyle     =   4  'Dash-Dot
      Index           =   6
      X1              =   4320
      X2              =   4680
      Y1              =   5100
      Y2              =   2700
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H0000C000&
      BorderStyle     =   4  'Dash-Dot
      Index           =   7
      X1              =   4470
      X2              =   4830
      Y1              =   5100
      Y2              =   2700
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H0000C000&
      BorderStyle     =   4  'Dash-Dot
      Index           =   8
      X1              =   4650
      X2              =   5010
      Y1              =   5100
      Y2              =   2700
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H0000C000&
      BorderStyle     =   4  'Dash-Dot
      Index           =   9
      X1              =   4770
      X2              =   5130
      Y1              =   5100
      Y2              =   2700
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H0000C000&
      BorderStyle     =   4  'Dash-Dot
      Index           =   10
      X1              =   4920
      X2              =   5280
      Y1              =   5100
      Y2              =   2700
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H0000C000&
      BorderStyle     =   4  'Dash-Dot
      Index           =   11
      X1              =   5100
      X2              =   5460
      Y1              =   5100
      Y2              =   2700
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H0000C000&
      BorderStyle     =   4  'Dash-Dot
      Index           =   12
      X1              =   5250
      X2              =   5610
      Y1              =   5100
      Y2              =   2700
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H0000C000&
      BorderStyle     =   4  'Dash-Dot
      Index           =   13
      X1              =   5370
      X2              =   5730
      Y1              =   5100
      Y2              =   2700
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H0000C000&
      BorderStyle     =   4  'Dash-Dot
      Index           =   14
      X1              =   5520
      X2              =   5880
      Y1              =   5100
      Y2              =   2700
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H0000C000&
      BorderStyle     =   4  'Dash-Dot
      Index           =   15
      X1              =   5730
      X2              =   6090
      Y1              =   5100
      Y2              =   2700
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H0000C000&
      BorderStyle     =   4  'Dash-Dot
      Index           =   16
      X1              =   5880
      X2              =   6240
      Y1              =   5100
      Y2              =   2700
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H0000C000&
      BorderStyle     =   4  'Dash-Dot
      Index           =   17
      X1              =   6060
      X2              =   6420
      Y1              =   5100
      Y2              =   2700
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H0000C000&
      BorderStyle     =   4  'Dash-Dot
      Index           =   18
      X1              =   6240
      X2              =   6600
      Y1              =   5100
      Y2              =   2700
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H0000C000&
      BorderStyle     =   4  'Dash-Dot
      Index           =   19
      X1              =   6390
      X2              =   6750
      Y1              =   5100
      Y2              =   2700
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H0000C000&
      BorderStyle     =   4  'Dash-Dot
      Index           =   20
      X1              =   6630
      X2              =   6990
      Y1              =   5100
      Y2              =   2700
   End
   Begin VB.Line GLine2 
      BorderColor     =   &H0000C000&
      BorderStyle     =   4  'Dash-Dot
      Index           =   21
      X1              =   6840
      X2              =   7200
      Y1              =   5100
      Y2              =   2700
   End
   Begin VB.Line Line2 
      X1              =   1815
      X2              =   1815
      Y1              =   1500
      Y2              =   8040
   End
   Begin VB.Line LineValue3 
      Index           =   2
      X1              =   1815
      X2              =   1725
      Y1              =   1785
      Y2              =   1785
   End
   Begin VB.Line LineValue3 
      Index           =   1
      X1              =   1815
      X2              =   1725
      Y1              =   4785
      Y2              =   4785
   End
   Begin VB.Line LineValue3 
      Index           =   0
      X1              =   1815
      X2              =   1725
      Y1              =   7785
      Y2              =   7785
   End
   Begin VB.Line LineValue2 
      Index           =   9
      X1              =   2565
      X2              =   2475
      Y1              =   2385
      Y2              =   2385
   End
   Begin VB.Line LineValue2 
      Index           =   8
      X1              =   2565
      X2              =   2475
      Y1              =   2985
      Y2              =   2985
   End
   Begin VB.Line LineValue2 
      Index           =   7
      X1              =   2565
      X2              =   2475
      Y1              =   3585
      Y2              =   3585
   End
   Begin VB.Line LineValue2 
      Index           =   6
      X1              =   2565
      X2              =   2475
      Y1              =   4185
      Y2              =   4185
   End
   Begin VB.Line LineValue2 
      Index           =   5
      X1              =   2565
      X2              =   2475
      Y1              =   4785
      Y2              =   4785
   End
   Begin VB.Line LineValue2 
      Index           =   4
      X1              =   2565
      X2              =   2475
      Y1              =   5385
      Y2              =   5385
   End
   Begin VB.Line LineValue2 
      Index           =   3
      X1              =   2565
      X2              =   2475
      Y1              =   5985
      Y2              =   5985
   End
   Begin VB.Line LineValue2 
      Index           =   2
      X1              =   2565
      X2              =   2475
      Y1              =   6585
      Y2              =   6585
   End
   Begin VB.Line LineValue2 
      Index           =   1
      X1              =   2565
      X2              =   2475
      Y1              =   7185
      Y2              =   7185
   End
   Begin VB.Line LineValue2 
      Index           =   0
      X1              =   2565
      X2              =   2475
      Y1              =   7785
      Y2              =   7785
   End
   Begin VB.Line Line9 
      X1              =   360
      X2              =   2730
      Y1              =   8790
      Y2              =   8790
   End
   Begin VB.Line Line8 
      X1              =   360
      X2              =   2730
      Y1              =   8550
      Y2              =   8550
   End
   Begin VB.Line Line7 
      X1              =   360
      X2              =   2760
      Y1              =   8295
      Y2              =   8295
   End
   Begin VB.Shape Shape1 
      Height          =   1020
      Left            =   360
      Top             =   8040
      Width           =   2985
   End
   Begin VB.Line Line6 
      X1              =   480
      X2              =   2910
      Y1              =   8265
      Y2              =   8040
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Failure Frequency Rate"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   2
      Left            =   1020
      TabIndex        =   116
      Top             =   8835
      Width           =   1725
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Breakdown Durability Rate"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   1
      Left            =   765
      TabIndex        =   115
      Top             =   8580
      Width           =   1980
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rate of Efficiency"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   1440
      TabIndex        =   114
      Top             =   8325
      Width           =   1305
   End
   Begin VB.Line Line5 
      X1              =   14100
      X2              =   14100
      Y1              =   1500
      Y2              =   8280
   End
   Begin VB.Label lblBarVal3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   3330
      TabIndex        =   113
      Top             =   8805
      Width           =   495
   End
   Begin VB.Label lblBarVal3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   3825
      TabIndex        =   112
      Top             =   8805
      Width           =   495
   End
   Begin VB.Label lblBarVal3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   4305
      TabIndex        =   111
      Top             =   8805
      Width           =   495
   End
   Begin VB.Label lblBarVal3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   4800
      TabIndex        =   110
      Top             =   8805
      Width           =   495
   End
   Begin VB.Label lblBarVal3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   5295
      TabIndex        =   109
      Top             =   8805
      Width           =   495
   End
   Begin VB.Label lblBarVal3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   5775
      TabIndex        =   108
      Top             =   8805
      Width           =   495
   End
   Begin VB.Label lblBarVal3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   6270
      TabIndex        =   107
      Top             =   8805
      Width           =   495
   End
   Begin VB.Label lblBarVal3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   6765
      TabIndex        =   106
      Top             =   8805
      Width           =   495
   End
   Begin VB.Label lblBarVal3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   7245
      TabIndex        =   105
      Top             =   8805
      Width           =   495
   End
   Begin VB.Label lblBarVal3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   7740
      TabIndex        =   104
      Top             =   8805
      Width           =   495
   End
   Begin VB.Label lblBarVal3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   8235
      TabIndex        =   103
      Top             =   8805
      Width           =   495
   End
   Begin VB.Label lblBarVal3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   8715
      TabIndex        =   102
      Top             =   8805
      Width           =   495
   End
   Begin VB.Label lblBarVal3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   9210
      TabIndex        =   101
      Top             =   8805
      Width           =   495
   End
   Begin VB.Label lblBarVal3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   9705
      TabIndex        =   100
      Top             =   8805
      Width           =   495
   End
   Begin VB.Label lblBarVal3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   10185
      TabIndex        =   99
      Top             =   8805
      Width           =   495
   End
   Begin VB.Label lblBarVal3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   10680
      TabIndex        =   98
      Top             =   8805
      Width           =   495
   End
   Begin VB.Label lblBarVal3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   11175
      TabIndex        =   97
      Top             =   8805
      Width           =   495
   End
   Begin VB.Label lblBarVal3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   11655
      TabIndex        =   96
      Top             =   8805
      Width           =   495
   End
   Begin VB.Label lblBarVal3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   12150
      TabIndex        =   95
      Top             =   8805
      Width           =   495
   End
   Begin VB.Label lblBarVal3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   12645
      TabIndex        =   94
      Top             =   8805
      Width           =   495
   End
   Begin VB.Label lblBarVal3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   13140
      TabIndex        =   93
      Top             =   8805
      Width           =   495
   End
   Begin VB.Line LineX 
      X1              =   990
      X2              =   14060
      Y1              =   8040
      Y2              =   8040
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
      Left            =   13620
      TabIndex        =   92
      Top             =   8040
      Width           =   495
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
      Left            =   3330
      TabIndex        =   91
      Top             =   8040
      Width           =   495
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   3330
      TabIndex        =   90
      Top             =   8295
      Width           =   495
   End
   Begin VB.Line Line4 
      X1              =   360
      X2              =   14100
      Y1              =   1500
      Y2              =   1500
   End
   Begin VB.Line Line3 
      X1              =   360
      X2              =   360
      Y1              =   1500
      Y2              =   8040
   End
   Begin VB.Line LineValue 
      Index           =   9
      X1              =   3330
      X2              =   3240
      Y1              =   2385
      Y2              =   2385
   End
   Begin VB.Line LineValue 
      Index           =   8
      X1              =   3330
      X2              =   3240
      Y1              =   2985
      Y2              =   2985
   End
   Begin VB.Label lblY3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "99.9"
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
      Left            =   1320
      TabIndex        =   89
      Top             =   7785
      Width           =   420
   End
   Begin VB.Label lblY3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "99.9"
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
      Left            =   1320
      TabIndex        =   88
      Top             =   4785
      Width           =   420
   End
   Begin VB.Label lblY3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "99.9"
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
      Left            =   1350
      TabIndex        =   87
      Top             =   1785
      Width           =   420
   End
   Begin VB.Label lblY2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "99.9"
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
      Index           =   10
      Left            =   2070
      TabIndex        =   86
      Top             =   1785
      Width           =   420
   End
   Begin VB.Label lblY2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "99.9"
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
      Left            =   2070
      TabIndex        =   85
      Top             =   2385
      Width           =   420
   End
   Begin VB.Label lblY2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "99.9"
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
      Left            =   2070
      TabIndex        =   84
      Top             =   2985
      Width           =   420
   End
   Begin VB.Label lblY2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "99.9"
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
      Left            =   2070
      TabIndex        =   83
      Top             =   7785
      Width           =   420
   End
   Begin VB.Label lblY2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "99.9"
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
      Left            =   2070
      TabIndex        =   82
      Top             =   7185
      Width           =   420
   End
   Begin VB.Label lblY2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "99.9"
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
      Left            =   2070
      TabIndex        =   81
      Top             =   6585
      Width           =   420
   End
   Begin VB.Label lblY2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "99.9"
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
      Left            =   2070
      TabIndex        =   80
      Top             =   5985
      Width           =   420
   End
   Begin VB.Label lblY2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "99.9"
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
      Left            =   2070
      TabIndex        =   79
      Top             =   5385
      Width           =   420
   End
   Begin VB.Label lblY2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "99.9"
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
      Left            =   2070
      TabIndex        =   78
      Top             =   4785
      Width           =   420
   End
   Begin VB.Label lblY2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "99.9"
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
      Left            =   2070
      TabIndex        =   77
      Top             =   4185
      Width           =   420
   End
   Begin VB.Label lblY2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "99.9"
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
      Left            =   2070
      TabIndex        =   76
      Top             =   3585
      Width           =   420
   End
   Begin VB.Line Line1 
      X1              =   2565
      X2              =   2565
      Y1              =   1500
      Y2              =   8040
   End
   Begin VB.Label lblY 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "99.9"
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
      Index           =   10
      Left            =   2820
      TabIndex        =   75
      Top             =   1785
      Width           =   420
   End
   Begin VB.Label lblY 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "99.9"
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
      Left            =   2820
      TabIndex        =   74
      Top             =   2385
      Width           =   420
   End
   Begin VB.Label lblY 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "99.9"
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
      Left            =   2820
      TabIndex        =   73
      Top             =   2985
      Width           =   420
   End
   Begin VB.Label lblY 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "99.9"
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
      Left            =   2820
      TabIndex        =   72
      Top             =   7785
      Width           =   420
   End
   Begin VB.Label lblY 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "99.9"
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
      Left            =   2820
      TabIndex        =   71
      Top             =   7185
      Width           =   420
   End
   Begin VB.Label lblY 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "99.9"
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
      Left            =   2820
      TabIndex        =   70
      Top             =   6585
      Width           =   420
   End
   Begin VB.Label lblY 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "99.9"
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
      Left            =   2820
      TabIndex        =   69
      Top             =   5985
      Width           =   420
   End
   Begin VB.Label lblY 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "99.9"
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
      Left            =   2820
      TabIndex        =   68
      Top             =   5385
      Width           =   420
   End
   Begin VB.Label lblY 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "99.9"
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
      Left            =   2820
      TabIndex        =   67
      Top             =   4785
      Width           =   420
   End
   Begin VB.Label lblY 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "99.9"
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
      Left            =   2820
      TabIndex        =   66
      Top             =   4185
      Width           =   420
   End
   Begin VB.Label lblY 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "99.9"
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
      Left            =   2820
      TabIndex        =   65
      Top             =   3585
      Width           =   420
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   13140
      TabIndex        =   64
      Top             =   8550
      Width           =   495
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   12645
      TabIndex        =   63
      Top             =   8550
      Width           =   495
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   12150
      TabIndex        =   62
      Top             =   8550
      Width           =   495
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   11655
      TabIndex        =   61
      Top             =   8550
      Width           =   495
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   11175
      TabIndex        =   60
      Top             =   8550
      Width           =   495
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   10680
      TabIndex        =   59
      Top             =   8550
      Width           =   495
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   10185
      TabIndex        =   58
      Top             =   8550
      Width           =   495
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   9705
      TabIndex        =   57
      Top             =   8550
      Width           =   495
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   9210
      TabIndex        =   56
      Top             =   8550
      Width           =   495
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   8715
      TabIndex        =   55
      Top             =   8550
      Width           =   495
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   8235
      TabIndex        =   54
      Top             =   8550
      Width           =   495
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   7740
      TabIndex        =   53
      Top             =   8550
      Width           =   495
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   7245
      TabIndex        =   52
      Top             =   8550
      Width           =   495
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   6765
      TabIndex        =   51
      Top             =   8550
      Width           =   495
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   6270
      TabIndex        =   50
      Top             =   8550
      Width           =   495
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   5775
      TabIndex        =   49
      Top             =   8550
      Width           =   495
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   5295
      TabIndex        =   48
      Top             =   8550
      Width           =   495
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   4800
      TabIndex        =   47
      Top             =   8550
      Width           =   495
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   4305
      TabIndex        =   46
      Top             =   8550
      Width           =   495
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   3825
      TabIndex        =   45
      Top             =   8550
      Width           =   495
   End
   Begin VB.Label lblBarVal2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   3330
      TabIndex        =   44
      Top             =   8550
      Width           =   495
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   13140
      TabIndex        =   43
      Top             =   8295
      Width           =   495
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
      Left            =   13140
      TabIndex        =   42
      Top             =   8040
      Width           =   495
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   12645
      TabIndex        =   41
      Top             =   8295
      Width           =   495
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   12150
      TabIndex        =   40
      Top             =   8295
      Width           =   495
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   11655
      TabIndex        =   39
      Top             =   8295
      Width           =   495
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   11175
      TabIndex        =   38
      Top             =   8295
      Width           =   495
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   10680
      TabIndex        =   37
      Top             =   8295
      Width           =   495
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   10185
      TabIndex        =   36
      Top             =   8295
      Width           =   495
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   9705
      TabIndex        =   35
      Top             =   8295
      Width           =   495
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   9210
      TabIndex        =   34
      Top             =   8295
      Width           =   495
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   8715
      TabIndex        =   33
      Top             =   8295
      Width           =   495
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   8235
      TabIndex        =   32
      Top             =   8295
      Width           =   495
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   7740
      TabIndex        =   31
      Top             =   8295
      Width           =   495
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   7245
      TabIndex        =   30
      Top             =   8295
      Width           =   495
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   6765
      TabIndex        =   29
      Top             =   8295
      Width           =   495
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   6270
      TabIndex        =   28
      Top             =   8295
      Width           =   495
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   5775
      TabIndex        =   27
      Top             =   8295
      Width           =   495
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   5295
      TabIndex        =   26
      Top             =   8295
      Width           =   495
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   4800
      TabIndex        =   25
      Top             =   8295
      Width           =   495
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   4305
      TabIndex        =   24
      Top             =   8295
      Width           =   495
   End
   Begin VB.Label lblBarVal1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   3825
      TabIndex        =   23
      Top             =   8295
      Width           =   495
   End
   Begin VB.Line GLine3 
      BorderColor     =   &H00FF00FF&
      BorderStyle     =   3  'Dot
      Index           =   21
      X1              =   10860
      X2              =   11220
      Y1              =   7830
      Y2              =   5430
   End
   Begin VB.Line GLine3 
      BorderColor     =   &H00FF00FF&
      BorderStyle     =   3  'Dot
      Index           =   20
      X1              =   10650
      X2              =   11010
      Y1              =   7830
      Y2              =   5430
   End
   Begin VB.Line GLine3 
      BorderColor     =   &H00FF00FF&
      BorderStyle     =   3  'Dot
      Index           =   19
      X1              =   10410
      X2              =   10770
      Y1              =   7830
      Y2              =   5430
   End
   Begin VB.Line GLine3 
      BorderColor     =   &H00FF00FF&
      BorderStyle     =   3  'Dot
      Index           =   18
      X1              =   10260
      X2              =   10620
      Y1              =   7830
      Y2              =   5430
   End
   Begin VB.Line GLine3 
      BorderColor     =   &H00FF00FF&
      BorderStyle     =   3  'Dot
      Index           =   17
      X1              =   10080
      X2              =   10440
      Y1              =   7830
      Y2              =   5430
   End
   Begin VB.Line GLine3 
      BorderColor     =   &H00FF00FF&
      BorderStyle     =   3  'Dot
      Index           =   16
      X1              =   9900
      X2              =   10260
      Y1              =   7830
      Y2              =   5430
   End
   Begin VB.Line GLine3 
      BorderColor     =   &H00FF00FF&
      BorderStyle     =   3  'Dot
      Index           =   15
      X1              =   9750
      X2              =   10110
      Y1              =   7830
      Y2              =   5430
   End
   Begin VB.Line GLine3 
      BorderColor     =   &H00FF00FF&
      BorderStyle     =   3  'Dot
      Index           =   14
      X1              =   9540
      X2              =   9900
      Y1              =   7830
      Y2              =   5430
   End
   Begin VB.Line GLine3 
      BorderColor     =   &H00FF00FF&
      BorderStyle     =   3  'Dot
      Index           =   13
      X1              =   9390
      X2              =   9750
      Y1              =   7830
      Y2              =   5430
   End
   Begin VB.Line GLine3 
      BorderColor     =   &H00FF00FF&
      BorderStyle     =   3  'Dot
      Index           =   12
      X1              =   9240
      X2              =   9600
      Y1              =   7830
      Y2              =   5430
   End
   Begin VB.Line GLine3 
      BorderColor     =   &H00FF00FF&
      BorderStyle     =   3  'Dot
      Index           =   11
      X1              =   9120
      X2              =   9480
      Y1              =   7830
      Y2              =   5430
   End
   Begin VB.Line GLine3 
      BorderColor     =   &H00FF00FF&
      BorderStyle     =   3  'Dot
      Index           =   10
      X1              =   8940
      X2              =   9300
      Y1              =   7830
      Y2              =   5430
   End
   Begin VB.Line GLine3 
      BorderColor     =   &H00FF00FF&
      BorderStyle     =   3  'Dot
      Index           =   9
      X1              =   8790
      X2              =   9150
      Y1              =   7830
      Y2              =   5430
   End
   Begin VB.Line GLine3 
      BorderColor     =   &H00FF00FF&
      BorderStyle     =   3  'Dot
      Index           =   8
      X1              =   8670
      X2              =   9030
      Y1              =   7830
      Y2              =   5430
   End
   Begin VB.Line GLine3 
      BorderColor     =   &H00FF00FF&
      BorderStyle     =   3  'Dot
      Index           =   7
      X1              =   8490
      X2              =   8850
      Y1              =   7830
      Y2              =   5430
   End
   Begin VB.Line GLine3 
      BorderColor     =   &H00FF00FF&
      BorderStyle     =   3  'Dot
      Index           =   6
      X1              =   8340
      X2              =   8700
      Y1              =   7830
      Y2              =   5430
   End
   Begin VB.Line GLine3 
      BorderColor     =   &H00FF00FF&
      BorderStyle     =   3  'Dot
      Index           =   5
      X1              =   8220
      X2              =   8580
      Y1              =   7830
      Y2              =   5430
   End
   Begin VB.Line GLine3 
      BorderColor     =   &H00FF00FF&
      BorderStyle     =   3  'Dot
      Index           =   4
      X1              =   8100
      X2              =   8460
      Y1              =   7830
      Y2              =   5430
   End
   Begin VB.Line GLine3 
      BorderColor     =   &H00FF00FF&
      BorderStyle     =   3  'Dot
      Index           =   3
      X1              =   7980
      X2              =   8340
      Y1              =   7830
      Y2              =   5430
   End
   Begin VB.Line GLine3 
      BorderColor     =   &H00FF00FF&
      BorderStyle     =   3  'Dot
      Index           =   2
      X1              =   7860
      X2              =   8220
      Y1              =   7830
      Y2              =   5430
   End
   Begin VB.Line GLine3 
      BorderColor     =   &H00FF00FF&
      BorderStyle     =   3  'Dot
      Index           =   1
      X1              =   7740
      X2              =   8100
      Y1              =   7830
      Y2              =   5430
   End
   Begin VB.Line GLine3 
      BorderColor     =   &H00FF00FF&
      BorderStyle     =   3  'Dot
      Index           =   0
      X1              =   7650
      X2              =   8010
      Y1              =   7830
      Y2              =   5430
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      Index           =   21
      X1              =   6420
      X2              =   6780
      Y1              =   7820
      Y2              =   5430
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      Index           =   20
      X1              =   6300
      X2              =   6660
      Y1              =   7820
      Y2              =   5430
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      Index           =   19
      X1              =   6150
      X2              =   6510
      Y1              =   7820
      Y2              =   5430
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      Index           =   18
      X1              =   6060
      X2              =   6420
      Y1              =   7820
      Y2              =   5430
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      Index           =   17
      X1              =   5970
      X2              =   6330
      Y1              =   7820
      Y2              =   5430
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      Index           =   16
      X1              =   5790
      X2              =   6150
      Y1              =   7820
      Y2              =   5430
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      Index           =   15
      X1              =   5670
      X2              =   6030
      Y1              =   7820
      Y2              =   5430
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      Index           =   14
      X1              =   5520
      X2              =   5880
      Y1              =   7820
      Y2              =   5430
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      Index           =   13
      X1              =   5400
      X2              =   5760
      Y1              =   7820
      Y2              =   5430
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      Index           =   12
      X1              =   5190
      X2              =   5550
      Y1              =   7820
      Y2              =   5430
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      Index           =   11
      X1              =   5070
      X2              =   5430
      Y1              =   7820
      Y2              =   5430
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      Index           =   10
      X1              =   4920
      X2              =   5280
      Y1              =   7820
      Y2              =   5430
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      Index           =   9
      X1              =   4770
      X2              =   5130
      Y1              =   7820
      Y2              =   5430
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      Index           =   8
      X1              =   4650
      X2              =   5010
      Y1              =   7820
      Y2              =   5430
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
      Left            =   12645
      TabIndex        =   22
      Top             =   8040
      Width           =   495
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
      Left            =   12150
      TabIndex        =   21
      Top             =   8040
      Width           =   495
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
      Left            =   11655
      TabIndex        =   20
      Top             =   8040
      Width           =   495
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
      Left            =   11175
      TabIndex        =   19
      Top             =   8040
      Width           =   495
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
      Left            =   10680
      TabIndex        =   18
      Top             =   8040
      Width           =   495
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
      Left            =   10185
      TabIndex        =   17
      Top             =   8040
      Width           =   495
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
      Left            =   9705
      TabIndex        =   16
      Top             =   8040
      Width           =   495
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
      Left            =   9210
      TabIndex        =   15
      Top             =   8040
      Width           =   495
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
      Left            =   8715
      TabIndex        =   14
      Top             =   8040
      Width           =   495
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
      Left            =   8235
      TabIndex        =   13
      Top             =   8040
      Width           =   495
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
      Left            =   7740
      TabIndex        =   12
      Top             =   8040
      Width           =   495
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
      Left            =   7245
      TabIndex        =   11
      Top             =   8040
      Width           =   495
   End
   Begin VB.Line LinePrt 
      BorderWidth     =   2
      X1              =   0
      X2              =   14500
      Y1              =   510
      Y2              =   510
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
      Left            =   6525
      TabIndex        =   9
      Top             =   870
      Width           =   1095
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Efficiency Control"
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
      Height          =   330
      Index           =   0
      Left            =   6045
      TabIndex        =   8
      Top             =   570
      Width           =   2055
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
      Left            =   540
      TabIndex        =   7
      Top             =   120
      Width           =   390
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      Index           =   7
      X1              =   4500
      X2              =   4860
      Y1              =   7820
      Y2              =   5430
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      Index           =   6
      X1              =   4380
      X2              =   4740
      Y1              =   7820
      Y2              =   5430
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      Index           =   5
      X1              =   4200
      X2              =   4560
      Y1              =   7820
      Y2              =   5430
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      Index           =   4
      X1              =   4020
      X2              =   4380
      Y1              =   7820
      Y2              =   5430
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      Index           =   3
      X1              =   3900
      X2              =   4260
      Y1              =   7820
      Y2              =   5430
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      Index           =   2
      X1              =   3780
      X2              =   4140
      Y1              =   7820
      Y2              =   5430
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      Index           =   1
      X1              =   3660
      X2              =   4020
      Y1              =   7820
      Y2              =   5430
   End
   Begin VB.Line LineValue 
      Index           =   7
      X1              =   3330
      X2              =   3240
      Y1              =   3585
      Y2              =   3585
   End
   Begin VB.Line LineValue 
      Index           =   6
      X1              =   3335
      X2              =   3255
      Y1              =   4185
      Y2              =   4185
   End
   Begin VB.Line GLine1 
      BorderColor     =   &H00FF8080&
      Index           =   0
      X1              =   3480
      X2              =   3840
      Y1              =   7895
      Y2              =   5430
   End
   Begin VB.Line LineYValue 
      X1              =   3330
      X2              =   3330
      Y1              =   1500
      Y2              =   8040
   End
   Begin VB.Line LineValue 
      Index           =   3
      X1              =   3335
      X2              =   3255
      Y1              =   5985
      Y2              =   5985
   End
   Begin VB.Line LineValue 
      Index           =   5
      X1              =   3335
      X2              =   3255
      Y1              =   4785
      Y2              =   4785
   End
   Begin VB.Line LineValue 
      Index           =   1
      X1              =   3335
      X2              =   3255
      Y1              =   7185
      Y2              =   7185
   End
   Begin VB.Line LineValue 
      Index           =   2
      X1              =   3335
      X2              =   3255
      Y1              =   6585
      Y2              =   6585
   End
   Begin VB.Line LineValue 
      Index           =   0
      X1              =   3335
      X2              =   3255
      Y1              =   7790
      Y2              =   7785
   End
   Begin VB.Line LineValue 
      Index           =   4
      X1              =   3335
      X2              =   3255
      Y1              =   5385
      Y2              =   5385
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
      Left            =   6765
      TabIndex        =   6
      Top             =   8040
      Width           =   495
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
      Left            =   6270
      TabIndex        =   5
      Top             =   8040
      Width           =   495
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
      Left            =   5775
      TabIndex        =   4
      Top             =   8040
      Width           =   495
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
      Left            =   5295
      TabIndex        =   3
      Top             =   8040
      Width           =   495
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
      Left            =   4800
      TabIndex        =   2
      Top             =   8040
      Width           =   495
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
      Left            =   4305
      TabIndex        =   1
      Top             =   8040
      Width           =   495
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
      Left            =   3825
      TabIndex        =   0
      Top             =   8040
      Width           =   495
   End
End
Attribute VB_Name = "frmEffGraph22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i  As Integer

Dim XLeft As Double, XRight As Double, XTop As Double, YTop As Double
Dim RangeY As Double, MaxY As Double, totLineY As Integer, distY As Double, distX As Double
Dim Percentage As Double, Percentage2 As Double, Percentage3 As Double

Public jmlBar As Double
Public MaxQty As Double, MaxQty2 As Double, MaxQty3 As Double
Public tot As Double, tot2 As Double, tot3 As Double

Public Sub viewGraph()
    XLeft = LineX.X1: XRight = LineX.x2: XTop = LineX.Y1 - 285: YTop = LineYValue.Y1
    distY = 600: distX = 500
    
    MaxQty = setRange(MaxQty, True)
    MaxQty2 = setRange(MaxQty2, True)
    MaxQty3 = RoundUp(MaxQty3)
    
    tot = 0: tot2 = 0: tot3 = 0
    Percentage = 0: Percentage2 = 0: Percentage2 = 30
    
    For i = 0 To 21
        If i <= 10 Then
            lblY(i) = Format((MaxQty / 10) * (i), gs_formatEfficiency)
            lblY2(i) = Format((MaxQty2 / 10) * (i), gs_formatEfficiency)
            If i <= 2 Then lblY3(i) = Format((MaxQty3 / 2) * (i), gs_formatEfficiency)
        End If

        If i < jmlBar Then
            GLine1(i).Visible = True: GLine2(i).Visible = True: GLine3(i).Visible = True
            Dot1(i).Visible = True: Dot2(i).Visible = True: Dot3(i).Visible = True
            
            If MaxQty = 0 Then Percentage = 0 Else Percentage = CDbl(lblBarVal1(i) / MaxQty) * 10 * distY
            If MaxQty2 = 0 Then Percentage2 = 0 Else Percentage2 = CDbl(lblBarVal2(i) / MaxQty2) * 10 * distY
            If MaxQty3 = 0 Then Percentage3 = 0 Else Percentage3 = CDbl(lblBarVal3(i) / MaxQty3) * 2 * (distY * 5)
            
            '*********** Set Top & Height *****************
            If i = 0 Then
                GLine1(i).X1 = lblX(i).Left + (distX / 2): GLine1(i).x2 = GLine1(i).X1
                GLine2(i).X1 = lblX(i).Left + (distX / 2): GLine2(i).x2 = GLine2(i).X1
                GLine3(i).X1 = lblX(i).Left + (distX / 2): GLine3(i).x2 = GLine3(i).X1
                
                GLine1(i).Y1 = XTop - Percentage: GLine1(i).Y2 = GLine1(i).Y1
                GLine2(i).Y1 = XTop - Percentage2: GLine2(i).Y2 = GLine2(i).Y1
                GLine3(i).Y1 = XTop - Percentage3: GLine3(i).Y2 = GLine3(i).Y1
            
            ElseIf i > 0 And i < 21 Then
                GLine1(i).x2 = lblX(i + 1).Left - (distX / 2)
                GLine2(i).x2 = lblX(i + 1).Left - (distX / 2)
                GLine3(i).x2 = lblX(i + 1).Left - (distX / 2)
            
            ElseIf i = 21 Then
                GLine1(i).x2 = XRight - (distX / 2)
                GLine2(i).x2 = XRight - (distX / 2)
                GLine3(i).x2 = XRight - (distX / 2)
            End If
            
            If i <> 0 Then
                GLine1(i).X1 = GLine1(i - 1).x2
                GLine2(i).X1 = GLine2(i - 1).x2
                GLine3(i).X1 = GLine3(i - 1).x2
                
                GLine1(i).Y1 = GLine1(i - 1).Y2: GLine1(i).Y2 = XTop - Percentage
                GLine2(i).Y1 = GLine2(i - 1).Y2: GLine2(i).Y2 = XTop - Percentage2
                GLine3(i).Y1 = GLine3(i - 1).Y2: GLine3(i).Y2 = XTop - Percentage3
            End If
            Dot1(i).Left = GLine1(i).x2: Dot1(i).top = GLine1(i).Y2
            Dot2(i).Left = GLine2(i).x2: Dot2(i).top = GLine2(i).Y2
            Dot3(i).Left = GLine3(i).x2: Dot3(i).top = GLine3(i).Y2
            tot = tot + lblBarVal1(i): tot2 = tot2 + lblBarVal2(i): tot3 = tot3 + lblBarVal3(i)
        Else
            lblX(i) = ""
            lblBarVal1(i) = "": lblBarVal2(i) = "": lblBarVal3(i) = ""
            GLine1(i).Visible = False: GLine2(i).Visible = False: GLine3(i).Visible = False
            Dot1(i).Visible = False: Dot2(i).Visible = False: Dot3(i).Visible = False
        End If
    Next i
    lblTot(0) = Format(tot / jmlBar, gs_formatEfficiency)
    lblTot(1) = Format(tot2 / jmlBar, gs_formatEfficiency)
    lblTot(2) = Format(tot3 / jmlBar, gs_formatEfficiency)
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




