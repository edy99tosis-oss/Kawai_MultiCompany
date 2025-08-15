VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_Production_Planning 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Production Planning "
   ClientHeight    =   10950
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   15045
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frm_Production_Planning.frx":0000
   ScaleHeight     =   10950
   ScaleWidth      =   15045
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Clear"
      Height          =   375
      Index           =   3
      Left            =   9150
      Style           =   1  'Graphical
      TabIndex        =   87
      Top             =   10050
      Width           =   1200
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   77
      Left            =   13080
      MaxLength       =   16
      TabIndex        =   81
      Top             =   8850
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   76
      Left            =   11400
      MaxLength       =   16
      TabIndex        =   80
      Top             =   8850
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   75
      Left            =   9720
      MaxLength       =   16
      TabIndex        =   79
      Top             =   8850
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   74
      Left            =   8040
      MaxLength       =   16
      TabIndex        =   78
      Top             =   8850
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   73
      Left            =   6360
      MaxLength       =   16
      TabIndex        =   77
      Top             =   8850
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   72
      Left            =   4710
      MaxLength       =   16
      TabIndex        =   76
      Top             =   8850
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   71
      Left            =   13080
      MaxLength       =   16
      TabIndex        =   75
      Top             =   8310
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   70
      Left            =   11400
      MaxLength       =   16
      TabIndex        =   74
      Top             =   8310
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   69
      Left            =   9720
      MaxLength       =   16
      TabIndex        =   73
      Top             =   8310
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   68
      Left            =   8040
      MaxLength       =   16
      TabIndex        =   72
      Top             =   8310
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   67
      Left            =   6360
      MaxLength       =   16
      TabIndex        =   71
      Top             =   8310
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   66
      Left            =   4710
      MaxLength       =   16
      TabIndex        =   70
      Top             =   8310
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   65
      Left            =   13080
      MaxLength       =   16
      TabIndex        =   69
      Top             =   7800
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   64
      Left            =   11400
      MaxLength       =   16
      TabIndex        =   68
      Top             =   7800
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   63
      Left            =   9720
      MaxLength       =   16
      TabIndex        =   67
      Top             =   7800
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   62
      Left            =   8040
      MaxLength       =   16
      TabIndex        =   66
      Top             =   7800
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   61
      Left            =   6360
      MaxLength       =   16
      TabIndex        =   65
      Top             =   7800
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   60
      Left            =   4710
      MaxLength       =   16
      TabIndex        =   64
      Top             =   7800
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   59
      Left            =   13080
      MaxLength       =   16
      TabIndex        =   63
      Top             =   7290
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   58
      Left            =   11400
      MaxLength       =   16
      TabIndex        =   62
      Top             =   7290
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   57
      Left            =   9720
      MaxLength       =   16
      TabIndex        =   61
      Top             =   7290
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   56
      Left            =   8040
      MaxLength       =   16
      TabIndex        =   60
      Top             =   7290
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   55
      Left            =   6360
      MaxLength       =   16
      TabIndex        =   59
      Top             =   7290
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   54
      Left            =   4710
      MaxLength       =   16
      TabIndex        =   58
      Top             =   7290
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   53
      Left            =   13080
      MaxLength       =   16
      TabIndex        =   57
      Top             =   6720
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   52
      Left            =   11400
      MaxLength       =   16
      TabIndex        =   56
      Top             =   6720
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   51
      Left            =   9720
      MaxLength       =   16
      TabIndex        =   55
      Top             =   6720
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   50
      Left            =   8040
      MaxLength       =   16
      TabIndex        =   54
      Top             =   6720
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   49
      Left            =   6360
      MaxLength       =   16
      TabIndex        =   53
      Top             =   6720
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   48
      Left            =   4710
      MaxLength       =   16
      TabIndex        =   52
      Top             =   6720
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   47
      Left            =   13080
      MaxLength       =   16
      TabIndex        =   51
      Top             =   6210
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   46
      Left            =   11400
      MaxLength       =   16
      TabIndex        =   50
      Top             =   6210
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   45
      Left            =   9720
      MaxLength       =   16
      TabIndex        =   49
      Top             =   6210
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   44
      Left            =   8040
      MaxLength       =   16
      TabIndex        =   48
      Top             =   6210
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   43
      Left            =   6360
      MaxLength       =   16
      TabIndex        =   47
      Top             =   6210
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   42
      Left            =   4710
      MaxLength       =   16
      TabIndex        =   46
      Top             =   6210
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   41
      Left            =   13080
      MaxLength       =   16
      TabIndex        =   45
      Top             =   5640
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   40
      Left            =   11400
      MaxLength       =   16
      TabIndex        =   44
      Top             =   5640
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   39
      Left            =   9720
      MaxLength       =   16
      TabIndex        =   43
      Top             =   5640
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   38
      Left            =   8040
      MaxLength       =   16
      TabIndex        =   42
      Top             =   5640
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   37
      Left            =   6360
      MaxLength       =   16
      TabIndex        =   41
      Top             =   5640
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   36
      Left            =   4710
      MaxLength       =   16
      TabIndex        =   40
      Top             =   5640
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   35
      Left            =   13080
      MaxLength       =   16
      TabIndex        =   39
      Top             =   5160
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   34
      Left            =   11400
      MaxLength       =   16
      TabIndex        =   38
      Top             =   5160
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   33
      Left            =   9720
      MaxLength       =   16
      TabIndex        =   37
      Top             =   5160
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   32
      Left            =   8040
      MaxLength       =   16
      TabIndex        =   36
      Top             =   5160
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   31
      Left            =   6360
      MaxLength       =   16
      TabIndex        =   35
      Top             =   5160
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   30
      Left            =   4710
      MaxLength       =   16
      TabIndex        =   34
      Top             =   5160
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   29
      Left            =   13080
      MaxLength       =   16
      TabIndex        =   33
      Top             =   4590
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   28
      Left            =   11400
      MaxLength       =   16
      TabIndex        =   32
      Top             =   4590
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   27
      Left            =   9720
      MaxLength       =   16
      TabIndex        =   31
      Top             =   4590
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   26
      Left            =   8040
      MaxLength       =   16
      TabIndex        =   30
      Top             =   4590
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   25
      Left            =   6360
      MaxLength       =   16
      TabIndex        =   29
      Top             =   4590
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   24
      Left            =   4710
      MaxLength       =   16
      TabIndex        =   28
      Top             =   4590
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   23
      Left            =   13080
      MaxLength       =   16
      TabIndex        =   27
      Top             =   4080
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   22
      Left            =   11400
      MaxLength       =   16
      TabIndex        =   26
      Top             =   4080
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   21
      Left            =   9720
      MaxLength       =   16
      TabIndex        =   25
      Top             =   4080
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   20
      Left            =   8040
      MaxLength       =   16
      TabIndex        =   24
      Top             =   4080
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   19
      Left            =   6360
      MaxLength       =   16
      TabIndex        =   23
      Top             =   4080
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   18
      Left            =   4710
      MaxLength       =   16
      TabIndex        =   22
      Top             =   4080
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   17
      Left            =   13080
      MaxLength       =   16
      TabIndex        =   21
      Top             =   3510
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   16
      Left            =   11400
      MaxLength       =   16
      TabIndex        =   20
      Top             =   3510
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   15
      Left            =   9720
      MaxLength       =   16
      TabIndex        =   19
      Top             =   3510
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   14
      Left            =   8040
      MaxLength       =   16
      TabIndex        =   18
      Top             =   3510
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   13
      Left            =   6360
      MaxLength       =   16
      TabIndex        =   17
      Top             =   3510
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   12
      Left            =   4710
      MaxLength       =   16
      TabIndex        =   16
      Top             =   3510
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   11
      Left            =   13080
      MaxLength       =   16
      TabIndex        =   15
      Top             =   2970
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   10
      Left            =   11400
      MaxLength       =   16
      TabIndex        =   14
      Top             =   2970
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   9
      Left            =   9720
      MaxLength       =   16
      TabIndex        =   13
      Top             =   2970
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   8
      Left            =   8040
      MaxLength       =   16
      TabIndex        =   12
      Top             =   2970
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   7
      Left            =   6360
      MaxLength       =   16
      TabIndex        =   11
      Top             =   2970
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   6
      Left            =   4710
      MaxLength       =   16
      TabIndex        =   10
      Top             =   2970
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   5
      Left            =   13080
      MaxLength       =   16
      TabIndex        =   9
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   4
      Left            =   11400
      MaxLength       =   16
      TabIndex        =   8
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   3
      Left            =   9720
      MaxLength       =   16
      TabIndex        =   7
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   2
      Left            =   8040
      MaxLength       =   16
      TabIndex        =   6
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   1
      Left            =   6360
      MaxLength       =   16
      TabIndex        =   5
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H0080FFFF&
      Caption         =   "Ca&ncel"
      Height          =   375
      Index           =   1
      Left            =   10410
      Style           =   1  'Graphical
      TabIndex        =   88
      Top             =   10050
      Width           =   1200
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Submit"
      Height          =   375
      Index           =   2
      Left            =   11670
      Style           =   1  'Graphical
      TabIndex        =   82
      Top             =   10050
      Width           =   1200
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Sub &Menu"
      Height          =   375
      Index           =   0
      Left            =   330
      Style           =   1  'Graphical
      TabIndex        =   89
      Top             =   10050
      Width           =   1200
   End
   Begin VB.CommandButton cmdMove 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Last Page"
      Height          =   375
      Index           =   3
      Left            =   7500
      Style           =   1  'Graphical
      TabIndex        =   86
      Top             =   10050
      Width           =   1200
   End
   Begin VB.CommandButton cmdMove 
      BackColor       =   &H00C0E0FF&
      Caption         =   "First Page"
      Height          =   375
      Index           =   0
      Left            =   3630
      Style           =   1  'Graphical
      TabIndex        =   83
      Top             =   10050
      Width           =   1200
   End
   Begin VB.CommandButton cmdMove 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Prev Page"
      Height          =   375
      Index           =   1
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   84
      Top             =   10050
      Width           =   1200
   End
   Begin VB.CommandButton cmdMove 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Next Page"
      Height          =   375
      Index           =   2
      Left            =   6210
      Style           =   1  'Graphical
      TabIndex        =   85
      Top             =   10050
      Width           =   1200
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   330
      TabIndex        =   132
      Top             =   9360
      Width           =   12540
      Begin VB.Label lblerror 
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
         Left            =   90
         TabIndex        =   133
         Top             =   180
         Width           =   12255
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   0
      Left            =   4710
      MaxLength       =   16
      TabIndex        =   4
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   330
      TabIndex        =   91
      Top             =   630
      Width           =   14505
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   300
         Left            =   4110
         TabIndex        =   139
         Top             =   750
         Width           =   300
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FDDFE3&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4500
         TabIndex        =   137
         Text            =   "Text2"
         Top             =   780
         Width           =   2715
      End
      Begin MSComCtl2.DTPicker Mydate 
         Height          =   315
         Left            =   8130
         TabIndex        =   2
         Top             =   750
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "MMM yyyy"
         Format          =   287571971
         UpDown          =   -1  'True
         CurrentDate     =   37867
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   11760
         TabIndex        =   136
         Top             =   960
         Visible         =   0   'False
         Width           =   2400
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Update"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   13320
         TabIndex        =   135
         Top             =   870
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   4530
         X2              =   7290
         Y1              =   1020
         Y2              =   1020
      End
      Begin VB.Label lbldesc 
         BackStyle       =   0  'Transparent
         Caption         =   "xxxxxxxxxxxxxxxxxxxxxxxxxx"
         Height          =   195
         Index           =   2
         Left            =   4500
         TabIndex        =   98
         Top             =   780
         Width           =   2760
      End
      Begin MSForms.ComboBox CboPart 
         Height          =   315
         Left            =   2100
         TabIndex        =   1
         Top             =   750
         Width           =   1965
         VariousPropertyBits=   746604571
         MaxLength       =   17
         DisplayStyle    =   3
         Size            =   "3466;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Code/Part No."
         Height          =   195
         Index           =   6
         Left            =   90
         TabIndex        =   97
         Top             =   855
         Width           =   1920
      End
      Begin VB.Label lbldesc 
         BackStyle       =   0  'Transparent
         Caption         =   "xxxxxxxxxxxxxxxxxxxxxxxxxx"
         Height          =   195
         Index           =   1
         Left            =   11790
         TabIndex        =   96
         Top             =   780
         Width           =   2190
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   11760
         X2              =   13950
         Y1              =   1020
         Y2              =   1020
      End
      Begin MSForms.ComboBox cbogroup 
         Height          =   315
         Left            =   10590
         TabIndex        =   3
         Top             =   750
         Width           =   1095
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "1931;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cbodealer 
         Height          =   315
         Left            =   2100
         TabIndex        =   0
         Top             =   300
         Width           =   1335
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2355;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Factory Code"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   95
         Top             =   330
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Period"
         Height          =   195
         Index           =   2
         Left            =   7440
         TabIndex        =   94
         Top             =   810
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Group Cls"
         Height          =   195
         Index           =   3
         Left            =   9630
         TabIndex        =   93
         Top             =   810
         Width           =   855
      End
      Begin VB.Label lbldesc 
         BackStyle       =   0  'Transparent
         Caption         =   "xxxxxxxxxxxxxxxxxxxxxxxxxx"
         Height          =   195
         Index           =   0
         Left            =   3600
         TabIndex        =   92
         Top             =   330
         Width           =   9240
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   3600
         X2              =   12840
         Y1              =   540
         Y2              =   540
      End
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   435
      Left            =   12990
      TabIndex        =   138
      Top             =   150
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   767
   End
   Begin VB.Label lblpage 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Page 0 of 0"
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
      Height          =   195
      Left            =   13560
      TabIndex        =   134
      Top             =   9450
      Width           =   1095
   End
   Begin VB.Line Line1 
      Index           =   14
      X1              =   4440
      X2              =   14850
      Y1              =   8730
      Y2              =   8730
   End
   Begin VB.Line Line1 
      Index           =   13
      X1              =   4440
      X2              =   14850
      Y1              =   8190
      Y2              =   8190
   End
   Begin VB.Line Line1 
      Index           =   12
      X1              =   4440
      X2              =   14850
      Y1              =   7680
      Y2              =   7680
   End
   Begin VB.Line Line1 
      Index           =   11
      X1              =   4440
      X2              =   14850
      Y1              =   7140
      Y2              =   7140
   End
   Begin VB.Line Line1 
      Index           =   10
      X1              =   4440
      X2              =   14850
      Y1              =   6630
      Y2              =   6630
   End
   Begin VB.Line Line1 
      Index           =   9
      X1              =   4440
      X2              =   14850
      Y1              =   6090
      Y2              =   6090
   End
   Begin VB.Line Line1 
      Index           =   8
      X1              =   4440
      X2              =   14850
      Y1              =   5550
      Y2              =   5550
   End
   Begin VB.Line Line1 
      Index           =   7
      X1              =   4440
      X2              =   14850
      Y1              =   5010
      Y2              =   5010
   End
   Begin VB.Line Line1 
      Index           =   6
      X1              =   4440
      X2              =   14850
      Y1              =   4470
      Y2              =   4470
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   4440
      X2              =   14850
      Y1              =   3930
      Y2              =   3930
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   4440
      X2              =   14880
      Y1              =   3390
      Y2              =   3390
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      Height          =   7365
      Index           =   2
      Left            =   4350
      Top             =   1920
      Width           =   90
   End
   Begin VB.Label Descitem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Index           =   12
      Left            =   420
      TabIndex        =   131
      Top             =   9000
      Width           =   960
   End
   Begin VB.Label lblitem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Part No./SEBANGO"
      Height          =   195
      Index           =   12
      Left            =   420
      TabIndex        =   130
      Top             =   8790
      Width           =   3450
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   45
      Index           =   11
      Left            =   300
      Top             =   8700
      Width           =   4065
   End
   Begin VB.Label Descitem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Index           =   11
      Left            =   420
      TabIndex        =   129
      Top             =   8460
      Width           =   960
   End
   Begin VB.Label lblitem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Part No./SEBANGO"
      Height          =   195
      Index           =   11
      Left            =   420
      TabIndex        =   128
      Top             =   8250
      Width           =   3450
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   45
      Index           =   10
      Left            =   300
      Top             =   8160
      Width           =   4095
   End
   Begin VB.Label Descitem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Index           =   10
      Left            =   420
      TabIndex        =   127
      Top             =   7920
      Width           =   960
   End
   Begin VB.Label lblitem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Part No./SEBANGO"
      Height          =   195
      Index           =   10
      Left            =   420
      TabIndex        =   126
      Top             =   7710
      Width           =   3450
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   45
      Index           =   9
      Left            =   300
      Top             =   7650
      Width           =   4065
   End
   Begin VB.Label Descitem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Index           =   9
      Left            =   420
      TabIndex        =   125
      Top             =   7410
      Width           =   960
   End
   Begin VB.Label lblitem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Part No./SEBANGO"
      Height          =   195
      Index           =   9
      Left            =   420
      TabIndex        =   124
      Top             =   7200
      Width           =   3450
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   45
      Index           =   8
      Left            =   300
      Top             =   7110
      Width           =   4065
   End
   Begin VB.Label Descitem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Index           =   8
      Left            =   420
      TabIndex        =   123
      Top             =   6870
      Width           =   960
   End
   Begin VB.Label lblitem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Part No./SEBANGO"
      Height          =   195
      Index           =   8
      Left            =   420
      TabIndex        =   122
      Top             =   6660
      Width           =   3450
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   45
      Index           =   7
      Left            =   300
      Top             =   6600
      Width           =   4065
   End
   Begin VB.Label Descitem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Index           =   7
      Left            =   420
      TabIndex        =   121
      Top             =   6360
      Width           =   960
   End
   Begin VB.Label lblitem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Part No./SEBANGO"
      Height          =   195
      Index           =   7
      Left            =   420
      TabIndex        =   120
      Top             =   6150
      Width           =   3450
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   45
      Index           =   6
      Left            =   300
      Top             =   6060
      Width           =   4065
   End
   Begin VB.Label Descitem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Index           =   6
      Left            =   420
      TabIndex        =   119
      Top             =   5820
      Width           =   960
   End
   Begin VB.Label lblitem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Part No./SEBANGO"
      Height          =   195
      Index           =   6
      Left            =   420
      TabIndex        =   118
      Top             =   5610
      Width           =   3450
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   45
      Index           =   5
      Left            =   300
      Top             =   5520
      Width           =   4065
   End
   Begin VB.Label Descitem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Index           =   5
      Left            =   390
      TabIndex        =   117
      Top             =   5280
      Width           =   960
   End
   Begin VB.Label lblitem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Part No./SEBANGO"
      Height          =   195
      Index           =   5
      Left            =   390
      TabIndex        =   116
      Top             =   5070
      Width           =   3450
   End
   Begin VB.Label lblitem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Part No./SEBANGO"
      Height          =   195
      Index           =   4
      Left            =   420
      TabIndex        =   115
      Top             =   4530
      Width           =   3450
   End
   Begin VB.Label Descitem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Index           =   4
      Left            =   420
      TabIndex        =   114
      Top             =   4740
      Width           =   960
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   45
      Index           =   4
      Left            =   300
      Top             =   4980
      Width           =   4065
   End
   Begin VB.Label lblitem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Part No./SEBANGO"
      Height          =   195
      Index           =   3
      Left            =   420
      TabIndex        =   113
      Top             =   3990
      Width           =   3450
   End
   Begin VB.Label Descitem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Index           =   3
      Left            =   420
      TabIndex        =   112
      Top             =   4200
      Width           =   960
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   45
      Index           =   3
      Left            =   300
      Top             =   4440
      Width           =   4065
   End
   Begin VB.Label lblitem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Part No./SEBANGO"
      Height          =   195
      Index           =   2
      Left            =   420
      TabIndex        =   111
      Top             =   3450
      Width           =   3450
   End
   Begin VB.Label Descitem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Index           =   2
      Left            =   420
      TabIndex        =   110
      Top             =   3660
      Width           =   960
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   45
      Index           =   2
      Left            =   300
      Top             =   3900
      Width           =   4065
   End
   Begin VB.Label lblitem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Part No./SEBANGO"
      Height          =   195
      Index           =   1
      Left            =   420
      TabIndex        =   109
      Top             =   2910
      Width           =   3450
   End
   Begin VB.Label Descitem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Index           =   1
      Left            =   420
      TabIndex        =   108
      Top             =   3120
      Width           =   960
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   45
      Index           =   1
      Left            =   300
      Top             =   3360
      Width           =   4065
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   4410
      X2              =   14820
      Y1              =   2820
      Y2              =   2820
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   45
      Index           =   0
      Left            =   300
      Top             =   2790
      Width           =   4065
   End
   Begin VB.Label Descitem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Index           =   0
      Left            =   420
      TabIndex        =   107
      Top             =   2550
      Width           =   960
   End
   Begin VB.Label lblitem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Part No./SEBANGO"
      Height          =   195
      Index           =   0
      Left            =   420
      TabIndex        =   106
      Top             =   2340
      Width           =   3450
   End
   Begin VB.Label lblmonth 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "June"
      Height          =   195
      Index           =   5
      Left            =   13080
      TabIndex        =   105
      Top             =   2010
      Width           =   390
   End
   Begin VB.Label lblmonth 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "May"
      Height          =   195
      Index           =   4
      Left            =   11400
      TabIndex        =   104
      Top             =   2010
      Width           =   345
   End
   Begin VB.Label lblmonth 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "April"
      Height          =   195
      Index           =   3
      Left            =   9720
      TabIndex        =   103
      Top             =   2010
      Width           =   390
   End
   Begin VB.Label lblmonth 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "March"
      Height          =   195
      Index           =   2
      Left            =   8040
      TabIndex        =   102
      Top             =   2010
      Width           =   510
   End
   Begin VB.Label lblmonth 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "February"
      Height          =   195
      Index           =   1
      Left            =   6360
      TabIndex        =   101
      Top             =   2010
      Width           =   765
   End
   Begin VB.Label lblmonth 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "January"
      Height          =   195
      Index           =   0
      Left            =   4710
      TabIndex        =   100
      Top             =   2010
      Width           =   675
   End
   Begin VB.Label Headeritem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Part No./Product Code"
      Height          =   195
      Index           =   0
      Left            =   600
      TabIndex        =   99
      Top             =   2010
      Width           =   1920
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      Height          =   7005
      Index           =   0
      Left            =   300
      Top             =   2280
      Width           =   14550
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   1
      Left            =   300
      Top             =   1920
      Width           =   14550
   End
   Begin VB.Label title 
      Alignment       =   2  'Center
      BackColor       =   &H00FDDFE3&
      BackStyle       =   0  'Transparent
      Caption         =   "Production Planning (Forecast)"
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
      Height          =   435
      Left            =   330
      TabIndex        =   90
      Top             =   150
      Width           =   14595
   End
End
Attribute VB_Name = "Frm_Production_Planning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rstcust As Recordset, rstgroup As Recordset, rstpart As Recordset
Dim rstplan As Recordset, tgl_sb As Byte
Dim sql As String, page As Integer, totalPage  As Integer, HakU As Integer
Dim i As Integer, Y As Integer, X As String, thn(5) As String, bln(5) As String
Dim blnsubmit As Boolean, thn_sb As String * 4
Dim boldelete As Boolean, TempQty As Double, bolzero As Boolean

Sub adtocombo()
sql = "select *, trade_code from trade_master where trade_code in (select distinct manufacture_code from manufacture_line)"
Set rstcust = New Recordset
rstcust.Open sql, Db, adOpenKeyset, adLockOptimistic
With cbodealer
    .clear
    .columnCount = 2
    .ColumnWidths = "50 pt;280 pt"
    .ListWidth = 330
    .ListRows = 15
i = 0
Do Until rstcust.EOF
    .AddItem ""
    .List(i, 0) = Trim(rstcust!Trade_Code)
    .List(i, 1) = Trim(rstcust!trade_name)
    i = i + 1
    rstcust.MoveNext
Loop
End With


sql = "select * from group_cls"
Set rstgroup = New Recordset
rstgroup.Open sql, Db, adOpenKeyset, adLockOptimistic
With cboGroup
    .clear
    .columnCount = 2
    .ColumnWidths = "50 pt;75 pt"
    .ListWidth = 180
    .ListRows = 15
i = 0
Do Until rstgroup.EOF
    .AddItem ""
    .List(i, 0) = Trim(rstgroup!group_cls)
    .List(i, 1) = Trim(rstgroup!Description)
    i = i + 1
    rstgroup.MoveNext
Loop
End With
End Sub

Sub adpartcombo()
Dim sq As String
If Trim(cboGroup.Text) = "" Then
    sq = ""
Else
    sq = "and group_cls = '" & cboGroup & "'"
End If
sql = "select item_code,rtrim(makeritem_code) makeritem_code, rtrim(item_name) item_name  from item_master where " & _
        " manufacture_code = '" & cbodealer.Text & "' and finishgoodpart_cls = '01' and production_cls = '01' " & sq & " and use_endday >= convert(char(8), getdate(), 112) " & _
        " order by item_code "

' Untuk KAWAI tidak hanya Finish Good
'Sql = "select item_code,rtrim(makeritem_code) makeritem_code, rtrim(item_name) item_name  from item_master where " & _
'        " manufacture_code = '" & cbodealer.Text & "' and production_cls = '01' " & sq & " and use_endday >= convert(char(8), getdate(), 112) " & _
'        " order by item_code "

Set rstpart = New Recordset
rstpart.Open sql, Db, adOpenKeyset, adLockOptimistic
With CboPart
    .clear
    .columnCount = 3
    .ColumnWidths = "100 pt;100 pt;250 pt"
    .ListWidth = 450
    .ListRows = 15
i = 0
Do Until rstpart.EOF
    .AddItem ""
    .List(i, 0) = Trim(rstpart!Item_Code)
    .List(i, 1) = Trim(rstpart!MakerItem_Code)
    .List(i, 2) = Trim(rstpart!item_name)
    i = i + 1
    rstpart.MoveNext
Loop
End With

End Sub

Private Sub cbodealer_Click()
If Trim(cbodealer.Text) <> "" Then
    rstcust.Requery
    rstcust.Find "trade_code = '" & cbodealer.Text & "'"
    If Not rstcust.EOF Then
        lbldesc(0) = Trim(rstcust!trade_name)
        If Trim(cboGroup) <> "" Then
            adpartcombo
            rstgroup.Requery
            rstgroup.Find "group_cls = '" & cboGroup.Text & "'"
            If Not rstgroup.EOF Then
                adpartcombo
                lbldesc(1) = Trim(rstgroup!Description)
                If Trim(CboPart) <> "" Then
                    rstpart.Requery
                    rstpart.Find "item_code = '" & CboPart.Text & "'"
                    If Not rstpart.EOF Then
                        CboPart_Click
                    Else
                        lblerror = DisplayMsg(4061)
                    End If
                Else
                    CboPart_Click
                End If
            Else
                lblerror = DisplayMsg(4064)
            End If
        Else
            adpartcombo
            If Trim(CboPart) <> "" Then
                rstpart.Requery
                rstpart.Find "item_code = '" & CboPart.Text & "'"
                If Not rstpart.EOF Then
                    CboPart_Click
                Else
                    lblerror = DisplayMsg(4061)
                End If
            Else
                CboPart_Click
            End If
        End If
    Else
        clear
        clearheader
        lblerror = DisplayMsg(4060)
        lblpage = "Page 0 of 0"
    End If
End If

End Sub

Private Sub cbodealer_GotFocus()
If edited Then Frame1.Enabled = False
End Sub

Private Sub cbodealer_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 13 Then cbodealer_Click
End Sub

Private Sub cbodealer_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = Asc("'") Then KeyAscii = 0
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cbodealer_LostFocus()
rstcust.Requery
rstcust.Find "trade_code = '" & cbodealer.Text & "'"
If rstcust.EOF Then
    rstcust.Requery
    lblerror = DisplayMsg(4060)
End If
End Sub

Private Sub CboGroup_Click()
If Trim(cbodealer.Text) <> "" Then
    rstcust.Requery
    rstcust.Find "trade_code = '" & cbodealer.Text & "'"
    If Not rstcust.EOF Then
        lbldesc(0) = Trim(rstcust!trade_name)
        If Trim(cboGroup) <> "" Then
            adpartcombo
            rstgroup.Requery
            rstgroup.Find "group_cls = '" & cboGroup.Text & "'"
            If Not rstgroup.EOF Then
                adpartcombo
                lbldesc(1) = Trim(rstgroup!Description)
                If Trim(CboPart) <> "" Then
                    rstpart.Requery
                    rstpart.Find "item_code = '" & CboPart.Text & "'"
                    If Not rstpart.EOF Then
                        CboPart_Click
                    Else
                        lblerror = DisplayMsg(4061)
                    End If
                Else
                    CboPart_Click
                End If
            Else
                lblerror = DisplayMsg(4064)
            End If
        Else
            adpartcombo
            If Trim(CboPart) <> "" Then
                rstpart.Requery
                rstpart.Find "item_code = '" & CboPart.Text & "'"
                If Not rstpart.EOF Then
                    CboPart_Click
                Else
                    lblerror = DisplayMsg(4061)
                End If
            Else
                CboPart_Click
            End If
        End If
    Else
        clear
        clearheader
        lblerror = DisplayMsg(4060)
        lblpage = "Page 0 of 0"
    End If
End If
End Sub

Private Sub cbogroup_GotFocus()
If edited Then Frame1.Enabled = False
End Sub

Private Sub CboGroup_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 13 Then CboGroup_Click
End Sub

Private Sub cbogroup_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub CboPart_Click()
Dim sqlP As String
rstcust.Requery
rstcust.Find "trade_code = '" & cbodealer.Text & "'"
If Not rstcust.EOF Then
    lbldesc(0) = Trim(rstcust!trade_name)
Else
    lblerror = DisplayMsg(4060)
    Exit Sub
End If

If Trim(cboGroup) <> "" Then
    rstgroup.Requery
    rstgroup.Find "group_cls ='" & cboGroup.Text & "'"
    If Not rstgroup.EOF Then
        lbldesc(1) = Trim(rstgroup!Description)
    Else
        lblerror = DisplayMsg(4064)
        Exit Sub
    End If
    rstgroup.Requery
    rstcust.Requery
End If

If Trim(CboPart) <> "" Then
    rstpart.Requery
    rstpart.Find "item_code ='" & CboPart.Text & "'"
    If rstpart.EOF Then lblerror = DisplayMsg(4061): rstpart.Requery: Exit Sub
    lbldesc(2) = rstpart!item_name
    rstpart.Requery
End If

If Trim(cboGroup) <> "" Then
    sqlP = "and IM.group_cls = '" & cboGroup.Text & "'"
Else
    sqlP = ""
End If
sql = "select  isnull(a.qty,0) bln1, isnull(b.qty,0) bln2, isnull(c.qty,0)bln3, " & vbCrLf & _
                "isnull(d.qty,0)bln4, isnull(e.qty,0)bln5, isnull(f.qty,0)bln6 ,IM.item_Code, rtrim(IM.makeritem_code) makercode, rtrim(IM.item_name) itemname, IM.unit_cls " & vbCrLf & _
        "From " & vbCrLf & _
            "item_master IM Left join trade_master TM on TM.trade_code = IM.manufacture_code " & vbCrLf & _
            "Left Join (select item_Code, qty from production_planning where prod_year = '" & thn(0) & "' and prod_month = '" & bln(0) & "') a on a.item_code=Im.Item_Code " & vbCrLf & _
            "Left Join (select item_Code, qty from production_planning where prod_year = '" & thn(1) & "' and prod_month = '" & bln(1) & "') b on b.item_code=Im.Item_Code " & vbCrLf & _
            "Left Join (select item_Code, qty from production_planning where prod_year = '" & thn(2) & "' and prod_month = '" & bln(2) & "') c on c.item_code=Im.Item_Code " & vbCrLf & _
            "Left Join (select item_Code, qty from production_planning where prod_year = '" & thn(3) & "' and prod_month = '" & bln(3) & "') d on d.item_code=Im.Item_Code " & vbCrLf & _
            "Left Join (select item_Code, qty from production_planning where prod_year = '" & thn(4) & "' and prod_month = '" & bln(4) & "') e on e.item_code=Im.Item_Code " & vbCrLf & _
            "Left Join (select item_Code, qty from production_planning where prod_year = '" & thn(5) & "' and prod_month = '" & bln(5) & "') f on f.item_code=Im.Item_Code " & vbCrLf & _
        "Where " & vbCrLf & _
            "   IM.item_code >= '" & CboPart & "'and IM.finishgoodpart_cls = '01' and IM.production_cls='01'" & vbCrLf & _
            "   and IM.manufacture_code= '" & cbodealer.Text & "'" & sqlP & " order by IM.item_code "

Set rstplan = New Recordset
rstplan.Open sql, Db, adOpenKeyset, adLockOptimistic
rstplan.PageSize = 13
page = 0
displayrecords
End Sub

Private Sub CboPart_GotFocus()
If edited Then Frame1.Enabled = False
End Sub

Private Sub CboPart_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 13 Then CboPart_Click
End Sub

Private Sub CboPart_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub cmdAction_Click(Index As Integer)

Select Case Index
    Case 0
            blnsubmit = False
            If edited Then Exit Sub
            Unload Me
            frmMainMenu.Show
    Case 1
            blnsubmit = True
            If edited Then
                displayrecords
                Frame1.Enabled = True
                lblerror = ""
            Else
                lblerror = ""
            End If
    Case 2
            blnsubmit = True
            If HakU = 0 Then _
                lblerror = DisplayMsg(3008): Exit Sub
            If errcheck = False Then
                If edited Then
                    insertupdate
                    Frame1.Enabled = True
                    'lblerror = ""
                Else
                    lblerror = DisplayMsg(4070)
                End If
            End If
    Case 3
        blnsubmit = False
        If Not edited Then
            clear
            clearheader
            cbodealer.ListIndex = -1
            cboGroup.ListIndex = -1
            CboPart.ListIndex = -1
            MYDate = Format(Now, "MMM YYYY")
        End If
End Select
End Sub

Private Sub cmdMove_Click(Index As Integer)
blnsubmit = False
Select Case Index
    Case 0
            If edited Then Exit Sub
            If page > 1 Then
                page = 1
                displayrecords
                lblerror = ""
            End If
    Case 1
            If edited Then Exit Sub
            If page > 1 Then
                page = page - 1
                displayrecords
                lblerror = ""
            Else
                lblerror = DisplayMsg("4020")
            End If
    Case 2
            If edited Then Exit Sub
            If page < totalPage Then
                page = page + 1
                displayrecords
                lblerror = ""
            Else
                lblerror = DisplayMsg("4021")
            End If
    Case 3
            If edited Then Exit Sub
            If page < totalPage Then
                page = totalPage
                displayrecords
                lblerror = ""
            End If
End Select
End Sub

Private Sub Command1_Click()
    Me.MousePointer = vbHourglass
    frm_BrowseItem.getItemCode = CboPart.Text
    frm_BrowseItem.Show 1
    CboPart.Text = frm_BrowseItem.getItemCode
    CboPart.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    lblerror.Caption = ErrMsg
End If
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
HakU = hakUpdate(Me.Name)
adtocombo
MYDate = Format(Now, "mmm yyyy")
X = Format(Me.MYDate, "mm")
For i = 0 To 5
    If X + i <= 12 Then
        LblMonth(i) = MonthName(X + i)
        thn(i) = Year(MYDate)
        bln(i) = (X + i)
    Else
        LblMonth(i) = MonthName((X + i) - 12) & " " & (Year(MYDate) + 1)
        bln(i) = ((X + i) - 12)
        thn(i) = Year(MYDate) + 1
    End If
Next
clear
clearheader
CtrlMenu1.FormName = Me.Name
Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
tgl_sb = Month(Now)
thn_sb = Year(Now)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub lbldesc_Change(Index As Integer)
If Index = 2 Then Text2 = lbldesc(2)
End Sub


Private Sub MYDate_Change()
blnsubmit = False
If edited Then
    MYDate.Month = tgl_sb
    MYDate.Year = thn_sb
    MYDate.Year = MYDate.Year
    Frame1.Enabled = False
    Exit Sub
End If

MYDate_Click
tgl_sb = MYDate.Month
thn_sb = MYDate.Year
X = Format(Me.MYDate, "mm")
For i = 0 To 5
    If X + i <= 12 Then
        LblMonth(i) = MonthName(X + i)
        thn(i) = Year(MYDate)
        bln(i) = (X + i)
    Else
        LblMonth(i) = MonthName((X + i) - 12) & " " & (Year(MYDate) + 1)
        bln(i) = ((X + i) - 12)
        thn(i) = Year(MYDate) + 1
    End If
Next
If Trim(cbodealer.Text) <> "" Then
    rstcust.Requery
    rstcust.Find "trade_code = '" & cbodealer.Text & "'"
    If Not rstcust.EOF Then
        If Trim(CboPart) <> "" Then
        rstpart.Requery
        rstpart.Find "item_code = '" & CboPart.Text & "'"
            If Not rstpart.EOF Then
                CboPart_Click
            Else
                lblerror = DisplayMsg(4061)
            End If
        Else
            CboPart_Click
        End If
    Else
        lblerror = DisplayMsg(4060)
    End If
End If
End Sub

Sub clear()
For i = 0 To 12
    Me.lblitem(i) = ""
    Me.Descitem(i) = ""
Next
For i = 0 To 77
    Text1(i) = 0
    Text1(i).BackColor = vbWhite
    Text1(i).DataChanged = False
    Text1(i).Enabled = False
Next
lblerror = ""
End Sub

Sub clearheader()
For i = 0 To 2
    lbldesc(i) = ""
Next
End Sub
Sub displayrecords()

totalPage = rstplan.PageCount
If page > totalPage Then page = totalPage
If page > 0 Then rstplan.AbsolutePage = page

clear
If rstplan.EOF = False Then
    For Y = 0 To 12
        If page = 0 Then page = 1
        If Not rstplan.EOF Then
            lblitem(Y) = Trim(rstplan!MakerCode) & "/" & Trim(rstplan!Item_Code)
            lblitem(Y).Tag = Trim(rstplan!Item_Code)
            Me.Descitem(Y) = Trim(rstplan!itemname)
            For i = 0 To 5
                Text1(i + Y * 6).Enabled = True
                If InStr(1, rstplan.Fields(i).Value, ".") Then
                    Text1(i + Y * 6).Text = Format(rstplan.Fields(i).Value, gs_formatQty)
                Else
                    Text1(i + Y * 6).Text = Format(rstplan.Fields(i).Value, gs_formatQty)
                End If
                Text1(i + Y * 6).Tag = rstplan!Unit_cls & "," & CDbl(rstplan.Fields(i).Value)
                Text1(i + Y * 6).BackColor = vbWhite
                Text1(i + Y * 6).DataChanged = False
            Next
        rstplan.MoveNext
        End If
    Next
End If

If page < 0 Then page = 0
If totalPage < 0 Then totalPage = 0
lblpage.Caption = "Page " & page & " of " & totalPage
Dim rstime As Recordset
Set rstime = New Recordset
rstime.Open "select Last_Update from production_planning order by Last_Update desc", Db, adOpenDynamic, adLockOptimistic
If Not rstime.EOF Then
    Label3.Caption = Format(rstime!Last_Update, "dd mmm yyyy hh:mm:ss")
Else
    Label3.Caption = ""
End If
End Sub

Private Sub MYDate_Click()
If MYDate.Month = 1 And Val(tgl_sb) = 12 Then MYDate.Year = MYDate.Year + 1
If MYDate.Month = 12 And Val(tgl_sb) = 1 Then MYDate.Year = MYDate.Year - 1
End Sub

Private Sub MYDate_GotFocus()
If edited Then Frame1.Enabled = False
End Sub

Private Sub Text1_Change(Index As Integer)
If InStr(1, Text1(Index).Text, ",") = 1 Then Text1(Index) = Mid(Text1(Index), 2, Len(Text1(Index)))
End Sub

Private Sub Text1_GotFocus(Index As Integer)
Text1(Index).SelStart = 0
Text1(Index).SelLength = Len(Text1(Index).Text)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0: Exit Sub
If KeyAscii = Asc(".") Then
    If InStr(1, Text1(Index).Text, ".") And KeyAscii <> vbKeyBack Then KeyAscii = 0
End If
If Trim(Text1(Index)) = "" Or Trim(Text1(Index)) = "." Then Exit Sub
If CDbl(Text1(Index).Text) > gd_MaxQty Then
    If KeyAscii = Asc(".") Then
        If InStr(1, Text1(Index).Text, ".") Then KeyAscii = 0
    Else
        If InStr(1, Text1(Index).Text, ".") Then
            If InStr(1, Right(Text1(Index).Text, 3), ".") = 1 And KeyAscii <> vbKeyBack Then
                KeyAscii = 0
            Else
                If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
            End If
        Else
            KeyAscii = 0
            Exit Sub
        End If
    End If
Else
    If KeyAscii = Asc(".") Then
        If InStr(1, Text1(Index).Text, ".") And KeyAscii <> vbKeyBack Then KeyAscii = 0
    Else
        If InStr("0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
        
    End If
End If
End Sub

Function edited() As Boolean
edited = False
For i = 0 To 77
    If Text1(i).DataChanged Then
        If blnsubmit = False Then Text1(i).BackColor = vbRed
        If Text1(i).Text = "." Then Text1(i) = 0
        edited = True
    Else
         Text1(i).BackColor = vbWhite
    End If
Next
If edited And blnsubmit = False Then lblerror = DisplayMsg("1049")
End Function

Sub insertupdate()
Dim rst As Recordset, tempcode As String
Dim rslot As Recordset, blnupdate As Boolean, lotno As String
Dim dbz As New Connection

dbz.Open Db.ConnectionString
sql = "select cast (lot_no as numeric) LOT_NO from production_planning order by lot_no desc"
Set rslot = New Recordset
rslot.Open sql, Db, adOpenDynamic, adLockOptimistic
blnupdate = False
For Y = 0 To 12
    If Trim(lblitem(Y)) = "" Then Exit For
    'tempcode = Split(lblitem(Y), "/")(1)
    tempcode = lblitem(Y).Tag
    For i = 0 To 5
        If Text1(i + Y * 6).DataChanged = True Then
            blnupdate = True
            sql = "select *  from production_planning where item_code ='" & tempcode & "' and prod_year = '" & thn(i) & "' and prod_month = '" & bln(i) & "'"
            Set rst = New Recordset
            rst.Open sql, Db, adOpenKeyset, adLockOptimistic
            With rst
            If .EOF Then
                .AddNew
                !Item_Code = tempcode
                !prod_year = thn(i)
                !prod_month = bln(i)
                !Qty = Text1(i + Y * 6).Text
                rslot.Requery
                If rslot.EOF Or rslot.BOF Then
                    !Lot_no = "0000001"
                Else
                    !Lot_no = Format(Val(rslot!Lot_no) + 1, "0000000")
                End If
                !Unit_cls = Split(Text1(i + Y * 6).Tag, ",")(0)
                !production_date = thn(i) & "-" & bln(i) & "-01"
                !Last_Update = Now
                !last_user = userLogin
                .update
            Else
                !Item_Code = tempcode
                !prod_year = thn(i)
                !prod_month = bln(i)
                !Qty = Text1(i + Y * 6).Text
                !Unit_cls = Split(Text1(i + Y * 6).Tag, ",")(0)
                !production_date = thn(i) & "-" & bln(i) & "-01"
                !Last_Update = Now
                !last_user = userLogin
                .update
            End If
            lotno = !Lot_no
            End With

        End If
    Next
Next
Set dbz = Nothing
rstplan.Requery
displayrecords
If blnupdate Then lblerror = DisplayMsg(1000)
End Sub

Private Sub Text1_LostFocus(Index As Integer)
If Trim(Text1(Index)) = "" Or Trim(Text1(Index)) = "." Then Text1(Index) = 0
If InStr(1, Text1(Index), ".") Then
    Text1(Index) = Format(Text1(Index), gs_formatQty)
Else
    If Val(Text1(Index)) = 0 Then Text1(Index) = 0: Exit Sub
    If CDbl(Text1(Index)) > gd_MaxQty Then
        Text1(Index) = Format(Left(CDbl(Text1(Index)), 7), gs_formatQty)
    Else
    Text1(Index) = Format(Text1(Index), gs_formatQty)
    End If
End If
End Sub


Function errcheck() As Boolean
errcheck = False
If Trim(cbodealer.Text) = "" Then
    lblerror = DisplayMsg(1040)
    errcheck = True
    Exit Function
Else
    cbodealer = Trim(cbodealer)
    If cbodealer.MatchFound Then
        If Trim(CboPart) <> "" Then
            CboPart = Trim(CboPart)
            If CboPart.MatchFound Then
                If Trim(cboGroup) <> "" Then
                    cboGroup = Trim(cboGroup)
                    If cboGroup.MatchFound = False Then
                        errcheck = True
                        lblerror = DisplayMsg(4064)
                        Exit Function
                    End If
                End If
            Else
                lblerror = DisplayMsg(4061)
                errcheck = True
                Exit Function
            End If
        End If
    Else
        lblerror = DisplayMsg(4060)
        errcheck = True
        Exit Function
    End If
End If
lblerror = ""
End Function

Function Actual(ByVal zitem As String, ByVal zlot As String, ByVal zthn As String, ByVal zbln As String, ByVal zqty As Double, ByVal zqtybefore As Double) As String
Dim rstbom As Recordset, zdatebom As String
Dim rsceka As Recordset
Dim dbz As New Connection
dbz.ConnectionString = Db.ConnectionString
dbz.Open
dbz.BeginTrans
    zdatebom = zthn & Format(zbln, "00") & "01"
        
    sql = "select * from bom_master where " & _
          "  parent_itemcode ='" & zitem & "' and start_date <= '" & zdatebom & "' " & _
          " and end_date >= '" & zdatebom & "'"
    Set rstbom = New Recordset
    rstbom.CursorLocation = adUseClient
    rstbom.Open sql, dbz, adOpenDynamic, adLockOptimistic
    If Not rstbom.EOF Then
        boldelete = False
        bolzero = False
        TempQty = 0
        Call CekWIP(zitem, zdatebom, zthn, zbln, zlot, zqty, zqtybefore, 1, dbz)
    End If
    dbz.Execute sql
dbz.CommitTrans
dbz.Close
Set dbz = Nothing
End Function

Function CekWIP(Item As String, StartDate As String, zthn As String, zbln As String, zlot As String, Optional ParentQty As Double, Optional ParentQtyBefore As Double, Optional qtyAnak As Double, Optional dbx As Connection) As Boolean
Dim rstcek As Recordset, rstcekanak As Recordset, zdate As String

zdate = zthn & "-" & Format(zbln, "00") & "-01"

Set rstcek = New Recordset
sql = "select * from bom_master where " & _
      "  parent_itemcode ='" & Item & "' and start_date <= '" & StartDate & "' " & _
      " and end_date >= '" & StartDate & "'"
Set rstcek = New Recordset
rstcek.Open sql, dbx, adOpenDynamic, adLockOptimistic
If Not rstcek.EOF Then
    Do While Not rstcek.EOF
        sql = "select * from bom_master where " & _
        "  parent_itemcode ='" & rstcek!Item_Code & "' and start_date <= '" & StartDate & "' " & _
            " and end_date >= '" & StartDate & "'"
        Set rstcekanak = New Recordset
        rstcekanak.Open sql, dbx, adOpenDynamic, adLockOptimistic
        If Not rstcekanak.EOF Then
            Call CekWIP(rstcekanak!parent_itemcode, StartDate, zthn, zbln, zlot, ParentQty, ParentQtyBefore, rstcek!Qty, dbx)
        End If
        rstcek.MoveNext
    Loop
End If
End Function

Function Root(Item As String, StartDate As String) As Boolean
Dim rsRoot As Recordset
sql = "select * from bom_master where " & _
      "  item_code ='" & Item & "' and start_date <= '" & StartDate & "' " & _
      " and end_date >= '" & StartDate & "'"
Set rsRoot = New Recordset
rsRoot.Open sql, Db, adOpenDynamic, adLockOptimistic
Root = True
If Not rsRoot.EOF Then Root = False
rsRoot.Close
Set rsRoot = Nothing
End Function

