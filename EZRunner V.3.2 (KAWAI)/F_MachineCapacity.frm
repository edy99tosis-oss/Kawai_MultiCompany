VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form F_MachineCapacity 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Machine Capacity (Item Code)"
   ClientHeight    =   9330
   ClientLeft      =   765
   ClientTop       =   2895
   ClientWidth     =   13860
   Icon            =   "F_MachineCapacity.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   9330
   ScaleWidth      =   13860
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox TxtQtyDay 
      Alignment       =   1  'Right Justify
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
      Height          =   350
      Left            =   11940
      Locked          =   -1  'True
      MaxLength       =   23
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   7140
      Width           =   1125
   End
   Begin VB.TextBox TxtQtyHour 
      Alignment       =   1  'Right Justify
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
      Height          =   350
      Left            =   10620
      Locked          =   -1  'True
      MaxLength       =   23
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   7140
      Width           =   1125
   End
   Begin VB.TextBox txtQtyProcess 
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
      Height          =   350
      Left            =   9210
      MaxLength       =   23
      TabIndex        =   8
      Top             =   7140
      Width           =   1125
   End
   Begin VB.TextBox txtQtyMachine 
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
      Height          =   350
      Left            =   7890
      MaxLength       =   23
      TabIndex        =   7
      Top             =   7140
      Width           =   1125
   End
   Begin VB.TextBox txtEficiency 
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
      Height          =   350
      Left            =   6630
      MaxLength       =   23
      TabIndex        =   6
      Top             =   7140
      Width           =   1125
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   473
      TabIndex        =   19
      Top             =   7695
      Width           =   12915
      Begin VB.Label LblErrMsg 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Left            =   105
         TabIndex        =   20
         Top             =   195
         Width           =   12690
      End
   End
   Begin VB.CommandButton CmdClear 
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
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8400
      Width           =   1200
   End
   Begin VB.CommandButton CmdCancel 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Cancel"
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
      Left            =   10905
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8400
      Width           =   1200
   End
   Begin VB.CommandButton CmdSubMenu 
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
      Left            =   473
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8400
      Width           =   1200
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
      Left            =   12188
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8400
      Width           =   1200
   End
   Begin VB.TextBox txtCycle 
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
      Height          =   350
      Left            =   5400
      MaxLength       =   23
      TabIndex        =   5
      Top             =   7140
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1680
      Left            =   473
      TabIndex        =   14
      Top             =   900
      Width           =   12915
      Begin VB.Label lblMachine 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Machine Name"
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
         Left            =   3360
         TabIndex        =   37
         Top             =   1170
         Width           =   3675
      End
      Begin VB.Label lblLineName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Line Name"
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
         Left            =   3360
         TabIndex        =   36
         Top             =   780
         Width           =   3660
      End
      Begin VB.Label lblFactory 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Factory Name"
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
         Left            =   3360
         TabIndex        =   35
         Top             =   360
         Width           =   1185
      End
      Begin VB.Line Line2 
         Index           =   1
         X1              =   3360
         X2              =   7020
         Y1              =   1410
         Y2              =   1410
      End
      Begin MSForms.ComboBox cboMachine 
         Height          =   345
         Left            =   1725
         TabIndex        =   2
         Top             =   1140
         Width           =   1515
         VariousPropertyBits=   746604571
         MaxLength       =   10
         DisplayStyle    =   3
         Size            =   "2672;609"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Machine Code"
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
         Left            =   405
         TabIndex        =   27
         Top             =   1215
         Width           =   1200
      End
      Begin VB.Label lblItem 
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
         Left            =   3465
         TabIndex        =   18
         Top             =   780
         Width           =   60
      End
      Begin VB.Label lblGroup 
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
         Left            =   3465
         TabIndex        =   17
         Top             =   345
         Width           =   60
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   3360
         X2              =   7020
         Y1              =   1020
         Y2              =   1020
      End
      Begin VB.Line Line1 
         X1              =   3360
         X2              =   7020
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Line Code"
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
         Left            =   810
         TabIndex        =   16
         Top             =   780
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Factory Code"
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
         TabIndex        =   15
         Top             =   345
         Width           =   1140
      End
      Begin MSForms.ComboBox cboLineCode 
         Height          =   345
         Left            =   1725
         TabIndex        =   1
         Top             =   690
         Width           =   1515
         VariousPropertyBits=   746604571
         MaxLength       =   10
         DisplayStyle    =   3
         Size            =   "2672;609"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboFactori 
         Height          =   345
         Left            =   1725
         TabIndex        =   0
         Top             =   270
         Width           =   1515
         VariousPropertyBits=   746604571
         MaxLength       =   10
         DisplayStyle    =   3
         Size            =   "2672;609"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   3675
      Left            =   473
      TabIndex        =   3
      Top             =   2850
      Width           =   12915
      _cx             =   22781
      _cy             =   6482
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   275
      RowHeightMax    =   275
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"F_MachineCapacity.frx":0E42
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
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   420
      Left            =   11460
      TabIndex        =   13
      Top             =   255
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   741
   End
   Begin VB.Label LblItemName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblItemName"
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
      Left            =   2220
      TabIndex        =   38
      Top             =   7230
      Width           =   1095
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qty / Day"
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
      Left            =   12060
      TabIndex        =   33
      Top             =   6765
      Width           =   840
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qty/ Hour"
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
      Left            =   10740
      TabIndex        =   31
      Top             =   6765
      Width           =   840
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qty Per Process"
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
      Left            =   9150
      TabIndex        =   30
      Top             =   6765
      Width           =   1365
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qty Machine"
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
      Left            =   7830
      TabIndex        =   29
      Top             =   6780
      Width           =   1050
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00A6D2FF&
      Height          =   585
      Index           =   0
      Left            =   473
      Top             =   7035
      Width           =   12915
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%Eficiency"
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
      Left            =   6660
      TabIndex        =   28
      Top             =   6765
      Width           =   930
   End
   Begin VB.Label Label5 
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
      Left            =   2235
      TabIndex        =   26
      Top             =   6765
      Width           =   960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Code"
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
      Left            =   630
      TabIndex        =   25
      Top             =   6765
      Width           =   915
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cycle Time"
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
      Left            =   5460
      TabIndex        =   22
      Top             =   6765
      Width           =   960
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Machine Capacity (Item Code)"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   480
      TabIndex        =   24
      Top             =   285
      Width           =   12915
   End
   Begin VB.Label Isian 
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
      Left            =   1935
      TabIndex        =   23
      Top             =   8460
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00A6D2FF&
      BackStyle       =   1  'Opaque
      Height          =   345
      Index           =   2
      Left            =   473
      Top             =   6705
      Width           =   12915
   End
   Begin VB.Label lblMachineName 
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
      Left            =   2205
      TabIndex        =   21
      Top             =   7230
      Width           =   60
   End
   Begin VB.Line Line4 
      X1              =   2190
      X2              =   5340
      Y1              =   7470
      Y2              =   7470
   End
   Begin MSForms.ComboBox cboItem 
      Height          =   345
      Left            =   600
      TabIndex        =   4
      Top             =   7140
      Width           =   1515
      VariousPropertyBits=   746604571
      MaxLength       =   2
      DisplayStyle    =   3
      Size            =   "2672;609"
      ListRows        =   10
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "F_MachineCapacity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ColItem As Byte, ColDesc As Byte, ColCycleTime As Byte, ColEficiency As Byte
Dim ColQtyMachine As Byte, colQtyProcess As Byte, ColQtyHour As Byte, ColQtyDay As Byte
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim rscek As New ADODB.Recordset
Dim Tekan As Integer

Private Sub header()
 ColItem = 1
 ColDesc = 2
 ColCycleTime = 3
 ColEficiency = 4
 ColQtyMachine = 5
 colQtyProcess = 6
 ColQtyHour = 7
 ColQtyDay = 8
 With Grid
  .ColS = 9
  .Rows = 1
    
  .TextMatrix(0, 0) = " "
  .TextMatrix(0, ColItem) = "Item Code"
  .TextMatrix(0, ColDesc) = "Description"
  .TextMatrix(0, ColCycleTime) = "Cycle Time"
  .TextMatrix(0, ColEficiency) = "%Eficiency"
  .TextMatrix(0, ColQtyMachine) = "Qty Machine"
  .TextMatrix(0, colQtyProcess) = "Qty Per Process"
  .TextMatrix(0, ColQtyHour) = "Qty/Hour"
  .TextMatrix(0, ColQtyDay) = "Qty/Day"

  .ColWidth(0) = 300
  .ColWidth(ColItem) = 1400
  .ColWidth(ColDesc) = 1900
  .ColWidth(ColCycleTime) = 1500
  .ColWidth(ColEficiency) = 1500
  .ColWidth(ColQtyMachine) = 1500
  .ColWidth(colQtyProcess) = 1500
  .ColWidth(ColQtyHour) = 1500
  .ColWidth(ColQtyDay) = 1500
    
  
    
  .Cell(flexcpAlignment, 0, 0, 0, ColQtyDay) = flexAlignCenterCenter
  '.ColHidden(lb_Item) = True
  .EditMaxLength = 1
 End With
End Sub
Private Sub browse()
 MousePointer = vbHourglass
 'Call header
 If cboFactori = "" Or cboLineCode = "" Or cboMachine = "" Then
 KosongBawah
 Exit Sub
 MousePointer = vbDefault
 End If
 
 Sql = "SELECT *,ISNULL((SELECT  Item_Name FROM Item_Master WHERE Item_code=a.Item_Code),'') Item_Name "
 Sql = Sql & " FROM Machine_Capacity a WHERE Factory_Code='" & cboFactori & "' AND LINE_Code ='" & cboLineCode & "'"
 Sql = Sql & " AND Machine_Code='" & cboMachine & "'"
  
 If rs.State <> adStateClosed Then rs.Close
 rs.Open Sql, Db, adOpenDynamic, adLockOptimistic, adCmdText
    With Grid
    .Rows = 1
    Do While Not rs.EOF
      .Rows = .Rows + 1
      .TextMatrix(.Rows - 1, 0) = " "
      .TextMatrix(.Rows - 1, ColItem) = Trim(rs!Item_Code)
      .TextMatrix(.Rows - 1, ColDesc) = is_null((rs!item_name))
      .TextMatrix(.Rows - 1, ColCycleTime) = Trim(rs!Cycle_Time)
      .TextMatrix(.Rows - 1, ColEficiency) = Trim(rs!Efficiency)
      .TextMatrix(.Rows - 1, colQtyProcess) = Trim(rs!Qty_Process)
      .TextMatrix(.Rows - 1, ColQtyMachine) = Trim(rs!Qty_Machine)
      .TextMatrix(.Rows - 1, ColQtyHour) = GetQtyHour(.Rows - 1)
      .TextMatrix(.Rows - 1, ColQtyDay) = GetQtyDay(.Rows - 1)
      
      .Cell(flexcpBackColor, .Rows - 1, 0) = &HFFFFFF
      rs.MoveNext
     
    Loop
    
    End With
Set rs = Nothing
MousePointer = vbDefault
End Sub
Function is_null(Data)
is_null = Trim(IIf(IsNull(Data), "", Data))

End Function
Function GetQtyDay(baris)
With Grid
GetQtyDay = 7 * .TextMatrix(baris, ColQtyHour)
End With
End Function

Function GetQtyHour(baris)
With Grid
GetQtyHour = (3600 / .TextMatrix(baris, ColCycleTime)) * (.TextMatrix(baris, ColEficiency) / 100) * (.TextMatrix(baris, ColQtyMachine)) * (.TextMatrix(baris, colQtyProcess))
End With
End Function
Private Sub AdCmbFactory()
'*****cboFactiory******
Call isiCbo(cboFactori, "Trade_Master", "Trade_Code", "Trade_Name", 50, 100, "Trade_Code,Trade_Name", , , "(Trade_Cls = 1)", , 0)
End Sub
Private Sub adtocomboitem()

With cboItem
    .clear
    .ColumnCount = 2
     Sql = "SELECT  Item_Code,Item_Name FroM ITEM_Master"
    .AddItem ""
    .List(.ListCount - 1, 0) = "Common"
    .List(.ListCount - 1, 1) = "Common"

    rs.Open Sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    While Not rs.EOF
        .AddItem ""
        .List(.ListCount - 1, 0) = Trim(rs.Fields("item_code"))
        .List(.ListCount - 1, 1) = Trim(rs.Fields("item_name"))
        
        rs.MoveNext
    Wend
    rs.Close
    .ListWidth = 300
    .ColumnWidths = "75 pt;100"
    .Text = ""
    .ListIndex = -1
 End With

Set rs = Nothing
End Sub
Private Sub AdToComboMachineNo()

With cboMachineNo
    .clear
    .ColumnCount = 2
    
    If cboItem.MatchFound = True Then
        Sql = "select ml.line_code, ml.line_name " & _
              "from Manufacture_Line ml " & _
              "where ml.Manufacture_Code = '" & Trim(cboItem.Column(3)) & "' "
                     
         rs.Open Sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
        
        While Not rs.EOF
            .AddItem ""
            .List(.ListCount - 1, 0) = Trim(rs.Fields("line_code"))
            .List(.ListCount - 1, 1) = Trim(rs.Fields("line_name"))
            rs.MoveNext
        Wend
        rs.Close
        
        .ListWidth = 150
        .ColumnWidths = "50 pt;100 pt"
        .Text = ""
    End If
    
End With

Set rs = Nothing
End Sub
Private Sub CboItem_Change()
    If cboItem.MatchFound Then
        LblItemName.Caption = cboItem.Column(1)
        LblErrMsg.Caption = ""
    Else
        LblItemName.Caption = ""
        'If cboItem.Text <> "" Then LblErrMsg = DisplayMsg(8084)
    End If
    
    
End Sub
Private Sub CboGroup_Change()
    If cbogroup.MatchFound Then
        lblGroup.Caption = cbogroup.Column(1)
        LblErrMsg.Caption = ""
    Else
        lblGroup.Caption = ""
        LblErrMsg.Caption = ""
        If cbogroup.Text <> "" Then LblErrMsg = DisplayMsg(8083)
    End If
    adtocomboitem
End Sub
 Function Get_Field(Sql, Field)
Dim Rdata As New ADODB.Recordset
Set Rdata = Db.Execute(Sql)
Get_Field = ""
If Not Rdata.EOF Then
 Get_Field = IIf(IsNull(Rdata.Fields(Field)), "", Rdata.Fields(Field))
End If
End Function
Private Sub cboFactori_Change()
 If cboFactori.MatchFound Then
       lblFactoryCode.Caption = cboFactori.List(cboFactori.ListIndex, 1)
       LblPesan = ""
        If CDbl(Get_Field("SELECT count(*) A FROM manufacture_line WHERE Manufacture_Code='" & cboFactori & "' and Line_code='" & cboLineCode & "'", "A")) = 0 Then
            Call isiCbo(cboLineCode, "manufacture_line", "Line_Code", "Line_Name", 50, 100, "Line_Code,Line_name", , , " Manufacture_Code='" & cboFactori & "'", , 0)
        End If
    If cboFactori.List(cboFactori.ListIndex, 0) <> cboFactori.Text Then Call browse
    Else
       lblFactory = ""
       LblPesan = DisplayMsg(4014) '"Location CD is not found !"
    End If
    MousePointer = vbDefault
End Sub

Private Sub cboFactori_Click()
cboFactori_Change
End Sub

Private Sub cboFactori_DropButtonClick()
cboFactori_Change
End Sub

Private Sub cboFactori_GotFocus()
cboFactori_Change
End Sub

Private Sub cboLineCode_Change()
MousePointer = vbHourglass
If cboLineCode.MatchFound Then
       lblLineName.Caption = cboLineCode.List(cboLineCode.ListIndex, 1)
       LblPesan = ""
       Call browse
    Else
       lblLineName = ""
       LblPesan = DisplayMsg(4014) '"Location CD is not found !"
    End If
    MousePointer = vbDefault
End Sub

Private Sub cboLineCode_DropButtonClick()
cboLineCode_Change
End Sub

Private Sub cboMachine_Change()
    If cboMachine.MatchFound Then
        lblMachine.Caption = cboMachine.Column(1)
        LblErrMsg.Caption = ""
        Call browse
    Else
        lblMachine.Caption = ""
        LblErrMsg.Caption = ""
        If cboMachine.Text <> "" Then LblErrMsg = DisplayMsg(4017)
    End If

End Sub

Private Sub cboMachine_DropButtonClick()
cboMachine_Change
End Sub

Private Sub AboveGridAreaClear()
 cbogroup.Text = ""
 lblGroup.Caption = ""
 cboItem.Text = ""
 lblItem.Caption = ""
 lblFactory.Caption = ""
End Sub
Private Sub BelowGridAreaClear()
 lblMachineName.Caption = ""

End Sub
Private Function InsertValidation() As Boolean
 InsertValidation = True
 If cbogroup.Text = "" Then
  LblErrMsg.Caption = DisplayMsg("8081") 'Please select Group Cls!
  cbogroup.SetFocus
  InsertValidation = False
  Exit Function
 ElseIf cbogroup.MatchFound = False Then
  LblErrMsg.Caption = DisplayMsg("8083") 'Record with This Group Cls not Found
  cbogroup.SetFocus
  InsertValidation = False
  Exit Function
 ElseIf cboItem.Text = "" Then
  LblErrMsg.Caption = DisplayMsg("8082") 'Please select Item Code!
  cboItem.SetFocus
  InsertValidation = False
  Exit Function
 ElseIf cboItem.MatchFound = False Then
  LblErrMsg = DisplayMsg("8084") 'Record with this Item Code Not Found
  cboItem.SetFocus
  InsertValidation = False
  Exit Function
 ElseIf cboMachineNo.Text = "" Then
  LblErrMsg.Caption = DisplayMsg("1041") 'Please input Line Code!
  cboMachineNo.SetFocus
  InsertValidation = False
  Exit Function
 ElseIf cboMachineNo.MatchFound = False Then
  LblErrMsg.Caption = DisplayMsg("4017") 'Record with This Line Code not found
  cboMachineNo.SetFocus
  InsertValidation = False
  Exit Function
 ElseIf txtCapacity.Text = "" Then
  LblErrMsg.Caption = DisplayMsg("8114") 'Please Input Capacity !
  txtCapacity.SetFocus
  InsertValidation = False
  Exit Function
 Else
  Sql = "Select * from Machine_Capacity mc " & _
        "Where mc.Item_Code = '" & Trim(cboItem.Text) & "' " & _
        "And mc.Line_Code = '" & Trim(cboMachineNo.Text) & "' "
 
  If rscek.State <> adStateClosed Then rscek.Close
  rscek.Open Sql, Db, adOpenDynamic, adLockReadOnly, adCmdText
 
  If rscek.EOF Then
   LblErrMsg.Caption = ""
   InsertValidation = True
  Else
   LblErrMsg.Caption = DisplayMsg("1023") 'Data already exist
   InsertValidation = False
   Exit Function
  End If
  Set rscek = Nothing
 End If
End Function
Private Function UpdateValidation() As Boolean
 UpdateValidation = True
 If cboMachineNo.Text = "" Then
  LblErrMsg.Caption = DisplayMsg("1041") 'Please input Line Code!
  cboMachineNo.SetFocus
  UpdateValidation = False
  Exit Function
 ElseIf cboMachineNo.MatchFound = False Then
  LblErrMsg.Caption = DisplayMsg("4017") 'Record with This Line Code not found
  cboMachineNo.SetFocus
  UpdateValidation = False
  Exit Function
 End If

 If Trim(Grid.TextMatrix(Grid.Row, lb_MachineNo)) = Trim(cboMachineNo.Text) Then Exit Function

 Sql = "Select * from Machine_Capacity mc " & _
       "Where mc.Item_Code = '" & Trim(Grid.TextMatrix(Grid.Row, lb_Item)) & "' " & _
       "And mc.Line_Code = '" & Trim(cboMachineNo.Text) & "' "
 
 If rscek.State <> adStateClosed Then rscek.Close
 rscek.Open Sql, Db, adOpenDynamic, adLockReadOnly, adCmdText
 
 If rscek.EOF Then
  LblErrMsg.Caption = ""
  UpdateValidation = True
 Else
  LblErrMsg.Caption = DisplayMsg("1023") 'Data already exist
  UpdateValidation = False
  Exit Function
 End If
 
 Set rscek = Nothing
 
 If txtCapacity.Text = "" Then
  LblErrMsg.Caption = DisplayMsg("8114") 'Please Input Capacity !
  txtCapacity.SetFocus
  UpdateValidation = False
 End If
End Function

Private Sub cmdBrowser_Click()
End Sub

Private Sub CmdCancel_Click()
 Call ClearS
 Call BelowGridAreaClear
 LblErrMsg.Caption = ""
End Sub

Private Sub cmdClear_Click()
Call ClearS
Call AboveGridAreaClear
Call BelowGridAreaClear
LblErrMsg.Caption = ""
cbogroup.SetFocus
End Sub

Private Sub CmdSubmit_Click()
Dim rddata As New ADODB.Recordset
Set rddata = Nothing
Me.MousePointer = vbHourglass
If hakUpdate(Me.Name) = 0 Then LblPesan = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub

''DELETE  dATA
With Grid
For i = 1 To Grid.Rows - 1
 If .TextMatrix(i, 0) = "D" Then
        LblInput = MsgBox("Do you really to delete  ?", _
                    vbYesNo + vbQuestion, "Confirmation")
        If LblInput = vbYes Then
                        Sql = " DELETE FROM dbo.Machine_Capacity WHERE Factory_Code ='" & cboFactori & "'"
                        Sql = Sql & "AND Line_Code='" & cboLineCode & "' AND Machine_Code='" & cboMachine & "' and Item_Code='" & .TextMatrix(i, ColItem) & "'"
                        Db.Execute Sql
                        LblErrMsg = DisplayMsg(1201)
                        Call browse
           End If
 
 MousePointer = vbDefault
 Exit Sub
 End If
Next
End With

'---- insert ATAU update data---------------
'mengecek kelengkapan data
If cboFactori = "" Then
    LblErrMsg = DisplayMsg(1040)
    cboFactori.SetFocus
    Me.MousePointer = vbDefault
Exit Sub
ElseIf cboLineCode = "" Then
    LblErrMsg = DisplayMsg(1041)
    cboLineCode.SetFocus
    MousePointer = vbDefault
    Exit Sub
ElseIf cboMachine = "" Then
    LblErrMsg = DisplayMsg(8018)
    cboMachine.SetFocus
    MousePointer = vbDefault
    Exit Sub
    
ElseIf cboItem = "" Then
    LblErrMsg = DisplayMsg(8082)
    cboItem.SetFocus
    MousePointer = vbDefault
    Exit Sub
ElseIf txtCycle = "" Then
    LblErrMsg = DisplayMsg(8125)
    txtCycle.SetFocus
    MousePointer = vbDefault
    Exit Sub
 ElseIf txtEficiency = "" Then
    LblErrMsg = DisplayMsg(8126)
    txtEficiency.SetFocus
    MousePointer = vbDefault
    Exit Sub
 ElseIf txtQtyMachine = "" Then
    LblErrMsg = DisplayMsg(8127)
    txtQtyMachine.SetFocus
    MousePointer = vbDefault
    Exit Sub
  ElseIf txtQtyProcess = "" Then
    LblErrMsg = DisplayMsg(8128)
    txtQtyProcess.SetFocus
    MousePointer = vbDefault
    Exit Sub
End If

    Dim Ritem As New ADODB.Recordset
    Dim sItem As String
    sItem = "SELECT * FROM Item_Master"
    If LCase(cboItem.Text) <> "common" Then sItem = sItem & " WHERE Item_Code='" & cboItem & "'"
    
    Ritem.Open sItem, Db, adOpenDynamic, adLockBatchOptimistic
    While Not Ritem.EOF
        Set rddata = Nothing
        Sql = " select * FROM dbo.Machine_Capacity WHERE Factory_Code ='" & cboFactori & "'"
        Sql = Sql & "AND Line_Code='" & cboLineCode & "' AND Machine_Code='" & cboMachine & "'"
        Sql = Sql & " AND Item_Code='" & Trim(Ritem!Item_Code) & "'"
        rddata.Open Sql, Db, 1, 3
        If rddata.EOF Then
        'tambah data
            rddata.AddNew
            rddata!Factory_Code = cboFactori
            rddata!Line_Code = cboLineCode
            rddata!Machine_code = cboMachine
        End If
        rddata!Item_Code = Trim(Ritem!Item_Code)
        rddata!Cycle_Time = txtCycle
        rddata!Efficiency = txtEficiency
        rddata!Qty_Machine = txtQtyMachine
        rddata!Qty_Process = txtQtyProcess
        rddata.update
        rddata.Close
      Ritem.MoveNext
     Wend




    
 
 KosongBawah
 LblErrMsg = DisplayMsg(1101)
 browse
'Call browse
Me.MousePointer = vbDefault
End Sub
Sub KosongBawah()
'cboItem.clear
cboItem = ""
txtCycle = "": txtEficiency = "": TxtQtyDay = "": TxtQtyHour = "": txtQtyMachine = "": txtQtyProcess = ""
LblErrMsg = ""
End Sub
Private Sub InsertData()
On Error GoTo ErrHandler
If InsertValidation = False Then Exit Sub

Sql = "Insert Into Machine_Capacity " & _
      vbLf & "(Item_Code, Line_Code, Capacity, Last_Update, Last_User, Register_Date) " & _
      vbLf & "Values( " & _
      vbLf & "'" & Trim(cboItem.Text) & "', " & _
      vbLf & "'" & Trim(cboMachineNo.Text) & "', " & _
      vbLf & CDbl(txtCapacity.Text) & ", " & _
      vbLf & "getdate(), " & _
      vbLf & "'" & Trim(userLogin) & "', " & _
      vbLf & "getdate() " & _
      vbLf & ") "

Db.BeginTrans
Db.Execute Sql
Db.CommitTrans
Call BelowGridAreaClear
LblErrMsg.Caption = DisplayMsg("1000") 'Data Saved Success!

Exit Sub
ErrHandler:
    Db.RollbackTrans
    LblErrMsg.Caption = "[" & Err.number & "] " & Err.Description
    Err.clear
End Sub
Private Sub DeleteData()
On Error GoTo ErrHandler
Dim i As Integer
Dim Certainty As Integer

Certainty = MsgBox("Do you really want to delete the record(s) ?", vbYesNo + vbQuestion, "Confirmation")

If Certainty = vbYes Then
 For i = 1 To Grid.Rows - 1
  If Trim(Grid.TextMatrix(i, lb_Action)) = "D" Then
   Sql = "Delete From Machine_Capacity " & _
         vbLf & "Where Item_Code = '" & Trim(Grid.TextMatrix(i, lb_Item)) & "' " & _
         vbLf & "And Line_Code = '" & Trim(Grid.TextMatrix(i, lb_MachineNo)) & "' "
   Db.BeginTrans
   Db.Execute Sql
   Db.CommitTrans
  End If
 Next i
Else
 Call ClearS
End If

Db.BeginTrans
Db.Execute Sql
Db.CommitTrans
Call BelowGridAreaClear
LblErrMsg.Caption = DisplayMsg("1201") 'Delete Record Success!

Exit Sub
ErrHandler:
    Db.RollbackTrans
    LblErrMsg.Caption = "[" & Err.number & "] " & Err.Description
    Err.clear
End Sub

Private Sub combobox1_Change()

End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Grid.Col > lb_Action Then Cancel = True
End Sub

Private Sub grid_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Call Grid_KeyPressEdit(Grid.RowSel, Grid.ColSel, KeyAscii)
End Sub

Private Sub Grid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Grid.Col = lb_Action Then
        If KeyAscii = 8 Then Grid.Text = ""
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii <> Asc("D") And KeyAscii <> Asc("S") And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyReturn Then
            KeyAscii = 0
        End If
    End If
    
End Sub
Private Sub ClearS(Optional Letter As String)
    Dim i As Integer
    Grid.Col = lb_Action
    If Letter <> "" Then
        For i = 1 To Grid.Rows - 1
            Grid.Row = i
            If Grid.Text = Letter Then Grid.Text = ""
        Next i
    Else
        For i = 1 To Grid.Rows - 1
            Grid.Row = i
            Grid.Text = ""
        Next i
    End If
End Sub
Private Sub grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim TextGrid As String, Data1 As String, rtrec As Integer, i As Integer
    
    TextGrid = Trim(Grid.Text)
    If TextGrid = "S" Then
     cboItem = Trim(Grid.TextMatrix(Row, ColItem))
     LblItemName = Trim(Grid.TextMatrix(Row, ColDesc))
     txtCycle = Trim(Grid.TextMatrix(Row, ColCycleTime))
     txtQtyMachine = Trim(Grid.TextMatrix(Row, ColQtyMachine))
     txtQtyProcess = Trim(Grid.TextMatrix(Row, colQtyProcess))
     TxtQtyHour = Trim(Grid.TextMatrix(Row, ColQtyHour))
     TxtQtyDay = Trim(Grid.TextMatrix(Row, ColQtyDay))
     txtEficiency = Trim(Grid.TextMatrix(Row, ColEficiency))
     Tekan = 1
     Call ClearS
     
    ElseIf TextGrid = "D" Then
     Tekan = 2
     Call ClearS("S")
     Call KosongBawah
    Else
     Tekan = 3
     Call BelowGridAreaClear
    End If

    Grid.TextMatrix(Row, Col) = TextGrid
    Grid.Col = Col
    Grid.Row = Row
End Sub
Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
    If ErrMsg = "" Then
        Unload Me
    Else
        LblPesan.Caption = ErrMsg
    End If
End Sub
Private Sub cmdSubMenu_Click()
 DoEvents
 frmMainMenu.Show
 DoEvents
 Unload Me
End Sub
Private Sub Form_Load()
 If gb_Simulation = True Then Call up_InitSimulation(Me)
    
 CtrlMenu1.FormName = Me.Name
 Me.Caption = "Machine Capacity (Item Code)"
 Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"
 lblFactoryCode = "": lblLineName = "": lblMachine = "": LblItemName = ""
 Call header
 Call AdCmbFactory
 Call adToLineCode
 Call adToCboMachine
  Call adtocomboitem
 Call browse
 Tekan = 3 'Insert directly
 MousePointer = vbDefault
End Sub
Sub adToCboMachine()
Call isiCbo(cboMachine, "Machine_master", "Machine_Code", "Machine_Name", 50, 100, "Machine_code,Machine_name", , , , , 0)
End Sub
Sub adToLineCode()
Call isiCbo(cboLineCode, "manufacture_line", "Line_Code", "Line_Name", 50, 100, "Line_Code,Line_name", , , , , 0)
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub txtCapacity_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc("'") Then
 KeyAscii = 0
ElseIf KeyAscii = Asc(".") Then
 If InStr(1, txtCapacity.Text, ".") > 0 Then KeyAscii = 0
ElseIf IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn And KeyAscii <> Asc(".") And KeyAscii <> Asc(",") Then
 KeyAscii = 0
Else
 If IsNumeric(Chr(KeyAscii)) = True Then
  If CDbl((txtCapacity.Text & Chr(KeyAscii))) < 10000000000000# Then      '18,5 (tp 2 digit decimal)
    'If InStr(1, txtCapacity.Text, ".") <= 0 Then Exit Sub
    'If Len(Right(txtCapacity.Text, (Len(txtCapacity.Text) - InStr(1, txtCapacity.Text, ".")))) >= 2 Then
     'txtCapacity.Text = Left(txtCapacity, InStr(1, txtCapacity.Text, ".") + 2)
    'End If
  Else
   KeyAscii = 0
  End If
 End If
End If

End Sub
Private Sub txtCapacity_Change()
 If InStr(1, txtCapacity.Text, ",") = 1 Then txtCapacity.Text = Right(txtCapacity, Len(txtCapacity) - 1)
End Sub
Private Sub txtCapacity_LostFocus()
 txtCapacity.Text = Format(txtCapacity.Text, gs_formatQty)
End Sub

Private Sub txtCycle_Change()
On Error Resume Next
TxtQtyHour = (3600 / IfNol(txtCycle)) * (IfNol(txtEficiency) / 100) * (IfNol(txtQtyMachine)) * (IfNol(txtQtyProcess))
End Sub

Private Sub txtCycle_KeyPress(KeyAscii As Integer)

If InStr("0123456789", Chr(KeyAscii)) = 0 Then
    If KeyAscii = 13 Then
      SendKeys vbTab
    ElseIf KeyAscii <> vbKeyBack Then
    KeyAscii = 0
    End If
End If
If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtEficiency_Change()
On Error Resume Next
TxtQtyHour = (3600 / IfNol(txtCycle)) * (IfNol(txtEficiency) / 100) * (IfNol(txtQtyMachine)) * (IfNol(txtQtyProcess))
End Sub

Private Sub txtEficiency_KeyPress(KeyAscii As Integer)
If InStr("0123456789", Chr(KeyAscii)) = 0 Then
    If KeyAscii = 13 Then
      SendKeys vbTab
    ElseIf KeyAscii <> vbKeyBack Then
    KeyAscii = 0
    End If
End If
If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub TxtQtyHour_Change()
On Error Resume Next
TxtQtyDay = 7 * TxtQtyHour
End Sub

Private Sub txtQtyMachine_Change()
On Error Resume Next
TxtQtyHour = (3600 / IfNol(txtCycle)) * (IfNol(txtEficiency) / 100) * (IfNol(txtQtyMachine)) * (IfNol(txtQtyProcess))
End Sub

Private Sub txtQtyMachine_KeyPress(KeyAscii As Integer)
If InStr("0123456789", Chr(KeyAscii)) = 0 Then
    If KeyAscii = 13 Then
      SendKeys vbTab
    ElseIf KeyAscii <> vbKeyBack Then
    KeyAscii = 0
    End If
End If
If KeyAscii = 39 Then KeyAscii = 0
End Sub
Function IfNol(Data)

IfNol = IIf(Data = "", 0, Data)
End Function

Private Sub txtQtyProcess_Change()
On Error Resume Next
TxtQtyHour = (3600 / IfNol(txtCycle)) * (IfNol(txtEficiency) / 100) * (IfNol(txtQtyMachine)) * (IfNol(txtQtyProcess))
End Sub

Private Sub txtQtyProcess_KeyPress(KeyAscii As Integer)
If InStr("0123456789", Chr(KeyAscii)) = 0 Then
    If KeyAscii = 13 Then
      SendKeys vbTab
    ElseIf KeyAscii <> vbKeyBack Then
    KeyAscii = 0
    End If
End If
If KeyAscii = 39 Then KeyAscii = 0
End Sub
