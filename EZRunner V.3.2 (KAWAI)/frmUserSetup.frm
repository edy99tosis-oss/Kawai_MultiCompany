VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmUserSetup 
   BackColor       =   &H00FDDFE3&
   Caption         =   "User Setup"
   ClientHeight    =   10365
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   15120
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUserSetup.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10365
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   12923
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   450
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   714
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3990
      Left            =   4920
      TabIndex        =   23
      Top             =   1140
      Width           =   9840
      Begin VB.TextBox txtPO 
         BackColor       =   &H00FFFFFF&
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
         Left            =   2280
         MaxLength       =   1
         TabIndex        =   32
         Tag             =   "1"
         Text            =   "Text2"
         Top             =   3480
         Width           =   315
      End
      Begin VB.TextBox txtPass2 
         BackColor       =   &H00FFFFFF&
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
         IMEMode         =   3  'DISABLE
         Left            =   2280
         MaxLength       =   18
         PasswordChar    =   "#"
         TabIndex        =   4
         Tag             =   "1"
         Text            =   "Text3"
         Top             =   1500
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FDDFE3&
         Caption         =   "Unlocked"
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
         Left            =   2280
         TabIndex        =   8
         Tag             =   "1"
         Top             =   2790
         Width           =   1335
      End
      Begin VB.TextBox txtDesc 
         BackColor       =   &H00FFFFFF&
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
         Left            =   2280
         MaxLength       =   25
         TabIndex        =   9
         Tag             =   "1"
         Text            =   "Text2"
         Top             =   3120
         Width           =   3855
      End
      Begin VB.TextBox txtUser 
         BackColor       =   &H00FFFFFF&
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
         Left            =   2280
         MaxLength       =   15
         TabIndex        =   1
         Tag             =   "1"
         Text            =   "AAAAAAAAAAAAAAA"
         Top             =   330
         Width           =   2295
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H00FFFFFF&
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
         Left            =   2280
         MaxLength       =   25
         TabIndex        =   2
         Tag             =   "1"
         Text            =   "Text2"
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txtPass1 
         BackColor       =   &H00FFFFFF&
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
         IMEMode         =   3  'DISABLE
         Left            =   2280
         MaxLength       =   18
         PasswordChar    =   "#"
         TabIndex        =   3
         Tag             =   "1"
         Text            =   "Text3"
         Top             =   1110
         Width           =   1455
      End
      Begin VB.Frame Frame3 
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
         Height          =   465
         Left            =   2280
         TabIndex        =   16
         Top             =   2220
         Width           =   2535
         Begin VB.OptionButton optStatus 
            BackColor       =   &H00FDDFE3&
            Caption         =   "No"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   1380
            TabIndex        =   7
            Top             =   150
            Value           =   -1  'True
            Width           =   780
         End
         Begin VB.OptionButton optStatus 
            BackColor       =   &H00FDDFE3&
            Caption         =   "Yes "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   135
            TabIndex        =   6
            Top             =   150
            Width           =   735
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Left            =   2100
         TabIndex        =   42
         Top             =   3540
         Width           =   75
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Index           =   19
         Left            =   2100
         TabIndex        =   41
         Top             =   3165
         Width           =   75
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Left            =   2100
         TabIndex        =   40
         Top             =   2790
         Width           =   75
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Left            =   2100
         TabIndex        =   39
         Top             =   2340
         Width           =   75
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Index           =   16
         Left            =   2100
         TabIndex        =   38
         Top             =   1950
         Width           =   75
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Index           =   15
         Left            =   2100
         TabIndex        =   37
         Top             =   1560
         Width           =   75
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Index           =   14
         Left            =   2100
         TabIndex        =   36
         Top             =   1155
         Width           =   75
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Index           =   13
         Left            =   2100
         TabIndex        =   35
         Top             =   765
         Width           =   75
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Index           =   12
         Left            =   2100
         TabIndex        =   34
         Top             =   375
         Width           =   75
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PO Initial"
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
         Index           =   11
         Left            =   360
         TabIndex        =   33
         Top             =   3540
         Width           =   780
      End
      Begin MSForms.ComboBox cboUser 
         Height          =   315
         Left            =   2280
         TabIndex        =   5
         Top             =   1890
         Width           =   2295
         VariousPropertyBits=   746604571
         MaxLength       =   7
         DisplayStyle    =   7
         Size            =   "4048;556"
         ColumnCount     =   2
         ListRows        =   20
         MatchEntry      =   0
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Group "
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
         TabIndex        =   31
         Top             =   1950
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password "
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
         TabIndex        =   30
         Top             =   1560
         Width           =   1620
      End
      Begin VB.Label Label1 
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
         Index           =   7
         Left            =   360
         TabIndex        =   29
         Top             =   3165
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Lock "
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
         TabIndex        =   28
         Top             =   2790
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
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
         TabIndex        =   27
         Top             =   375
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
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
         TabIndex        =   26
         Top             =   1155
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Full Name "
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
         Left            =   360
         TabIndex        =   25
         Top             =   765
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status Admin"
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
         Index           =   10
         Left            =   360
         TabIndex        =   24
         Top             =   2340
         Width           =   1140
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Delete"
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
      Left            =   12683
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   9840
      Width           =   1000
   End
   Begin VB.CommandButton cmdSubMenu 
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
      Left            =   413
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   9840
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
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
      Left            =   13763
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9840
      Width           =   1000
   End
   Begin VB.CommandButton Command1 
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
      Left            =   11603
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9840
      Width           =   1000
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
      Height          =   585
      Left            =   413
      TabIndex        =   20
      Top             =   9120
      Width           =   14355
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
         Height          =   270
         Left            =   105
         TabIndex        =   21
         Top             =   195
         Width           =   14130
      End
   End
   Begin MSComctlLib.ListView lvw1 
      Height          =   3915
      Left            =   420
      TabIndex        =   0
      Top             =   1215
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   6906
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "UserName"
         Object.Width           =   2866
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Full Name"
         Object.Width           =   4718
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   3555
      Left            =   420
      TabIndex        =   10
      Top             =   5535
      Width           =   14355
      _cx             =   25321
      _cy             =   6271
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
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   275
      RowHeightMax    =   275
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
      Editable        =   2
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
   Begin VB.Label lblNama 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name            :"
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
      Left            =   2910
      TabIndex        =   22
      Top             =   5220
      Width           =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Properties For User Name :"
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
      Left            =   420
      TabIndex        =   19
      Top             =   5235
      Width           =   2340
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "User Setup"
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
      Index           =   0
      Left            =   413
      TabIndex        =   18
      Top             =   450
      Width           =   14355
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Available User :"
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
      Left            =   413
      TabIndex        =   17
      Top             =   930
      Width           =   1395
   End
End
Attribute VB_Name = "frmUserSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsUser As New ADODB.Recordset
Dim RS As New ADODB.Recordset
Dim sql As String
Dim ubah As Boolean
Dim i As Integer
Dim pilihAdmin As Integer

Dim bteColNo As Byte
Dim bteColSubMenu As Byte
Dim bteColMenuID As Byte
Dim bteColMenuName As Byte
Dim bteColAccess As Byte
Dim bteColUpdate As Byte
Dim bteColPrice As Byte

Private Sub headerGrid()

    bteColNo = 0
    bteColSubMenu = 1
    bteColMenuID = 2
    bteColMenuName = 3
    bteColAccess = 4
    bteColUpdate = 5
    bteColPrice = 6
    
    With Grid
        .clear
        .ColS = 7
        .Rows = 1
        
        .TextMatrix(0, bteColNo) = "No"
        .TextMatrix(0, bteColSubMenu) = "Sub Menu"
        .TextMatrix(0, bteColMenuID) = "Menu ID"
        .TextMatrix(0, bteColMenuName) = "Menu Name"
        .TextMatrix(0, bteColAccess) = "Access"
        .TextMatrix(0, bteColUpdate) = "Update"
        .TextMatrix(0, bteColPrice) = "Display Price"
        
        .ColWidth(bteColNo) = 500
        .ColWidth(bteColSubMenu) = 2000
        .ColWidth(bteColMenuID) = 1000
        .ColWidth(bteColMenuName) = 4000
        .ColWidth(bteColAccess) = 1000
        .ColWidth(bteColUpdate) = 1000
        .ColWidth(bteColPrice) = 1500
        
        .ColAlignment(bteColNo) = flexAlignLeftCenter
        .ColAlignment(bteColSubMenu) = flexAlignLeftCenter
        .ColAlignment(bteColMenuID) = flexAlignCenterCenter
        .ColAlignment(bteColMenuName) = flexAlignLeftCenter
        .ColAlignment(bteColAccess) = flexAlignCenterCenter
        .ColAlignment(bteColUpdate) = flexAlignCenterCenter
        .ColAlignment(bteColPrice) = flexAlignCenterCenter
        
        .Cell(flexcpChecked, bteColNo, bteColAccess) = flexUnchecked
        .Cell(flexcpChecked, bteColNo, bteColUpdate) = flexUnchecked
        .Cell(flexcpChecked, bteColNo, bteColPrice) = flexUnchecked
    End With
End Sub

Sub Kosong()
    txtUser = ""
    txtUser.Enabled = True
    txtName = ""
    txtPass1 = ""
    txtPass2 = ""
    cboUser = ""
    txtDesc = ""
    txtpo = ""
    Check1.Value = 0
    optStatus(1).Value = True
    ubah = False
    lblNama = ""
    Call isiList
    Call isiCbo
    LblErrMsg = ""
End Sub

Sub isiCbo() 'Isi Combo User Group (spy privilege nya sama dgn user yg dipilih)
    rsUser.filter = ""
    rsUser.Requery
    
    cboUser.clear
    i = 0
    Do While Not rsUser.EOF
        cboUser.AddItem ""
        cboUser.List(i, 0) = Trim(rsUser("userName"))
        cboUser.List(i, 1) = Trim(rsUser("Name"))
        i = i + 1
        rsUser.MoveNext
    Loop
    cboUser.ListWidth = 200
    cboUser.ColumnWidths = "75pt;100pt"
End Sub

Private Sub isiList() 'Isi List View
Dim aItem As ListItem
Dim kdIsiList As String
Dim nmUser As String

    lvw1.ListItems.clear
    rsUser.filter = ""
    rsUser.Requery
    If Not rsUser.EOF Then
        Do While Not rsUser.EOF
            nmUser = Trim(rsUser("username"))
            kdIsiList = "n" & nmUser
            Set aItem = lvw1.ListItems.Add(, kdIsiList, nmUser)
            aItem.SubItems(1) = Trim(rsUser("name"))
            rsUser.MoveNext
        Loop
    End If
End Sub

Sub IsiGrid(Optional lama As Integer, Optional teks As String)
Dim rsPriv As New ADODB.Recordset

    Call headerGrid
            
    If lama = 1 Then
'        sql = "select distinct b.Menu_Id ,menu_Desc ," & _
'            "isnull(status,0) as status, isnull(Allow_Update,0) as statusU, isnull(Allow_Price,0) as statusP," & _
'            "Group_ID,Menu_Indeks " & _
'            "from user_Privilege a right join user_Menu b " & _
'            " on a.App_ID = 'P01' and a.Menu_Id = b.Menu_Id " & _
'            "where (a.UserName ='" & teks & _
'            "' and b.Group_ID <> 'Security System') or  (b.Menu_ID not in (select distinct Menu_ID From user_Privilege) and b.Group_ID <> 'Security System') "

 sql = "DECLARE @User Char(15)='" & teks & "'" & vbCrLf & _
        "SELECT Menu_ID,Menu_Desc," & vbCrLf & _
        " [Status]=Coalesce((select coalesce(Status,0) From User_Privilege where UserName=@User and Menu_ID=A.Menu_ID),0), " & vbCrLf & _
        " statusU=Coalesce((select coalesce(Allow_Update,0) From User_Privilege where UserName=@User and Menu_ID=A.Menu_ID),0), " & vbCrLf & _
        " statusP=Coalesce((select coalesce(Allow_Price,0) From User_Privilege where UserName=@User and Menu_ID=A.Menu_ID),0), " & vbCrLf & _
        " Group_ID , Menu_Indeks " & vbCrLf & _
        " FROM User_Menu A " & vbCrLf & _
        "where Group_ID <> 'Security System' order by   Menu_ID,menu_Indeks " & vbCrLf
    Else
        sql = "select distinct b.Menu_Id ,menu_Desc ," & _
            "0 as status, 0 as statusU, 0 as statusP, Group_ID,Menu_Indeks " & _
            "from user_Privilege a right join user_Menu b " & _
            "on a.App_ID = 'P01' and a.Menu_Id =b.Menu_Id " & _
            "where b.Group_ID <> 'Security System' "
            
            sql = sql & "order by b.Menu_ID,menu_Indeks"
    End If
    
    
    
    If rsPriv.State <> adStateClosed Then rsPriv.Close
    rsPriv.CursorLocation = adUseClient
    rsPriv.Open sql, Db, adOpenDynamic, adLockOptimistic
    
With Grid
    For i = 1 To rsPriv.RecordCount
        DoEvents
        .Rows = .Rows + 1
        .TextMatrix(i, bteColNo) = " " & i
        .TextMatrix(i, bteColSubMenu) = Trim(rsPriv("Group_ID"))
        .MergeCol(bteColSubMenu) = True
        .MergeCells = flexMergeRestrictColumns
        
        .TextMatrix(i, bteColMenuID) = Trim(rsPriv("Menu_ID"))
        .TextMatrix(i, bteColMenuName) = Trim(rsPriv("Menu_Desc"))
        .Cell(flexcpChecked, i, bteColAccess) = IIf(rsPriv("status") = "0", flexUnchecked, flexChecked)
        .Cell(flexcpChecked, i, bteColUpdate) = IIf(rsPriv("statusU") = "0", flexUnchecked, flexChecked)
        .Cell(flexcpChecked, i, bteColPrice) = IIf(rsPriv("statusP") = "0", flexUnchecked, flexChecked)
        
        rsPriv.MoveNext
    Next i
    Call cekSelect(CLng(bteColAccess))  'utk col 4 dicek atau tidak
    Call cekSelect(CLng(bteColUpdate)) 'utk col 5 dicek atau tidak
    Call cekSelect(CLng(bteColPrice)) 'utk col 6 dicek atau tidak
End With
    
    If rsPriv.State <> adStateClosed Then rsPriv.Close
End Sub

Sub cekSelect(kol As Long)
Dim cek, noCek As Integer

With Grid
    '******** agar cekBox nya jika semuanya udah ke-select/not
    cek = 0
    For i = 1 To .Rows - 1
        If .Cell(flexcpChecked, i, kol) = flexChecked Then
            cek = cek + 1
        Else
            noCek = noCek + 1
        End If
    Next i
    
    If cek = .Rows - 1 Then
        .Cell(flexcpChecked, 0, kol) = flexChecked
    ElseIf noCek >= 1 Then
        .Cell(flexcpChecked, 0, kol) = flexUnchecked
    End If
    
    '****************
End With
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim cek As Integer

With Grid
    If Row <> 0 Then
        If Col = bteColAccess Or Col = bteColUpdate Or Col = bteColPrice Then Call cekSelect(Col)
    Else
        If Col = bteColAccess Or Col = bteColUpdate Or Col = bteColPrice Then
            If .Cell(flexcpChecked, Row, Col) = 1 Then
                cek = 1 'flexChecked
            Else
                cek = 2 'flexUnchecked
            End If
                
            For i = 1 To .Rows - 1
                .Cell(flexcpChecked, i, Col) = cek
            Next i
        End If
    End If
End With
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With Grid
    If Col <> bteColAccess And Col <> bteColUpdate And Col <> bteColPrice Then Cancel = 1
End With
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
    sql = "select * from user_Setup order by userName"
    If rsUser.State <> adStateClosed Then rsUser.Close
    rsUser.Open sql, Db, adOpenDynamic, adLockOptimistic
    Call Kosong
    Call IsiGrid

    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
    
End Sub

Private Sub cboUser_Click()
    If cboUser <> "" Then
        If cboUser.MatchFound Then
            Call IsiGrid(1, cboUser)
        Else
            LblErrMsg = DisplayMsg(4001)
        End If
    End If
End Sub

Private Sub Command1_Click(Index As Integer)

Me.MousePointer = vbHourglass
Select Case Index
    Case 0: 'Simpan & Ubah
        If txtUser = "" Then
            txtUser.SetFocus
            LblErrMsg = DisplayMsg(1002)
        ElseIf txtName = "" Then
            txtName.SetFocus
            LblErrMsg = DisplayMsg(1003)
        ElseIf txtPass1 = "" Or txtPass2 = "" Then
            txtPass1.SetFocus
            LblErrMsg = DisplayMsg(1004)
        ElseIf txtPass1 <> txtPass2 Then
            txtPass2.SetFocus
            LblErrMsg = DisplayMsg(1005)
        Else
            'Penambahan Validasi Charater dan symbol 20250128
            If CheckUserPassword(txtPass1) = False Then
                'LblErrMsg = ""
                Me.MousePointer = vbDefault
                Exit Sub
            End If
            
            With rsUser
                .filter = ""
                .Requery
                .filter = "userName ='" & txtUser & "' and App_ID = 'P01'"
                    
                If Not (.EOF) And ubah = False Then LblErrMsg = DisplayMsg(1001): Me.MousePointer = vbDefault: Exit Sub
                
                If .EOF Then .AddNew
                !app_ID = "P01"
                !userName = txtUser
                !Name = txtName
                !Password = fc_Encrypt(txtPass1)
                !Description = txtDesc
                !InitPO = txtpo
                !status_Admin = IIf(optStatus(0).Value, 1, 0)
                !locked = Check1.Value
                !InvalidLogin = 0
                !Last_Update = Now
                !last_user = userLogin
                .update
            End With
            Call simpanGrid
            
            Call isiCbo
            Call isiList
            Call IsiGrid(1, txtUser)
            txtUser.Enabled = False
            ubah = True
            LblErrMsg = DisplayMsg(1000)
        End If
    
    Case 1: 'Clear
        Call Kosong
        Call IsiGrid
    
    Case 2: 'Delete
    Dim tanya
    Dim rsCek As New ADODB.Recordset
        tanya = MsgBox("Do You Really want to Delete User " & txtUser & " ?", vbQuestion & vbYesNo, "Confirmation")
        If tanya = vbYes Then
            sql = "select count(userName) as jml From  User_Setup " & _
                "where app_ID = 'P01' and status_Admin ='1'"
            Set rsCek = Db.Execute(sql)
            
            If rsCek("jml") = 1 And pilihAdmin = 1 Then
                LblErrMsg = DisplayMsg(1206)
            Else
                sql = "delete user_setup where userName = '" & txtUser & "' and App_ID ='P01'"
                Db.Execute sql
                Call Kosong
                Call IsiGrid
                LblErrMsg = DisplayMsg(1201)
            End If
        End If
End Select
Me.MousePointer = vbDefault

End Sub

Sub simpanGrid()
Dim rsSimpan As New ADODB.Recordset
Dim sqlB As String
    
    With Grid
        sqlB = "select * from user_Privilege " & _
            "where userName = '" & txtUser & "' and App_ID = 'P01'"
        
        For i = 1 To Grid.Rows - 1
            sql = sqlB & " and Menu_Id ='" & Trim(.TextMatrix(i, bteColMenuID)) & "'"
            If rsSimpan.State <> adStateClosed Then rsSimpan.Close
            rsSimpan.Open sql, Db, adOpenStatic, adLockOptimistic
            
            If rsSimpan.EOF Then
                sql = "insert into user_Privilege (App_ID,userName,Menu_Id,status,Allow_Update,Allow_Price,Last_Update,Last_User) " & _
                    "values ('P01','" & txtUser & "','" & .TextMatrix(i, bteColMenuID) & "','" & _
                    IIf(.Cell(flexcpChecked, i, bteColAccess) = flexUnchecked, "0", "1") & "','" & _
                    IIf(.Cell(flexcpChecked, i, bteColUpdate) = flexUnchecked, "0", "1") & "','" & _
                    IIf(.Cell(flexcpChecked, i, bteColPrice) = flexUnchecked, "0", "1") & "', getdate(),'" & userLogin & "')"
                Db.Execute sql
            Else
                sql = "update user_Privilege " & _
                    "set status ='" & IIf(.Cell(flexcpChecked, i, bteColAccess) = flexUnchecked, "0", "1") & "', " & _
                    "Allow_Update ='" & IIf(.Cell(flexcpChecked, i, bteColUpdate) = flexUnchecked, "0", "1") & "', " & _
                    "Allow_Price ='" & IIf(.Cell(flexcpChecked, i, bteColPrice) = flexUnchecked, "0", "1") & "', " & _
                    "Last_Update =getdate(), " & _
                    "Last_User ='" & userLogin & "' " & _
                    "where App_ID = 'P01' and " & _
                    "UserName = '" & txtUser & "' and " & _
                    "Menu_Id ='" & .TextMatrix(i, bteColMenuID) & "'"
                Db.Execute sql
            End If
        Next i
    End With
    If rsSimpan.State <> adStateClosed Then rsSimpan.Close
End Sub

Private Sub lvw1_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1
End Sub

Private Sub lvw1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    ubah = True
    LblErrMsg = ""
    Call tampil   'utk menampilkan data yg diklik
End Sub

Private Sub Check1_Click()
    Check1.Caption = IIf(Check1.Value = 1, "Locked", "Unlocked")
End Sub

Sub tampil()
With rsUser
    .filter = ""
    .Requery
    .filter = "userName ='" & lvw1.selectedItem & "' and App_ID = 'P01'"
    If Not (.EOF) Then
        txtUser = Trim(!userName)
        txtUser.Enabled = False
        txtName = Trim(!Name)
        txtPass1 = fc_Decrypt(Trim(!Password))
        txtPass2 = fc_Decrypt(Trim(!Password))
        If !status_Admin = 1 Then
            optStatus(0).Value = True
            pilihAdmin = 1
        Else
            optStatus(1).Value = True
            pilihAdmin = 0
        End If
        
        Check1.Value = !locked
        Call Check1_Click
        txtDesc = Trim(!Description)
        txtpo = Trim(!InitPO & "")
        cboUser = ""
    End If
End With
Call IsiGrid(1, txtUser)
End Sub


Private Sub txtPO_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub txtUser_Change()
    lblNama = txtUser
End Sub

Private Sub txtUser_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub txtPass1_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub txtPass2_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub cboUser_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = Asc("'") Then KeyAscii = 0
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub CmdSubMenu_Click()
    DoEvents
    frmMainMenu.Show
    DoEvents
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If rsUser.State <> adStateClosed Then rsUser.Close
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    LblErrMsg.Caption = ErrMsg
End If
End Sub

Public Function CheckUserPassword(ByVal Password As String) As Boolean
sql = "EXEC dbo.sp_Check_SP '" & txtPass1.Text & "' "

If RS.State <> adStateClosed Then RS.Close
RS.Open sql, Db, adOpenForwardOnly, adLockReadOnly
    
If Not RS.EOF Then
    If RS.Fields("Message").Value = "OK !" Then
        CheckUserPassword = True
    Else
        CheckUserPassword = False
        LblErrMsg = RS.Fields("Message").Value
    End If
End If

End Function
