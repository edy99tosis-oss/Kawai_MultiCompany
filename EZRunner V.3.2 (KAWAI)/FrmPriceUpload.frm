VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmPriceUpload 
   BackColor       =   &H00FDDFE3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Upload"
   ClientHeight    =   9915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14385
   Icon            =   "FrmPriceUpload.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9915
   ScaleWidth      =   14385
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd 
      Left            =   3240
      Top             =   9240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdBrowse 
      BackColor       =   &H0080FFFF&
      Caption         =   "..."
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
      Left            =   12105
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   420
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   2475
      Left            =   157
      TabIndex        =   11
      Top             =   6000
      Width           =   13950
      Begin VB.Label LblCheck12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(0)"
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
         Left            =   13320
         TabIndex        =   35
         Top             =   1875
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ": Data Already Exist"
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
         Left            =   11160
         TabIndex        =   34
         Top             =   1875
         Width           =   1725
      End
      Begin VB.Label LblCheck11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(0)"
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
         Left            =   13320
         TabIndex        =   33
         Top             =   1395
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ": OK "
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
         Left            =   11160
         TabIndex        =   32
         Top             =   1395
         Width           =   450
      End
      Begin VB.Label LblCheck10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(0)"
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
         Left            =   13320
         TabIndex        =   31
         Top             =   915
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ": Invalid Reason"
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
         Left            =   11160
         TabIndex        =   30
         Top             =   915
         Width           =   1410
      End
      Begin VB.Label LblCheck9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(0)"
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
         Left            =   13320
         TabIndex        =   29
         Top             =   435
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ": Invalid End Date"
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
         Left            =   11160
         TabIndex        =   28
         Top             =   435
         Width           =   1560
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFF00&
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   11
         Left            =   10320
         Top             =   1800
         Width           =   735
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00000000&
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   10
         Left            =   10320
         Top             =   1320
         Width           =   735
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00000000&
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   9
         Left            =   10320
         Top             =   840
         Width           =   735
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00000000&
         FillColor       =   &H00004080&
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   8
         Left            =   10320
         Top             =   360
         Width           =   735
      End
      Begin VB.Label LblCheck8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(0)"
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
         Left            =   8520
         TabIndex        =   27
         Top             =   1875
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ": Invalid Start Date"
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
         Left            =   6360
         TabIndex        =   26
         Top             =   1875
         Width           =   1665
      End
      Begin VB.Label LblCheck7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(0)"
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
         Left            =   8520
         TabIndex        =   25
         Top             =   1395
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ": Invalid Unit"
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
         Left            =   6360
         TabIndex        =   24
         Top             =   1395
         Width           =   1110
      End
      Begin VB.Label LblCheck6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(0)"
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
         Left            =   8520
         TabIndex        =   23
         Top             =   915
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ": Invalid Price"
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
         Left            =   6360
         TabIndex        =   22
         Top             =   915
         Width           =   1200
      End
      Begin VB.Label LblCheck5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(0)"
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
         Left            =   8520
         TabIndex        =   21
         Top             =   435
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ": Invalid Price Cls"
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
         Left            =   6360
         TabIndex        =   20
         Top             =   435
         Width           =   1530
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFC0FF&
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   7
         Left            =   5520
         Top             =   1800
         Width           =   735
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00000000&
         FillColor       =   &H00C0C000&
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   6
         Left            =   5520
         Top             =   1320
         Width           =   735
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00000000&
         FillColor       =   &H00FF00FF&
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   5
         Left            =   5520
         Top             =   840
         Width           =   735
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00000000&
         FillColor       =   &H00C000C0&
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   4
         Left            =   5520
         Top             =   360
         Width           =   735
      End
      Begin VB.Label LblCheck4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(0)"
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
         TabIndex        =   19
         Top             =   1875
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ": Invalid Currency"
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
         Left            =   1200
         TabIndex        =   18
         Top             =   1875
         Width           =   1575
      End
      Begin VB.Label LblCheck3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(0)"
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
         TabIndex        =   17
         Top             =   1395
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ": Invalid Priority "
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
         Left            =   1200
         TabIndex        =   16
         Top             =   1395
         Width           =   1455
      End
      Begin VB.Label LblCheck2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(0)"
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
         TabIndex        =   15
         Top             =   915
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ": Invalid Trade Code "
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
         Left            =   1200
         TabIndex        =   14
         Top             =   915
         Width           =   1845
      End
      Begin VB.Label LblCheck1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(0)"
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
         TabIndex        =   13
         Top             =   435
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ": Invalid Item Code "
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
         Left            =   1200
         TabIndex        =   12
         Top             =   435
         Width           =   1755
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00000000&
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   3
         Left            =   360
         Top             =   1800
         Width           =   735
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00000000&
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   2
         Left            =   360
         Top             =   1320
         Width           =   735
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00000000&
         FillColor       =   &H000040C0&
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   1
         Left            =   360
         Top             =   840
         Width           =   735
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00000000&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   0
         Left            =   360
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.CommandButton CmdExcel 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Excel"
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
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9315
      Width           =   1155
   End
   Begin VB.CommandButton CmdImport 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Import"
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
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   1515
   End
   Begin VB.CommandButton CmdBack 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Back"
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
      Left            =   157
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9315
      Width           =   1155
   End
   Begin VB.TextBox txtBrowse 
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
      Height          =   315
      Left            =   720
      MaxLength       =   200
      TabIndex        =   0
      Top             =   360
      Width           =   10620
   End
   Begin VB.CommandButton CmdUpload 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Upload"
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
      Left            =   12915
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9315
      Width           =   1140
   End
   Begin VB.CommandButton CmdCheck 
      BackColor       =   &H0080FFFF&
      Caption         =   "C&heck"
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
      Left            =   11625
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9315
      Width           =   1155
   End
   Begin VSFlex8Ctl.VSFlexGrid gridSearch 
      Height          =   4995
      Left            =   150
      TabIndex        =   9
      Top             =   900
      Width           =   13950
      _cx             =   2088787998
      _cy             =   2088772203
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
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   275
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmPriceUpload.frx":0E42
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
      Editable        =   0
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   157
      TabIndex        =   7
      Top             =   8535
      Width           =   13950
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
         Height          =   240
         Left            =   105
         TabIndex        =   8
         Top             =   195
         Width           =   13710
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File"
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
      Left            =   150
      TabIndex        =   10
      Top             =   360
      Width           =   405
   End
End
Attribute VB_Name = "FrmPriceUpload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public getItemCode As String
Public getPartNumber As String
Private mrsData As Recordset

Dim bteColPriceCls As Byte
Dim bteColItemCode As Byte
Dim bteColTradeCode As Byte
Dim bteColPriority As Byte
Dim bteColCurr As Byte
Dim bteColPrice As Byte
Dim bteColUnit As Byte
Dim bteColStartDate As Byte
Dim bteColEndDate As Byte
Dim bteColReason As Byte
Dim bteColRemarks As Byte
Dim bteColErrMessege As Byte
Dim strFileName As String

Dim Check1 As Integer, Check2 As Integer, Check3 As Integer, Check4 As Integer, Check5 As Integer, Check6 As Integer
Dim Check7 As Integer, Check8 As Integer, Check9 As Integer, Check10 As Integer, Check11 As Integer, Check12 As Integer

Public Sub isiCbx()
   
End Sub

Public Sub IsiGrid()
    Dim RS As New ADODB.Recordset
    Dim query As String
        
    query = " SELECT    Item_Code, MakerItem_Code, Item_Name " & vbCrLf & _
            " FROM  Item_Master " & vbCrLf & _
            " WHERE WH_Code IN (SELECT WH_Code FROM User_Warehouse WHERE UserName = '" & Trim(userLogin) & "' ) " & vbCrLf & _
            "       OR Manufacture_Code IN (SELECT Trade_Code FROM User_Factory WHERE UserName = '" & Trim(userLogin) & "') " & vbCrLf & _
            "       OR Section_Cls IN (SELECT Section_Cls FROM User_Section WHERE UserName = '" & Trim(userLogin) & "') "
            
    If RS.State = adStateOpen Then RS.Close
    RS.Open query, Db, adOpenForwardOnly, adLockReadOnly

    gridSearch.clear
    Header
    
    gridSearch.ColS = 3
    gridSearch.Editable = flexEDNone

    i = 0
    While Not RS.EOF
        With gridSearch
        i = i + 1
        .AddItem ""
        .TextMatrix(i, 0) = Trim(RS!Item_Code & "")
        .TextMatrix(i, 1) = Trim(RS!MakerItem_Code & "")
        .TextMatrix(i, 2) = Trim(RS!item_name & "")
        RS.MoveNext
        End With
    Wend
End Sub

Public Sub search()
   
End Sub

Public Sub Header()
   
    bteColPriceCls = 0
    bteColItemCode = 1
    bteColTradeCode = 2
    bteColPriority = 3
    bteColCurr = 4
    bteColPrice = 5
    bteColUnit = 6
    bteColStartDate = 7
    bteColEndDate = 8
    bteColReason = 9
    bteColRemarks = 10
    bteColErrMessege = 11
    
    With gridSearch
        .clear
        .Rows = 1
        .ColS = 12
        
        .TextMatrix(0, bteColPriceCls) = "Cls"
        .TextMatrix(0, bteColItemCode) = "Item Code"
        .TextMatrix(0, bteColTradeCode) = "Trade Code"
        .TextMatrix(0, bteColPriority) = "Priority"
        .TextMatrix(0, bteColCurr) = "Curr"
        .TextMatrix(0, bteColPrice) = "Price"
        .TextMatrix(0, bteColUnit) = "Unit"
        .TextMatrix(0, bteColStartDate) = "Start Date"
        .TextMatrix(0, bteColEndDate) = "End Date"
        .TextMatrix(0, bteColReason) = "Reason"
        .TextMatrix(0, bteColRemarks) = "Remarks"
        .TextMatrix(0, bteColErrMessege) = "Error Message"
        
        .ColWidth(bteColPriceCls) = 800
        .ColWidth(bteColItemCode) = 3000
         .ColWidth(bteColTradeCode) = 3000
        .ColWidth(bteColPriority) = 800
        .ColWidth(bteColCurr) = 900
        .ColWidth(bteColPrice) = 1700
        .ColWidth(bteColUnit) = 1000
        .ColWidth(bteColStartDate) = 1400
        .ColWidth(bteColEndDate) = 1400
        .ColWidth(bteColReason) = 2200
        .ColWidth(bteColRemarks) = 2200
        .ColWidth(bteColErrMessege) = 2200
                
        .ColDataType(bteColPrice) = flexDTCurrency
        .ColDataType(bteColStartDate) = flexDTDate
        .ColDataType(bteColEndDate) = flexDTDate
        
     
        
        .Cell(flexcpAlignment, 0, 0, 0, bteColErrMessege) = flexAlignCenterCenter
        .ColAlignment(bteColPriceCls) = flexAlignLeftCenter
        .ColAlignment(bteColItemCode) = flexAlignLeftCenter
        .ColAlignment(bteColTradeCode) = flexAlignLeftCenter
        .ColAlignment(bteColPriority) = flexAlignCenterCenter
        .ColAlignment(bteColCurr) = flexAlignCenterCenter
        .ColAlignment(bteColPrice) = flexAlignRightCenter
        .ColAlignment(bteColUnit) = flexAlignCenterCenter
        .ColAlignment(bteColStartDate) = flexAlignCenterCenter
        .ColAlignment(bteColEndDate) = flexAlignCenterCenter
        .ColAlignment(bteColReason) = flexAlignLeftCenter
        .ColAlignment(bteColRemarks) = flexAlignLeftCenter
        .ColAlignment(bteColErrMessege) = flexAlignLeftCenter
        
        .EditMaxLength = 1
    End With
   
End Sub



Private Sub cmdBack_Click()
    Unload Me
End Sub

Private Sub cmdBack2_Click()

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Public Sub awal()
    txtBrowse.Text = ""
    LblErrMsg = ""
    
    checkStatus ("Load")
End Sub

Private Sub cmdBrowse_Click()

On Error GoTo errHandler
    cd.CancelError = True
    cd.InitDir = serverPath
    cd.filter = "Excel Files |*.xls*"
    cd.ShowOpen
    txtBrowse.Text = cd.filename
    
ErrExit:
    Me.MousePointer = vbDefault
    Exit Sub
errHandler:
    LblErrMsg.Caption = "[" & err.number & "] " & err.Description
    err.clear
    Resume ErrExit
End Sub

Private Sub CmdCheck_Click()
    Dim i As Integer, j As Integer, X As Integer
    Dim rsCek As New ADODB.Recordset
    Dim Status As Boolean
    Check1 = 0: Check2 = 0: Check3 = 0: Check4 = 0: Check5 = 0: Check6 = 0
    Check7 = 0: Check8 = 0: Check9 = 0: Check10 = 0: Check11 = 0: Check12 = 0
    i = 1
    Status = True
    LblErrMsg = ""
    With gridSearch
        'looping baris
        For i = 1 To .Rows - 1
            j = 0
            'looping kolom
            For j = 0 To bteColRemarks
                'cek item code
                '========================================================================================
                If j = bteColItemCode Then
                    sql = "EXEC sp_UploadPrice_Validate '1', '" & RTrim(.TextMatrix(i, bteColItemCode)) & "'"
                    If rsCek.State <> adStateClosed Then rsCek.Close
                    rsCek.CursorLocation = adUseClient
                    rsCek.Open sql, Db, adOpenForwardOnly, adLockReadOnly
                    'Cek item sudah ada di item master atau belum
                    If rsCek.RecordCount > 0 Then
                        Status = True
                    'Invalid Item Code
                    Else
                        For X = 0 To bteColErrMessege
                            .Cell(flexcpBackColor, i, X) = &HFF&
                        Next X
                        .TextMatrix(i, bteColErrMessege) = "Invalid Item Code"
                        Check1 = Check1 + 1
                        Status = False
                        Exit For
                    End If
                End If
                '========================================================================================
                
                'cek trade code
                '========================================================================================
                If j = bteColTradeCode Then
                    sql = "EXEC sp_UploadPrice_Validate '2','','" & RTrim(.TextMatrix(i, bteColTradeCode)) & "'"
                    If rsCek.State <> adStateClosed Then rsCek.Close
                    rsCek.CursorLocation = adUseClient
                    rsCek.Open sql, Db, adOpenForwardOnly, adLockReadOnly
                    'Cek trade sudah ada di trade master atau belum
                    If RTrim(.TextMatrix(i, bteColTradeCode)) <> "000000" Then
                        If rsCek.RecordCount > 0 Then
                            Status = True
                        'Invalid Trade Code
                        Else
                            For X = 0 To bteColErrMessege
                                .Cell(flexcpBackColor, i, X) = &H40C0&
                            Next X
                            .TextMatrix(i, bteColErrMessege) = "Invalid Trade Code"
                            Check2 = Check2 + 1
                            Status = False
                            Exit For
                        End If
                    Else
                        Status = True
                    End If
                End If
                
                '========================================================================================
                
                'cek Priority
                '========================================================================================
                If j = bteColPriority Then
                    'Cek priority 0 atau 1
                    If .TextMatrix(i, bteColPriority) = "0" Or .TextMatrix(i, bteColPriority) = "1" Then
                        Status = True
                    'Invalid Priority
                    Else
                        For X = 0 To bteColErrMessege
                            .Cell(flexcpBackColor, i, X) = &H80FF&
                        Next X
                        .TextMatrix(i, bteColErrMessege) = "Invalid Priority"
                        Check3 = Check3 + 1
                        Status = False
                        Exit For
                    End If
                End If
                '========================================================================================
                
                'cek currency
                '========================================================================================
                If j = bteColCurr Then
                    sql = "EXEC sp_UploadPrice_Validate '3','','','" & RTrim(.TextMatrix(i, bteColCurr)) & "'"
                    If rsCek.State <> adStateClosed Then rsCek.Close
                    rsCek.CursorLocation = adUseClient
                    rsCek.Open sql, Db, adOpenForwardOnly, adLockReadOnly
                    'Cek Currency sudah ada di Curr Cls atau belum
                    If rsCek.RecordCount > 0 Then
                        Status = True
                    'Invalid Currency
                    Else
                        For X = 0 To bteColErrMessege
                            .Cell(flexcpBackColor, i, X) = &HFFFF&
                        Next X
                        .TextMatrix(i, bteColErrMessege) = "Invalid Currency"
                        Check4 = Check4 + 1
                        Status = False
                        Exit For
                    End If
                End If
                '========================================================================================
                
                'cek Price Cls
                '========================================================================================
                If j = bteColPriceCls Then
                    'Cek Price Cls 01, 02, 03, 04, 05
                    If .TextMatrix(i, bteColPriceCls) = "01" Or .TextMatrix(i, bteColPriceCls) = "02" Or .TextMatrix(i, bteColPriceCls) = "03" Or .TextMatrix(i, bteColPriceCls) = "04" Or .TextMatrix(i, bteColPriceCls) = "05" Or .TextMatrix(i, bteColPriceCls) = "06" Or .TextMatrix(i, bteColPriceCls) = "09" Or .TextMatrix(i, bteColPriceCls) = "10" Then
                        Status = True
                    'Invalid Price Cls
                    Else
                        For X = 0 To bteColErrMessege
                            .Cell(flexcpBackColor, i, X) = &HC000C0
                        Next X
                        .TextMatrix(i, bteColErrMessege) = "Invalid Price Cls"
                        Check5 = Check5 + 1
                        Status = False
                        Exit For
                    End If
                End If
                '========================================================================================
                'cek Price
                '========================================================================================
                If j = bteColPrice Then
                     If IsNumeric(.TextMatrix(i, bteColPrice)) Then
                        Status = True
                    'Invalid Price
                    Else
                        For X = 0 To bteColErrMessege
                            .Cell(flexcpBackColor, i, X) = &HFF00FF
                        Next X
                        .TextMatrix(i, bteColErrMessege) = "Invalid Price "
                        Check6 = Check6 + 1
                        Status = False
                        Exit For
                    End If
                End If
                '========================================================================================
                If j = bteColPrice Then
                    'Cek Price numeric
                    If Len(.TextMatrix(i, bteColPrice)) <= 23 Then
                        Status = True
                     Else
                        For X = 0 To bteColErrMessege
                            .Cell(flexcpBackColor, i, X) = &HFF00FF
                        Next X
                        .TextMatrix(i, bteColErrMessege) = "Invalid Price"
                        Check6 = Check6 + 1
                        Status = False
                        Exit For
                    End If
                End If
                
                'cek unit
                '========================================================================================
                If j = bteColUnit Then
                    sql = "EXEC sp_UploadPrice_Validate '4','','','','" & RTrim(.TextMatrix(i, bteColUnit)) & "'"
                    If rsCek.State <> adStateClosed Then rsCek.Close
                    rsCek.CursorLocation = adUseClient
                    rsCek.Open sql, Db, adOpenForwardOnly, adLockReadOnly
                    'Cek unit sudah ada di unit Cls atau belum
                    If rsCek.RecordCount > 0 Then
                        Status = True
                    Else
                        For X = 0 To bteColErrMessege
                            .Cell(flexcpBackColor, i, X) = &HC0C000
                        Next X
                        .TextMatrix(i, bteColErrMessege) = "Invalid Unit"
                        Check7 = Check7 + 1
                        Status = False
                        Exit For
                    End If
                End If
                '========================================================================================
                
                'cek Start Date
                '========================================================================================
                If j = bteColStartDate Then
                    'Cek Start Date
                    If IsDate(.TextMatrix(i, bteColStartDate)) Then
                        Status = True
                    Else
                        For X = 0 To bteColErrMessege
                            .Cell(flexcpBackColor, i, X) = &HFFC0FF
                        Next X
                        .TextMatrix(i, bteColErrMessege) = "Invalid Start Date "
                        Check8 = Check8 + 1
                        Status = False
                        Exit For
                    End If
                End If
                '========================================================================================
                
                'cek End Date
                '========================================================================================
                If j = bteColEndDate Then
                    'Cek End Date
                    If IsDate(.TextMatrix(i, bteColEndDate)) Or .TextMatrix(i, bteColEndDate) = "9999-99-99" Or .TextMatrix(i, bteColEndDate) = "99/99/9999" Then
                        Status = True
                    Else
                        For X = 0 To bteColErrMessege
                            .Cell(flexcpBackColor, i, X) = &H4080&
                        Next X
                        .TextMatrix(i, bteColErrMessege) = "Invalid End Date "
                        Check9 = Check9 + 1
                        Status = False
                        Exit For
                    End If
                End If
                '========================================================================================
                
                'cek Reason - Updated by Berth
                '========================================================================================
                If j = bteColReason Then
                   sql = "EXEC sp_UploadPrice_Validate '5','','','','','" & RTrim(.TextMatrix(i, bteColReason)) & "'"
                    If rsCek.State <> adStateClosed Then rsCek.Close
                    rsCek.CursorLocation = adUseClient
                    rsCek.Open sql, Db, adOpenForwardOnly, adLockReadOnly
                    'Cek unit sudah ada di unit Cls atau belum
                    If rsCek.RecordCount > 0 Then
                        Status = True
                    'Invalid Unit
                    Else
                        For X = 0 To bteColErrMessege
                            .Cell(flexcpBackColor, i, X) = &HC0C0C0
                        Next X
                        .TextMatrix(i, bteColErrMessege) = "Invalid Reason"
                        Check10 = Check10 + 1
                        Status = False
                        Exit For
                    End If
                End If
                
            Next j
            
            'jika semua kolom sudah dicek dan ok
            If Status = True Then
                sql = "EXEC sp_UploadPrice_Validate '6','" & RTrim(.TextMatrix(i, bteColItemCode)) & "', " & vbCrLf & _
                      " '" & RTrim(.TextMatrix(i, bteColTradeCode)) & "','','','','" & RTrim(.TextMatrix(i, bteColPriority)) & "', " & vbCrLf & _
                      " '" & RTrim(.TextMatrix(i, bteColPriceCls)) & "','" & Format(RTrim(.TextMatrix(i, bteColStartDate)), "YYYYMMdd") & "'"

                If rsCek.State <> adStateClosed Then rsCek.Close
                rsCek.CursorLocation = adUseClient
                rsCek.Open sql, Db, adOpenForwardOnly, adLockReadOnly
                
                'Data Already Exist
                If rsCek.RecordCount > 0 Then
                    For X = 0 To bteColErrMessege
                        .Cell(flexcpBackColor, i, X) = &HFFFF00
                    Next X
                    .TextMatrix(i, bteColErrMessege) = "Data Already Exist"
                    Check12 = Check12 + 1
                'Data OK
                Else
                    For X = 0 To bteColErrMessege
                        .Cell(flexcpBackColor, i, X) = &HFF00&
                    Next X
                    .TextMatrix(i, bteColErrMessege) = "OK"
                    Check11 = Check11 + 1
                End If
            End If
            Status = True
        Next i
        LblCheck1 = "(" & Check1 & ")": LblCheck2 = "(" & Check2 & ")": LblCheck3 = "(" & Check3 & ")"
        LblCheck4 = "(" & Check4 & ")": LblCheck5 = "(" & Check5 & ")": LblCheck6 = "(" & Check6 & ")"
        LblCheck7 = "(" & Check7 & ")": LblCheck8 = "(" & Check8 & ")": LblCheck9 = "(" & Check9 & ")"
        LblCheck10 = "(" & Check10 & ")": LblCheck11 = "(" & Check11 & ")": LblCheck12 = "(" & Check12 & ")"
        
        checkStatus ("Check")
        
    End With
    
ErrExit:
    Me.MousePointer = vbDefault
    Exit Sub
errHandler:
    LblErrMsg.Caption = "[" & err.number & "] " & err.Description
    err.clear
    Resume ErrExit
End Sub

Private Sub CmdExcel_Click()

    Me.MousePointer = vbHourglass
 
    LblErrMsg.Caption = ""
    If gridSearch.Rows > 1 Then
        Call up_ExcelOpen
    Else
        LblErrMsg.Caption = DisplayMsg("0013")
        Me.MousePointer = vbDefault
    End If
    

End Sub
Private Sub up_ExcelOpen()
    Dim ExlFile As New Excel.application
    Dim i As Integer
    On Error GoTo errHandler
    
    With ExlFile
        
        i = 0
        
        .Workbooks.Add
        .Range("A1:K1").Font.Bold = True
        .Range("A1:K1").Borders.color = xlGray16
        .Range("A1:K" & gridSearch.Rows).Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range("A" & gridSearch.Rows & ":K" & gridSearch.Rows).Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("A1:A" & gridSearch.Rows).Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("B1:B" & gridSearch.Rows).Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("C1:C" & gridSearch.Rows).Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("D1:D" & gridSearch.Rows).Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("E1:E" & gridSearch.Rows).Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("F1:F" & gridSearch.Rows).Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("G1:G" & gridSearch.Rows).Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("H1:H" & gridSearch.Rows).Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("I1:I" & gridSearch.Rows).Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("J1:J" & gridSearch.Rows).Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("K1:K" & gridSearch.Rows).Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("K1:K" & gridSearch.Rows).Borders(xlEdgeRight).LineStyle = xlContinuous
            
        .Range("A1") = "Price Cls"
        .Range("B1") = "Item Code"
        .Range("C1") = "Trade Code"
        .Range("D1") = "Priority Cls"
        .Range("E1") = "Currency"
        .Range("F1") = "Price"
        .Range("G1") = "Start Date"
        .Range("H1") = "End Date"
        .Range("I1") = "Reason"
        .Range("J1") = "Remarks"
        .Range("K1") = "Error Message"
        For i = 1 To gridSearch.Rows - 1
            .Range("A" & i + 1) = "'" & gridSearch.TextMatrix(i, bteColPriceCls)
            .Range("B" & i + 1) = "'" & gridSearch.TextMatrix(i, bteColItemCode)
            .Range("C" & i + 1) = gridSearch.TextMatrix(i, bteColTradeCode)
            .Range("D" & i + 1) = gridSearch.TextMatrix(i, bteColPriority)
            .Range("E" & i + 1) = "'" & gridSearch.TextMatrix(i, bteColCurr)
            .Range("F" & i + 1) = gridSearch.TextMatrix(i, bteColPrice)
            .Range("F" & i + 1).NumberFormat = "#,##0.00"
            .Range("G" & i + 1) = gridSearch.TextMatrix(i, bteColStartDate)
'            .Range("G" & i + 1).NumberFormat = "YYYY-MM-DD"
            .Range("H" & i + 1) = gridSearch.TextMatrix(i, bteColEndDate)
'            .Range("H" & i + 1).NumberFormat = "YYYY-MM-DD"
            .Range("I" & i + 1) = gridSearch.TextMatrix(i, bteColReason)
            .Range("J" & i + 1) = gridSearch.TextMatrix(i, bteColRemarks)
            .Range("K" & i + 1) = gridSearch.TextMatrix(i, bteColErrMessege)
            .Range("A" & i + 1 & ":K" & i + 1).Borders(xlEdgeBottom).LineStyle = xlContinuous
        Next i
     
    End With
    ExlFile.Visible = True
    ExlFile.ActiveSheet.PageSetup.PaperSize = xlPaperA4
    ExlFile.ActiveSheet.PageSetup.Orientation = 2
    ExlFile.ActiveSheet.PageSetup.LeftMargin = application.InchesToPoints(0.4)
    ExlFile.ActiveSheet.PageSetup.RightMargin = application.InchesToPoints(0.4)
    ExlFile.Range("A:K").Columns.AutoFit
    ExlFile.WindowState = xlMaximized

    Me.MousePointer = vbHourglass
ErrExit:
    'Set rsCek = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
errHandler:
    LblErrMsg.Caption = "[" & err.number & "] " & err.Description
    err.clear
    Resume ErrExit
End Sub
Private Sub CmdImport_Click()
Dim i As Integer, Row As Integer, sql As String
Dim ExlFile As New Excel.application
Dim batas As Boolean, j As Integer
Dim rsCek As New Recordset
On Error GoTo errHandler
    Header
    batas = True
    LblCheck1 = "(0)": LblCheck2 = "(0)": LblCheck3 = "(0)"
    LblCheck4 = "(0)": LblCheck5 = "(0)": LblCheck6 = "(0)"
    LblCheck7 = "(0)": LblCheck8 = "(0)": LblCheck9 = "(0)"
    LblCheck10 = "(0)": LblCheck11 = "(0)": LblCheck12 = "(0)"
    If txtBrowse <> "" Then
        Me.MousePointer = vbHourglass
        i = 0
        ExlFile.Workbooks.Open txtBrowse.Text
        '############ GET ROWS ############
'        For j = 1 To 10
'            sql = " select * from item_master where Item_Code='" & RTrim(ExlFile.Range("B" & j)) & "' "
'
'            If rsCek.State <> adStateClosed Then rsCek.Close
'            rsCek.CursorLocation = adUseClient
'            rsCek.Open sql, Db, adOpenForwardOnly, adLockReadOnly
'
'            If Not (rsCek.BOF And rsCek.EOF) Then
'                Row = j
'                Exit For
'            End If
'        Next j
        '##################################
        Row = 2
        If ExlFile.Range("A" & Row) = "" Then
            batas = False
            LblErrMsg.Caption = DisplayMsg("0013")
            'Exit Sub
        Else
            Do Until batas = False
                With gridSearch
                    .AddItem ""
                    i = i + 1
                    .TextMatrix(i, bteColPriceCls) = ExlFile.Range("A" & Row)
                    .TextMatrix(i, bteColItemCode) = ExlFile.Range("B" & Row)
                    .TextMatrix(i, bteColTradeCode) = ExlFile.Range("C" & Row)
                    .TextMatrix(i, bteColPriority) = ExlFile.Range("D" & Row)
                    .TextMatrix(i, bteColCurr) = ExlFile.Range("E" & Row)

                    If .TextMatrix(i, bteColCurr) = "03" Then
                        .TextMatrix(i, bteColPrice) = Format(ExlFile.Range("F" & Row), gs_formatPriceIDR)
                    Else
                        .TextMatrix(i, bteColPrice) = Format(ExlFile.Range("F" & Row), gs_formatPrice)
                    End If
                    
                    sql = "Select rtrim(Unit_Cls)Unit_Cls From Item_Master Where Item_Code='" & ExlFile.Range("B" & Row) & "'"
                    If rsCek.State <> adStateClosed Then rsCek.Close
                    rsCek.CursorLocation = adUseClient
                    rsCek.Open sql, Db, adOpenForwardOnly, adLockReadOnly
                    If rsCek.RecordCount > 0 Then
                        .TextMatrix(i, bteColUnit) = IIf(IsNull(rsCek!Unit_cls), "", rsCek!Unit_cls)
                    Else
                        .TextMatrix(i, bteColUnit) = ""
                    End If
                    .TextMatrix(i, bteColStartDate) = ExlFile.Range("G" & Row)
                    .TextMatrix(i, bteColEndDate) = ExlFile.Range("H" & Row)
                    .TextMatrix(i, bteColReason) = ExlFile.Range("I" & Row)
                    .TextMatrix(i, bteColRemarks) = ExlFile.Range("J" & Row)
                    Row = Row + 1
                    If ExlFile.Range("A" & Row) = "" Then batas = False
                End With
            Loop
            LblErrMsg = ""
            
        End If
        ExlFile.Workbooks.Close
    End If
    
    checkStatus ("Import")
    
ErrExit:
    Set rsCek = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
errHandler:
    LblErrMsg.Caption = "[" & err.number & "] " & err.Description
    err.clear
    Resume ErrExit
End Sub
Private Function uf_readExcel(ByVal sFile As String) As Recordset
    On Error GoTo errHandler
    Dim RS As Recordset
    Dim sconn As String
    
    
    Set RS = New Recordset
    
    RS.CursorLocation = adUseClient
    RS.CursorType = adOpenForwardOnly
    RS.LockType = adLockReadOnly
    
    sconn = "DRIVER=Microsoft Excel Driver (*.xls);" & _
        "DBQ=" & sFile
        sql = " select * from [sheet1$]"
    RS.Open " select * from [sheet1$]", sconn
    Set uf_readExcel = RS
    
    Set RS = Nothing
    Exit Function
    
ErrExit:
    Set RS = Nothing
    Me.MousePointer = vbDefault
    Exit Function
errHandler:
    LblErrMsg.Caption = "[" & err.number & "] " & err.Description
    err.clear
    Resume ErrExit
End Function

Private Sub cmdUpload_Click()
    Dim i As Integer, j As Integer
    Dim rsCek As New ADODB.Recordset
    'On Error GoTo errHandler
    Me.MousePointer = vbHourglass
    i = 1: j = 0
    Call CmdCheck_Click
    If Check11 + Check12 <> gridSearch.Rows - 1 Then
        LblErrMsg = "[8159] Incorrect Data to Upload! "
        GoTo ErrExit
    End If
        
    With gridSearch
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                If .TextMatrix(i, bteColErrMessege) = "OK" Then
                        sql = " EXEC sp_UploadPrice_InsUpd '" & RTrim(.TextMatrix(i, bteColPriceCls)) & "', " & vbCrLf & _
                            " '" & RTrim(.TextMatrix(i, bteColItemCode)) & "'," & vbCrLf & _
                            " '" & RTrim(.TextMatrix(i, bteColTradeCode)) & "', " & vbCrLf & _
                            " '" & RTrim(.TextMatrix(i, bteColPriority)) & "', " & vbCrLf & _
                            " '" & RTrim(.TextMatrix(i, bteColCurr)) & "', " & vbCrLf & _
                            " " & CDbl(RTrim(.TextMatrix(i, bteColPrice))) & ", " & vbCrLf & _
                            " '" & RTrim(.TextMatrix(i, bteColUnit)) & "', " & vbCrLf & _
                            " '" & Format(RTrim(.TextMatrix(i, bteColStartDate)), "YYYYMMdd") & "', " & vbCrLf & _
                            " '" & IIf((.TextMatrix(i, bteColEndDate)) = "9999-99-99", "99999999", Format(RTrim(.TextMatrix(i, bteColEndDate)), "YYYYMMdd")) & "', " & vbCrLf & _
                            " '" & RTrim(IIf(.TextMatrix(i, bteColReason) = "", "00", .TextMatrix(i, bteColReason))) & "', " & vbCrLf & _
                            " '" & RTrim(.TextMatrix(i, bteColRemarks)) & "', " & vbCrLf & _
                            " '" & userLogin & "' "

                        Db.Execute (sql)
                        j = j + 1

                End If
            Next i
            If j = 0 Then
                LblErrMsg = "[8159] Incorrect Data to Upload! "
            Else
                LblErrMsg = "[8157] " & j & " Data Successfully Upload!"
            End If
        Else
            LblErrMsg = "[8158] There is No Data to Upload! "
        End If
    End With
ErrExit:
    Set rsCek = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
errHandler:
    LblErrMsg.Caption = "[" & err.number & "] " & err.Description
    err.clear
    Resume ErrExit
End Sub

Private Sub Form_Load()
    'IsiGrid
    Header
    awal
    Me.MousePointer = vbDefault
End Sub

Private Sub gridSearch_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub gridSearch_Click()
    'message
End Sub

Private Sub gridSearch_DblClick()
    'cmdBack_Click
End Sub

Private Sub gridSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = vbKeyReturn Then cmdBack_Click
End Sub

Private Sub gridSearch_RowColChange()
    'LblErrMsg.Caption = "Item Code = " & gridSearch.TextMatrix(gridSearch.RowSel, 0)
End Sub

Public Sub message()
    LblErrMsg.Caption = "Item Code = " & gridSearch.TextMatrix(gridSearch.RowSel, 0)
End Sub

Public Function validasi() As Boolean
    If txtBrowse.Text = "" Then
        validasi = False
    ElseIf txtBrowse.Text <> 0 Then
        validasi = True
    End If
End Function

Public Sub checkStatus(ByVal Status As String)
    If Status = "Load" Then
        CmdCheck.Enabled = False
        CmdExcel.Enabled = False
        CmdUpload.Enabled = False
    ElseIf Status = "Import" Then
        CmdCheck.Enabled = True
        CmdExcel.Enabled = True
    ElseIf Status = "Check" Then
        If Check11 = (gridSearch.Rows - 1) Then
            CmdUpload.Enabled = True
        Else
            CmdUpload.Enabled = False
        End If
    End If
End Sub

