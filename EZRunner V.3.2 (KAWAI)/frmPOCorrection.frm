VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPOCorrection 
   BackColor       =   &H00FDDFE3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PO Correction"
   ClientHeight    =   10830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   Icon            =   "frmPOCorrection.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10830
   ScaleWidth      =   15270
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   240
      TabIndex        =   18
      Top             =   9360
      Width           =   14835
      Begin VB.Label LblErr 
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
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   195
         Width           =   14370
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Caption         =   "General"
      Height          =   1215
      Left            =   240
      TabIndex        =   7
      Top             =   3000
      Width           =   14775
      Begin VB.TextBox TxtReason 
         Height          =   615
         Left            =   11400
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   360
         Width           =   3135
      End
      Begin VB.TextBox TxtTransport 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.#0"
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
         Left            =   8640
         MaxLength       =   15
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   720
         Width           =   1650
      End
      Begin VB.TextBox TxtPacking 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.#0"
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
         Left            =   8640
         MaxLength       =   15
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   300
         Width           =   1650
      End
      Begin VB.TextBox TxtPrice 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.#0"
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
         Left            =   3600
         MaxLength       =   15
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   720
         Width           =   1650
      End
      Begin VB.TextBox TxtPayment 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.#0"
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
         Left            =   3600
         MaxLength       =   15
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   300
         Width           =   1650
      End
      Begin MSForms.ComboBox cboTransport 
         Height          =   315
         Left            =   7080
         TabIndex        =   25
         Top             =   720
         Width           =   975
         VariousPropertyBits=   612386843
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "1720;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox CboPacking 
         Height          =   315
         Left            =   7080
         TabIndex        =   24
         Top             =   300
         Width           =   975
         VariousPropertyBits=   612386843
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "1720;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboPriceCondition 
         Height          =   315
         Left            =   2040
         TabIndex        =   23
         Top             =   720
         Width           =   975
         VariousPropertyBits=   612386843
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "1720;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboPayment 
         Height          =   315
         Left            =   2040
         TabIndex        =   22
         Top             =   300
         Width           =   975
         VariousPropertyBits=   612386843
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "1720;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label LblPart 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reason"
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
         Left            =   10680
         TabIndex        =   16
         Top             =   360
         Width           =   630
      End
      Begin VB.Label LblPart 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transportation"
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
         Left            =   5520
         TabIndex        =   13
         Top             =   780
         Width           =   1245
      End
      Begin VB.Label LblPart 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Packing"
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
         Left            =   5520
         TabIndex        =   12
         Top             =   360
         Width           =   660
      End
      Begin VB.Label LblPart 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Price Condition"
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
         TabIndex        =   9
         Top             =   780
         Width           =   1290
      End
      Begin VB.Label LblPart 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Term"
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
         TabIndex        =   8
         Top             =   360
         Width           =   1260
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      ForeColor       =   &H00000000&
      Height          =   2175
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   14775
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.#0"
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
         Height          =   345
         Left            =   3480
         MaxLength       =   15
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   742
         Width           =   2370
      End
      Begin VB.CommandButton cmdsearch 
         BackColor       =   &H0080FFFF&
         Caption         =   "Sea&rch"
         Height          =   345
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1680
         Width           =   1515
      End
      Begin VB.OptionButton OptUpdate 
         BackColor       =   &H00FDDFE3&
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1920
         TabIndex        =   27
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton OptCreate 
         BackColor       =   &H00FDDFE3&
         Caption         =   "Create"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   360
         TabIndex        =   26
         Top             =   240
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker SDate 
         Height          =   345
         Left            =   1920
         TabIndex        =   30
         Top             =   1200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
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
         CustomFormat    =   "MMM yyyy"
         Format          =   293535747
         UpDown          =   -1  'True
         CurrentDate     =   37868
      End
      Begin MSComCtl2.DTPicker EDate 
         Height          =   345
         Left            =   3720
         TabIndex        =   33
         Top             =   1200
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   609
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
         CustomFormat    =   "MMM yyyy"
         Format          =   293535747
         UpDown          =   -1  'True
         CurrentDate     =   37868
      End
      Begin MSForms.ComboBox cboSupp 
         Height          =   345
         Left            =   1920
         TabIndex        =   36
         Top             =   720
         Width           =   1575
         VariousPropertyBits=   746604571
         MaxLength       =   35
         DisplayStyle    =   3
         Size            =   "2778;609"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label LblPart 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Code "
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
         TabIndex        =   35
         Top             =   795
         Width           =   1275
      End
      Begin MSForms.ComboBox CboPOnO 
         Height          =   345
         Left            =   1920
         TabIndex        =   31
         Top             =   1680
         Width           =   1575
         VariousPropertyBits=   746604571
         MaxLength       =   35
         DisplayStyle    =   3
         Size            =   "2778;609"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox CboPOC 
         Height          =   345
         Left            =   3480
         TabIndex        =   28
         Top             =   270
         Width           =   2370
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "4180;609"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label LblPart 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PO No"
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
         TabIndex        =   21
         Top             =   1755
         Width           =   525
      End
      Begin VB.Label LblPart 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Periode"
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
         TabIndex        =   20
         Top             =   1275
         Width           =   645
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "to"
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
         Left            =   3530
         TabIndex        =   6
         Top             =   1320
         Width           =   165
      End
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
      Left            =   13920
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   10080
      Width           =   1125
   End
   Begin VB.CommandButton cmdsubmenu 
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   10050
      Width           =   1125
   End
   Begin VB.CommandButton CmdClear 
      BackColor       =   &H0080FFFF&
      Caption         =   "Clea&r"
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
      TabIndex        =   1
      Top             =   10080
      Width           =   1125
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   420
      Left            =   13080
      TabIndex        =   4
      Top             =   120
      Width           =   1860
      _extentx        =   3281
      _extenty        =   741
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   4575
      Left            =   240
      TabIndex        =   37
      Top             =   4320
      Width           =   14805
      _cx             =   26114
      _cy             =   8070
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
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
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
   Begin VB.Label LblTotal 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Record"
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
      Height          =   255
      Left            =   12840
      TabIndex        =   29
      Top             =   9000
      Width           =   2175
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PO Correction"
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
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   15105
   End
End
Attribute VB_Name = "frmPOCorrection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim li_HakU As Integer, ls_Answer As String
Dim sql As String
Dim RS As New ADODB.Recordset

Dim bteColItemCode As Byte
Dim bteColDesciption As Byte
Dim bteColOrderQty As Byte
Dim bteColCur As Byte
Dim bteColPrice As Byte
Dim bteColAmount As Byte

Private Sub cboPacking_Change()
Call cboPacking_Click
End Sub

Private Sub cboPacking_Click()
If CboPacking.MatchFound = True Then
   TxtPacking.Text = Trim(CboPacking.Column(1))
Else
   TxtPacking.Text = ""
End If
End Sub

Private Sub cboPayment_Change()
Call cbopayment_Click
End Sub

Private Sub cbopayment_Click()
If cboPayment.MatchFound = True Then
    TxtPayment.Text = Trim(cboPayment.Column(1))
Else
    TxtPayment.Text = ""
End If
End Sub

Private Sub CboPOC_Click()

If CboPOC.MatchFound = True Then
    CboPOnO.Text = Trim(CboPOC.Column(2))
Else
    CboPOnO.Text = ""
End If

If CboPOC.MatchFound = True Then
   cboSupp.Text = Trim(CboPOC.Column(1))
Else
    cboSupp.Text = ""
End If

If CboPOC.MatchFound = True Then
    cboPayment.Text = Trim(CboPOC.Column(4))
Else
    cboPayment.Text = ""
End If

If CboPOC.MatchFound = True Then
    cboPriceCondition.Text = Trim(CboPOC.Column(5))
Else
    cboPriceCondition.Text = ""
End If

If CboPOC.MatchFound = True Then
    CboPacking.Text = Trim(CboPOC.Column(6))
Else
    CboPacking.Text = ""
End If

If CboPOC.MatchFound = True Then
    cboTransport.Text = Trim(CboPOC.Column(7))
Else
    cboTransport.Text = ""
End If

If CboPOC.MatchFound = True Then
    TxtReason.Text = Trim(CboPOC.Column(8))
Else
    TxtReason.Text = ""
End If

'If CboPOC.MatchFound = True Then
'    sdate.Value = CboPOC.Column(3)
'Else
'    sdate.Value = Format(Now, "yyyymm")
'End If
End Sub

Private Sub cbopricecondition_Change()
Call cbopricecondition_Click
End Sub

Private Sub cbopricecondition_Click()
If cboPriceCondition.MatchFound = True Then
    TxtPrice.Text = Trim(cboPriceCondition.Column(1))
Else
    TxtPrice.Text = ""
End If
End Sub

Private Sub cboTransport_Change()
Call cbotransport_Click
End Sub
 
Private Sub cbotransport_Click()
If cboTransport.MatchFound = True Then
    TxtTransport.Text = Trim(cboTransport.Column(1))
Else
    TxtTransport.Text = ""
End If
End Sub

Private Sub cmdClear_Click()
Call blank
OptUpdate.Value = True
End Sub

Private Sub CmdSubmit_Click()
   Call uf_Validate
   If TxtReason.Text = "" Then
    TxtReason.SetFocus
    LblErr.Caption = DisplayMsg(1058)
    
    Exit Sub
     End If
   Call insertupdate
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub CboSupp_Change()
Call CboSupp_Click
End Sub

Private Sub CboSupp_Click()
Call adtocbopono
If cboSupp.MatchFound = True Then
    txtName.Text = Trim(cboSupp.Column(1))
Else
    txtName.Text = ""
End If
End Sub

Private Sub cmdSearch_Click()
Call uf_Validate
Call up_GridSearch
Call search
LblErr.Caption = ""
End Sub

Sub general()
Dim sql1 As String
Dim rs1 As New Recordset

    Call up_FillCombo(cboPayment, "PaymentTerm_Cls")
    cboPayment.ColumnWidths = "25pt;175pt"
    cboPayment.ListWidth = 200
    
    Call up_FillCombo(CboPacking, "POPacking_Cls")
    CboPacking.ColumnWidths = "25pt;175pt"
    CboPacking.ListWidth = 200
    
    Call up_FillCombo(cboTransport, "Transportation_Cls")
    cboTransport.ColumnWidths = "25pt;175pt"
    cboTransport.ListWidth = 200
    
    sql1 = "select * from PriceCondition_cls"
    If rs1.State <> adStateClosed Then rs1.Close
    rs1.Open sql1, Db, adOpenKeyset, adLockOptimistic
    If Not (rs1.BOF And rs1.EOF) Then
        i = 0
        With cboPriceCondition
            .clear
            .columnCount = 2
            .ColumnWidths = "25pt;175pt"
            .ListWidth = 200
            .ListRows = 15

            Do While Not rs1.EOF
                .AddItem ""
                .List(i, 0) = Trim(rs1!PriceCondition_Cls)
                .List(i, 1) = Trim(rs1!Description)
                i = i + 1
                rs1.MoveNext
            Loop
        End With
    End If
    Set rs1 = Nothing
End Sub

Sub uf_Validate()
If cboSupp.Text = "" Then
    cboSupp.SetFocus
    LblErr.Caption = DisplayMsg("1054")
ElseIf CboPOnO.Text = "" Then
    CboPOnO.SetFocus
    LblErr.Caption = DisplayMsg("1048")
End If
End Sub

Private Sub CmdSubMenu_Click()
    Unload Me
    frmMainMenu.Show
End Sub

Private Sub edate_Change()
CboSupp_Click
End Sub

Private Sub Form_Load()
  CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
Call Header
Call adtocombo
Call adtocbopono
Call blank
OptUpdate.Value = True
Call general
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If grid.TextMatrix(Row, bteColCur) = "IDR" Then
   grid.TextMatrix(Row, bteColAmount) = Format(grid.TextMatrix(Row, bteColOrderQty) * grid.TextMatrix(Row, bteColPrice), gs_formatAmountIDR)
   grid.TextMatrix(Row, bteColPrice) = Format(grid.TextMatrix(Row, bteColPrice), gs_formatPriceIDR)
Else
   grid.TextMatrix(Row, bteColAmount) = Format(grid.TextMatrix(Row, bteColOrderQty) * grid.TextMatrix(Row, bteColPrice), gs_formatAmount)
   grid.TextMatrix(Row, bteColPrice) = Format(grid.TextMatrix(Row, bteColPrice), gs_formatPrice)
End If
End Sub

Private Sub Grid_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not (Col = 4) Then
        Cancel = True
    End If
End Sub

Sub opt()
If OptCreate.Value = True Then
    Call PONO(Right(Year(EDate), 4), Format(Month(EDate), "0#"))
ElseIf OptUpdate.Value = True Then
End If
End Sub

Sub PONO(ByVal thn As String, ByVal bln As String)
    Dim sqlno As String, SqlS As String
    Dim rsno As New Recordset, rsS As New Recordset
    
    sqlno = " SELECT top 1 urut = rtrim(PO_Correction_no) from PurchaseOrder_Master_History " + _
        " where substring(rtrim(PO_Correction_no),5,8) = '" & Right(Trim$(thn), 4) & "." & Trim$(bln) & ".' and left(PO_Correction_no,4) = 'POC.' " & vbCrLf & _
        " order by urut desc "
        If rsno.State <> adStateClosed Then rsno.Close
    rsno.Open sqlno, Db, 1, 3
    
    If Not (rsno.BOF And rsno.EOF) Then
        CboPOC.Text = Left(Trim(rsno!Urut), 12) + Format(Str(Int(Right(Trim(rsno!Urut), 4)) + 1), "0000")
    Else
        CboPOC.Text = "POC." & Right(thn, 4) & "." & bln & "." & "0001"
    End If
    rsno.Close
    Set rsno = Nothing
End Sub

Sub adtocbopono()
Dim sqlno As String
Dim rsno As New Recordset
If Trim(txtName.Text) = "" Then Exit Sub
    sqlno = " select * From Purchaseorder_master " & vbCrLf & _
                           "Where period>='" & Format(SDate, "yyyymm") & "' And period<='" & Format(EDate, "yyyymm") & "' " & vbCrLf & _
                            IIf(cboSupp.Text = "ALL", "", "and  Supplier_Code='" & cboSupp.Text & "' ") & vbCrLf & _
                           ""
    Set rsno = Db.Execute(sqlno)
    With CboPOnO
        .clear
        Do While Not rsno.EOF
            .AddItem Trim(rsno("PO_No"))
            rsno.MoveNext
        Loop
        .ColumnWidths = "150pt"
        .ListWidth = 150
        .ListRows = 15
    End With
    Set rsno = Nothing
End Sub

Sub search()
Dim sqlno As String
Dim rsno As New Recordset
If Trim(txtName.Text) = "" Then Exit Sub
                    sqlno = " select * from Purchaseorder_master" & vbCrLf & _
                           "Where period>='" & Format(SDate, "yyyymm") & "' And period<='" & Format(EDate, "yyyymm") & "' " & vbCrLf & _
                            IIf(cboSupp.Text = "ALL", "", "and  Supplier_Code='" & cboSupp.Text & "' ") & vbCrLf & _
                           "and PO_no='" & CboPOnO.Text & "' " & vbCrLf & _
                           ""
    Set rsno = Db.Execute(sqlno)
    
         Do While Not rsno.EOF
            
                    cboPayment.Text = IIf(IsNull(rsno("PaymentTerm_cls")), "", Trim(rsno!PaymentTerm_Cls))
                    cboPriceCondition.Text = IIf(IsNull(rsno("PriceCondition_Cls")), "", Trim(rsno!PriceCondition_Cls))
                    CboPacking.Text = IIf(IsNull(rsno("POPacking_Cls")), "", Trim(rsno!POPacking_Cls))
                    cboTransport.Text = IIf(IsNull(rsno("Transportation_Cls")), "", Trim(rsno!Transportation_Cls))
            
        rsno.MoveNext
        Loop
End Sub

Sub adtocombo()
sql = "SELECT Trade_Code, Trade_Name FROM Trade_Master where trade_cls='2' or trade_cls='3'"
Set RS = Db.Execute(sql)

With cboSupp

.clear
.columnCount = 2
.ColumnWidths = "80 pt;300 pt"
.ListWidth = 380
.ListRows = 15
.AddItem ""
.List(0, 0) = strAll
.List(0, 1) = strAll
i = 1
Do Until RS.EOF
    .AddItem ""
    .List(i, 0) = Trim(RS!Trade_Code)
    .List(i, 1) = Trim(RS!trade_name)
    i = i + 1
    RS.MoveNext
Loop
.ListIndex = 0
End With
End Sub




Private Sub OptCreate_Click()
    Call blank
    Call PONO(Right(Year(EDate), 4), Format(Month(EDate), "0#"))
End Sub

Sub update()
sql = "SELECT * FROM PurchaseOrder_Master_History"
Set RS = Db.Execute(sql)

With CboPOC

.clear
.columnCount = 2
.ColumnWidths = "150 pt;150 pt"
.ListWidth = 300
.ListRows = 15
.AddItem ""
.List(0, 0) = ""
.List(0, 1) = ""
.List(0, 2) = ""
.List(0, 3) = ""
.List(0, 4) = ""
.List(0, 5) = ""
.List(0, 6) = ""
.List(0, 7) = ""
.List(0, 8) = ""

i = 1
Do Until RS.EOF
    .AddItem ""
    .List(i, 0) = IIf(IsNull(RS("PO_Correction_no")), "", Trim(RS!PO_Correction_no))
    .List(i, 1) = IIf(IsNull(RS("Supplier_Code")), "", Trim(RS!Supplier_Code))
    .List(i, 2) = IIf(IsNull(RS("po_no")), "", Trim(RS!po_no))
    .List(i, 3) = IIf(IsNull(RS("Period")), "", Trim(RS!Period))
    .List(i, 4) = IIf(IsNull(RS("PaymentTerm_cls")), "", Trim(RS!PaymentTerm_Cls))
    .List(i, 5) = IIf(IsNull(RS("PriceCondition_Cls")), "", Trim(RS!PriceCondition_Cls))
    .List(i, 6) = IIf(IsNull(RS("POPacking_Cls")), "", Trim(RS!POPacking_Cls))
    .List(i, 7) = IIf(IsNull(RS("Transportation_Cls")), "", Trim(RS!Transportation_Cls))
    .List(i, 8) = IIf(IsNull(RS("Reason")), "", Trim(RS!reason))

    i = i + 1
    RS.MoveNext
Loop
.ListIndex = 0
End With
End Sub

Private Sub OptUpdate_Click()
Call update
End Sub

Private Sub sdate_Change()
CboSupp_Click
End Sub

Sub blank()
    cboSupp.Text = ""
    txtName.Text = ""
    SDate.Value = Format(Now, "dd MMM YYYY")
    EDate.Value = Format(Now, "dd MMM YYYY")
    CboPOC.Text = ""
    CboPOnO.Text = ""
    cboPayment.Text = ""
    cboPriceCondition.Text = ""
    CboPacking.Text = ""
    cboTransport.Text = ""
    TxtPayment.Text = ""
    TxtPrice.Text = ""
    TxtPacking.Text = ""
    TxtTransport.Text = ""
    LblTotal.Caption = ""
    LblErr.Caption = ""
    TxtReason.Text = ""
End Sub

Sub Header()
  With grid
    .clear

    .Rows = 1
    .ColS = 6
    
    bteColItemCode = 0
    bteColDesciption = 1
    bteColOrderQty = 2
    bteColCur = 3
    bteColPrice = 4
    bteColAmount = 5

    .ColWidth(bteColItemCode) = 1300
    .ColWidth(bteColDesciption) = 3500
    .ColWidth(bteColOrderQty) = 1200
    .ColWidth(bteColCur) = 800
    .ColWidth(bteColPrice) = 1100
    .ColWidth(bteColAmount) = 1500

    .TextMatrix(0, bteColItemCode) = "Item Code"
    .TextMatrix(0, bteColDesciption) = "Description"
    .TextMatrix(0, bteColOrderQty) = "Order QTY"
    .TextMatrix(0, bteColCur) = "Cur"
    .TextMatrix(0, bteColPrice) = "Price"
    .TextMatrix(0, bteColAmount) = "Amount"
  End With
End Sub

Private Sub up_GridSearch()
    Dim li_count As Integer
    Dim ls_sql As String
    Dim RS As New ADODB.Recordset
    Dim li_Row As Integer
    Dim prg As Integer
   Header
   
    If OptCreate.Value = True Then
    ls_sql = uf_SQLSearch
    ElseIf OptUpdate.Value = True Then
    ls_sql = uf_SQLSearchUpdate
    End If
    
    If RS.State = adStateOpen Then RS.Close
    RS.CursorLocation = adUseClient
    RS.Open ls_sql, Db, adOpenDynamic, adLockOptimistic
    Set RS = Db.Execute(ls_sql)
    
    With grid
        While Not RS.EOF
            .Rows = .Rows + 1
            li_Row = .Rows - 1
            
            .TextMatrix(li_Row, bteColItemCode) = Trim(RS!Item_Code)
            .TextMatrix(li_Row, bteColDesciption) = Trim(RS!item_name)
            .TextMatrix(li_Row, bteColCur) = uf_GetCurrencyDescription(Trim(RS("Currency_code")))
            .TextMatrix(li_Row, bteColPrice) = Trim(RS!Price)
            
            If .TextMatrix(li_Row, bteColCur) = "IDR" Then
            .TextMatrix(li_Row, bteColOrderQty) = Format(Trim(RS!Qty), gs_formatQty)
            .TextMatrix(li_Row, bteColPrice) = Format(Trim(RS!Price), gs_formatPriceIDR)
            .TextMatrix(li_Row, bteColAmount) = Format(Trim(RS!Amount), gs_formatAmountIDR)
            
            Else
            .TextMatrix(li_Row, bteColOrderQty) = Format(Trim(RS!Qty), gs_formatQty)
            .TextMatrix(li_Row, bteColPrice) = Format(Trim(RS!Price), gs_formatPrice)
            .TextMatrix(li_Row, bteColAmount) = Format(Trim(RS!Amount), gs_formatAmount)
            
            End If
            DoEvents
'
            RS.MoveNext
        Wend
        
        'up_CheckHeader
        RS.Close
        LblTotal.Caption = "Total : " & .Rows - 1 & " Record (s)"
    End With
    
If grid.Rows > 1 Then
    grid.Cell(flexcpBackColor, 1, bteColPrice, grid.Rows - 1, bteColPrice) = &HFFFFFF
End If
    
End Sub
    
Sub insertupdate()
Dim RS As New ADODB.Recordset
Dim rsCek As New ADODB.Recordset
Dim ls_sql As String
'Dim ls_sqlupdate As String
Dim X As Double
    
'#insert header
    
If OptCreate.Value = True Then
    ls_sql = " insert into PurchaseOrder_Master_History " & vbCrLf & _
                      " select  '" & Trim(CboPOC.Text) & "',*,'" & Trim(TxtReason.Text) & "',NULL,NULL,NULL From PurchaseOrder_Master " & vbCrLf & _
                      " where po_no='" & Trim(CboPOnO.Text) & "' " & vbCrLf & _
                      " update PurchaseOrder_Master_History " & vbCrLf & _
                      " set PaymentTerm_Cls='" & Trim(cboPayment.Text) & "',PriceCondition_Cls='" & Trim(cboPriceCondition.Text) & "',POPacking_Cls='" & Trim(CboPacking.Text) & "',Transportation_Cls='" & Trim(cboTransport.Text) & "' " & vbCrLf & _
                      " where PO_Correction_no='" & Trim(CboPOC.Text) & "' " & vbCrLf & _
                      "  "
    Db.Execute (ls_sql)
'#insert Detail
    
    ls_sql = " insert into PurchaseOrder_Detail_History  " & vbCrLf & _
                      " select  '" & Trim(CboPOC.Text) & "',* From PurchaseOrder_Detail " & vbCrLf & _
                      " where po_no='" & CboPOnO.Text & "' " & vbCrLf
    Db.Execute ls_sql
    
    For X = 1 To grid.Rows - 1
    
    ls_sql = "  " & vbCrLf & _
                  " update PurchaseOrder_Detail_History " & vbCrLf & _
                  " set price=" & CDbl(grid.TextMatrix(X, bteColPrice)) & ", " & vbCrLf & _
                  " Amount=" & CDbl(grid.TextMatrix(X, bteColAmount)) & " " & vbCrLf & _
                  " where PO_Correction_no='" & Trim(CboPOC.Text) & "' and PO_No='" & CboPOnO.Text & "' and Item_Code='" & Trim(grid.TextMatrix(X, bteColItemCode)) & "' " & vbCrLf & _
                  " "
    Db.Execute ls_sql
    Next X
    LblErr = DisplayMsg(1000)
ElseIf OptUpdate.Value = True Then
    
    ls_sql = " update PurchaseOrder_Master_History " & vbCrLf & _
             " set PaymentTerm_Cls='" & Trim(cboPayment.Text) & "',PriceCondition_Cls='" & Trim(cboPriceCondition.Text) & "',POPacking_Cls='" & Trim(CboPacking.Text) & "',Transportation_Cls='" & Trim(cboTransport.Text) & "' " & vbCrLf & _
             " where PO_Correction_no='" & Trim(CboPOC.Text) & "' " & vbCrLf & _
             " "
    Db.Execute (ls_sql)
    
    For X = 1 To grid.Rows - 1
    ls_sql = " update PurchaseOrder_Detail_History " & vbCrLf & _
                  " set price=" & CDbl(grid.TextMatrix(X, bteColPrice)) & ", " & vbCrLf & _
                  " Amount=" & CDbl(grid.TextMatrix(X, bteColAmount)) & " " & vbCrLf & _
                  " where PO_Correction_no='" & Trim(CboPOC.Text) & "' and PO_No='" & CboPOnO.Text & "' and Item_Code='" & Trim(grid.TextMatrix(X, bteColItemCode)) & "' " & vbCrLf & _
                  " "
        Db.Execute (ls_sql)
    Next X
    LblErr = DisplayMsg(1101)
    
End If
End Sub

Private Function uf_SQLSearch() As String
uf_SQLSearch = " Select * from PurchaseOrder_Detail left join Item_Master on PurchaseOrder_Detail.item_code=Item_Master.item_code  " & vbCrLf & _
               " where PurchaseOrder_Detail.PO_No='" & CboPOnO.Text & "' "
End Function
 
Private Function uf_SQLSearchUpdate() As String
uf_SQLSearchUpdate = " Select * from PurchaseOrder_Detail_History left join Item_Master on PurchaseOrder_Detail_History.item_code=Item_Master.item_code  " & vbCrLf & _
                     " where PurchaseOrder_Detail_History.PO_No='" & CboPOnO.Text & "' "
End Function
