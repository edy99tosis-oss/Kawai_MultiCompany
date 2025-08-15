VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmProcessMaster 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Process Master"
   ClientHeight    =   10980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15120
   Icon            =   "FrmProcessMaster.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   10980
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text4 
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
      Height          =   315
      Left            =   12623
      TabIndex        =   8
      Text            =   "XX"
      Top             =   8400
      Width           =   2205
   End
   Begin VB.TextBox Text3 
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
      Height          =   315
      Left            =   8483
      Locked          =   -1  'True
      TabIndex        =   27
      TabStop         =   0   'False
      Text            =   "XX"
      Top             =   8400
      Width           =   2985
   End
   Begin VB.TextBox Text2 
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
      Height          =   315
      Left            =   5633
      TabIndex        =   5
      Text            =   "XX"
      Top             =   8400
      Width           =   1425
   End
   Begin VB.TextBox Text1 
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
      Height          =   315
      Left            =   2423
      Locked          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Text            =   "XX"
      Top             =   8400
      Width           =   3135
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   435
      Left            =   13043
      TabIndex        =   22
      Top             =   270
      Width           =   1845
      _extentx        =   3254
      _extenty        =   767
   End
   Begin VB.TextBox txtseq_no 
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
      Left            =   383
      MaxLength       =   2
      TabIndex        =   3
      Text            =   "XX"
      Top             =   8400
      Width           =   585
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
      Left            =   12563
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9750
      Width           =   1140
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
      Left            =   13763
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9750
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   600
      Left            =   293
      TabIndex        =   15
      Top             =   9000
      Width           =   14625
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
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   14265
      End
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
      Left            =   293
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9750
      Width           =   1140
   End
   Begin VB.CommandButton CmdCancel 
      BackColor       =   &H0080FFFF&
      Caption         =   "Ca&ncel"
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
      Left            =   11363
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   9750
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1605
      Left            =   293
      TabIndex        =   14
      Top             =   900
      Width           =   14625
      Begin VB.Line Line5 
         X1              =   6660
         X2              =   12240
         Y1              =   1350
         Y2              =   1350
      End
      Begin VB.Label LblPNameCopy 
         BackColor       =   &H00FDDFE3&
         Caption         =   "Product Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6660
         TabIndex        =   36
         Top             =   1110
         Width           =   5595
      End
      Begin VB.Line Line4 
         X1              =   6660
         X2              =   12240
         Y1              =   990
         Y2              =   990
      End
      Begin VB.Label LblPName 
         BackColor       =   &H00FDDFE3&
         Caption         =   "Product Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6660
         TabIndex        =   35
         Top             =   750
         Width           =   5595
      End
      Begin MSForms.ComboBox CboPCodeCopy 
         Height          =   315
         Left            =   2070
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1080
         Width           =   2355
         VariousPropertyBits=   612386843
         MaxLength       =   15
         DisplayStyle    =   3
         Size            =   "4154;556"
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label LblPCodeCopy 
         BackColor       =   &H00FDDFE3&
         Caption         =   "Product Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4530
         TabIndex        =   34
         Top             =   1125
         Width           =   1995
      End
      Begin VB.Line Line3 
         X1              =   4530
         X2              =   6510
         Y1              =   1350
         Y2              =   1350
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FDDFE3&
         Caption         =   "Copy Item Code"
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
         Left            =   210
         TabIndex        =   33
         Top             =   1110
         Width           =   1755
      End
      Begin VB.Line Line2 
         X1              =   4530
         X2              =   7920
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label LblPcls 
         BackColor       =   &H00FDDFE3&
         Caption         =   "Process Cls"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4530
         TabIndex        =   32
         Top             =   345
         Width           =   3495
      End
      Begin MSForms.ComboBox CboPCls 
         Height          =   315
         Left            =   2070
         TabIndex        =   0
         Top             =   300
         Width           =   2355
         VariousPropertyBits=   612386843
         MaxLength       =   2
         DisplayStyle    =   3
         Size            =   "4154;556"
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FDDFE3&
         Caption         =   "Group Cls"
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
         Left            =   210
         TabIndex        =   31
         Top             =   330
         Width           =   1095
      End
      Begin MSForms.ComboBox CboPCode 
         Height          =   315
         Left            =   2070
         TabIndex        =   1
         Top             =   690
         Width           =   2355
         VariousPropertyBits=   612386843
         MaxLength       =   15
         DisplayStyle    =   3
         Size            =   "4154;556"
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label LblPCode 
         BackColor       =   &H00FDDFE3&
         Caption         =   "Product Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4530
         TabIndex        =   18
         Top             =   750
         Width           =   1995
      End
      Begin VB.Line Line1 
         X1              =   4530
         X2              =   6510
         Y1              =   990
         Y2              =   990
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FDDFE3&
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
         Height          =   255
         Left            =   210
         TabIndex        =   17
         Top             =   720
         Width           =   1275
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid grid 
      Height          =   4890
      Left            =   293
      TabIndex        =   20
      Top             =   2670
      Width           =   14625
      _cx             =   25797
      _cy             =   8625
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
      FocusRect       =   5
      HighLight       =   2
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   3
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   0
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
      WordWrap        =   -1  'True
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
   Begin VB.Label Label11 
      BackColor       =   &H00A6D2FF&
      Caption         =   "Cost / Minute"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   12623
      TabIndex        =   30
      Top             =   8040
      Width           =   1425
   End
   Begin MSForms.ComboBox ComboBox3 
      Height          =   315
      Left            =   11543
      TabIndex        =   7
      Top             =   8400
      Width           =   975
      VariousPropertyBits=   612386843
      DisplayStyle    =   3
      Size            =   "1720;556"
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label10 
      BackColor       =   &H00A6D2FF&
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
      Height          =   225
      Left            =   11543
      TabIndex        =   29
      Top             =   8040
      Width           =   975
   End
   Begin VB.Label Label9 
      BackColor       =   &H00A6D2FF&
      Caption         =   "TradeName"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   8483
      TabIndex        =   28
      Top             =   8040
      Width           =   1215
   End
   Begin MSForms.ComboBox ComboBox2 
      Height          =   315
      Left            =   7193
      TabIndex        =   6
      Top             =   8400
      Width           =   1215
      VariousPropertyBits=   612386843
      DisplayStyle    =   3
      Size            =   "2143;556"
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label8 
      BackColor       =   &H00A6D2FF&
      Caption         =   "Trade Code"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   7193
      TabIndex        =   26
      Top             =   8040
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackColor       =   &H00A6D2FF&
      Caption         =   "Std. Time (min)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5693
      TabIndex        =   25
      Top             =   8040
      Width           =   1425
   End
   Begin VB.Label Label5 
      BackColor       =   &H00A6D2FF&
      Caption         =   "Process Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2423
      TabIndex        =   24
      Top             =   8040
      Width           =   1215
   End
   Begin MSForms.ComboBox ComboBox1 
      Height          =   315
      Left            =   1013
      TabIndex        =   4
      Top             =   8400
      Width           =   1290
      VariousPropertyBits=   612386843
      DisplayStyle    =   3
      Size            =   "2275;556"
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label4 
      BackColor       =   &H00A6D2FF&
      Caption         =   "Process Cls"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1073
      TabIndex        =   21
      Top             =   8040
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00A6D2FF&
      Caption         =   "Seq No"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   353
      TabIndex        =   19
      Top             =   8040
      Width           =   705
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00A6D2FF&
      BackStyle       =   1  'Opaque
      Height          =   360
      Index           =   1
      Left            =   293
      Top             =   7950
      Width           =   14655
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080C0FF&
      Height          =   855
      Index           =   0
      Left            =   293
      Top             =   7980
      Width           =   14655
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Process Master"
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
      Left            =   293
      TabIndex        =   13
      Top             =   270
      Width           =   14580
   End
End
Attribute VB_Name = "FrmProcessMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ColS As Byte
Dim ColSeq As Byte
Dim ColProcessCode As Byte
Dim ColProcessDesc As Byte
Dim ColStandardTime As Byte
Dim ColTradeCode As Byte
Dim ColTradeName As Byte
Dim ColCurrency As Byte
Dim ColCostMinute As Byte
Private Function uf_Validate() As Boolean
uf_Validate = False
If Trim(txtseq_no) = "" Then
    LblErrMsg = DisplayMsg(8076)
    Exit Function
End If
If uf_ValidateComboData(CboPCode, "4003", LblErrMsg, LblPCode) = False Then Exit Function
If uf_ValidateComboData(ComboBox1, "8068", LblErrMsg, Text1) = False Then Exit Function
If IsNumeric(Text2) = False Then
    LblErrMsg = DisplayMsg(8072)
    Exit Function
End If

If CDbl(Text2) = 0 Then
    LblErrMsg = DisplayMsg(8072)
    Exit Function
End If

If CDbl(Text2) > gd_MaxTime Then
    LblErrMsg = DisplayMsg(8074) & " " & gd_MaxTime
    Exit Function
End If
'If uf_ValidateComboData(ComboBox2, "4013", LblErrMsg, Text3) = False Then Exit Function

If Trim(ComboBox2.Text) = "" Then
    If Trim(ComboBox3) = "" Then
        LblErrMsg = DisplayMsg(1028)
        Exit Function
    End If
    Text4 = 0
    If CDbl(Text4) > gd_MaxAmount Then
        LblErrMsg = DisplayMsg(8075) & " " & gd_MaxAmount
        Exit Function
    End If
    If IsNumeric(Text4) = False Then
    LblErrMsg = DisplayMsg(8073)
    Exit Function
End If
End If
uf_Validate = True
End Function

Private Sub CboPCls_Change()
If CboPCls.ListIndex < 0 Then Exit Sub
LblPcls = CboPCls.List(CboPCls.ListIndex, 1)
Call up_SettingComboItemMaster
End Sub
Private Sub CboPCode_Change()
Call up_HeaderGrid
If CboPCode.ListIndex < 0 Then Exit Sub
LblPCode = CboPCode.List(CboPCode.ListIndex, 1)
LblPName = CboPCode.List(CboPCode.ListIndex, 2)
Call up_SettingGrid
End Sub
Private Sub CboPCodeCopy_Change()
Call up_HeaderGrid
If CboPCodeCopy.ListIndex < 0 Then Exit Sub
LblPCodeCopy = CboPCodeCopy.List(CboPCodeCopy.ListIndex, 1)
LblPNameCopy = CboPCodeCopy.List(CboPCodeCopy.ListIndex, 2)
Call up_CopyData
End Sub

Private Sub up_CopyData()
If (MsgBox(DisplayMsg(8079), vbYesNo, "Confirmation")) = vbYes Then
    Dim ls_sql As String
    Db.BeginTrans
    ls_sql = " delete from [Process_Master] " & _
                  "  WHERE [Item_Code] = '" & Trim(CboPCode) & "'"
    Db.Execute ls_sql
   ls_sql = " INSERT INTO [Process_Master] " & _
                  "            ([Item_Code] " & _
                  "            ,[SeqNo] " & _
                  "            ,[Process_Cls] " & _
                  "            ,[Standard_Time] " & _
                  "            ,[Trade_Code] " & _
                  "            ,[Currency_Code] " & _
                  "            ,[Cost_Minute]) " & _
                  "       " & _
                  "  ( select "
    
    ls_sql = ls_sql + " '" & Trim(CboPCode) & "' item_Code, " & _
                  "            [SeqNo] " & _
                  "            ,[Process_Cls] " & _
                  "            ,[Standard_Time] " & _
                  "            ,[Trade_Code] " & _
                  "            ,[Currency_Code] " & _
                  "            ,[Cost_Minute] " & _
                  "  from [Process_Master] " & _
                  "  WHERE [Item_Code] = '" & Trim(CboPCodeCopy) & "' )"
                  
    Db.Execute ls_sql
    Db.CommitTrans
End If
CboPCodeCopy = ""
LblPCodeCopy = ""
LblPNameCopy = ""
Call up_SettingGrid
LblErrMsg = DisplayMsg(8080)
End Sub

Private Sub CmdSubMenu_Click()
Unload Me
frmMainMenu.Show
End Sub

Private Sub ComboBox2_Change()
up_EnableCost
End Sub
Private Sub up_EnableCost()
If Trim(ComboBox2) = "" Then
    ComboBox3.Enabled = True
    Text4.Enabled = True
    Text3.Text = ""
Else
    ComboBox3.Enabled = False
    Text4.Enabled = False
    Text4 = ""
    ComboBox3 = ""
End If
End Sub

Private Sub ComboBox2_LostFocus()
up_EnableCost
If ComboBox3.Enabled = True Then
ComboBox3.SetFocus
End If
End Sub

Private Sub ComboBox3_KeyPress(KeyAscii As MSForms.ReturnInteger)
KeyAscii = 0
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
CtrlMenu1.FormName = Me.Name
Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
ColS = 0
ColSeq = 1
ColProcessCode = 2
ColProcessDesc = 3
ColStandardTime = 4
ColTradeCode = 5
ColTradeName = 6
ColCurrency = 7
ColCostMinute = 8
Call up_SettingCombo
Call up_ClearAll
End Sub

Private Sub up_ClearAll()
CboPCode.Text = ""
CboPCls.Text = strAll
CboPCodeCopy.Text = ""

LblPCode = ""
LblPName = ""
LblPCodeCopy = ""
LblPNameCopy = ""
LblPcls = strAll

'up_SettingComboItemMaster
up_HeaderGrid
up_ClearInputArea
End Sub
Private Sub up_ClearInputArea()
txtseq_no.Text = ""
LblErrMsg = ""
ComboBox1 = ""
ComboBox1.Enabled = True
ComboBox2 = ""
ComboBox3 = ""
ComboBox3.Enabled = False
Text4.Enabled = False
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
End Sub

Private Sub up_SettingCombo()

Call up_FillCombo(CboPCls, "Group_CLs", "*", "", True)
CboPCls.ColumnWidths = "40 pt;120 pt"
CboPCls.ListWidth = 160
CboPCls.ListRows = 15
CboPCls.ListIndex = 0

'Call up_SettingComboItemMaster

DoEvents
Call up_FillCombo(ComboBox1, "Process_Cls")
ComboBox1.ColumnWidths = "75 pt;250 pt"
ComboBox1.ListWidth = 325
ComboBox1.ListRows = 15

Call up_FillCombo(ComboBox2, "Trade_Master", "Trade_Code,Trade_Name", " where trade_cls='3' ")
ComboBox2.ColumnWidths = "75 pt;250 pt"
ComboBox2.ListWidth = 325
ComboBox2.ListRows = 15

Call up_FillCombo(ComboBox3, "curr_cls", "description,curr_cls")
ComboBox3.ColumnWidths = "60 pt;0 pt"

End Sub

Private Sub up_SettingComboItemMaster()

Call up_FillCombo(CboPCode, "Item_Master", "*", IIf(Trim(CboPCls) <> strAll, " where group_cls= '" & Trim(CboPCls) & "' ", ""))
CboPCode.ColumnWidths = "75 pt;250 pt"
CboPCode.ListWidth = 325
CboPCode.ListRows = 15

    Dim rscbo As New ADODB.Recordset 'Isi Combo
    
    '******** Isi Combo *********
    sql = "select Item_Code as a,MakerItem_Code b,Item_Name c " & _
        "from Item_Master "
    
    If Trim(CboPCls) <> strAll Then sql = sql & " where group_cls = '" & Trim(CboPCls) & "' "
    
    sql = sql & "order by Item_Code"
    Set rscbo = Db.Execute(sql)
     
    '**** Isi combo Parent dan Item Code
    CboPCode.clear
    CboPCode.columnCount = 3
    CboPCode.TextColumn = 1
    
    CboPCodeCopy.clear
    CboPCodeCopy.columnCount = 3
    CboPCodeCopy.TextColumn = 1
    
    i = 0
    Do While Not (rscbo.EOF)
        CboPCode.AddItem ""
        CboPCode.List(i, 0) = Trim(rscbo("a"))
        CboPCode.List(i, 1) = Trim(rscbo("b"))
        CboPCode.List(i, 2) = Trim(rscbo("c"))
        
        CboPCodeCopy.AddItem ""
        CboPCodeCopy.List(i, 0) = Trim(rscbo("a"))
        CboPCodeCopy.List(i, 1) = Trim(rscbo("b"))
        CboPCodeCopy.List(i, 2) = Trim(rscbo("c"))
        
        i = i + 1
        rscbo.MoveNext
     Loop
     CboPCode.ListWidth = 400
     CboPCode.ColumnWidths = "80 pt;80 pt;240 pt"
     CboPCode.ListRows = 15
     
     CboPCodeCopy.ListWidth = 400
     CboPCodeCopy.ColumnWidths = "80 pt;80 pt;240 pt"
     CboPCodeCopy.ListRows = 15
     
     '************************
     Set rscbo = Nothing

End Sub

Private Sub up_HeaderGrid()
grid.clear
grid.RowHeightMin = 250
grid.Rows = 1
grid.ColS = 9

grid.TextMatrix(0, ColS) = "S"
grid.TextMatrix(0, ColSeq) = "Seq No"
grid.TextMatrix(0, ColProcessCode) = "Process Cls"
grid.TextMatrix(0, ColProcessDesc) = "Process Name"
grid.TextMatrix(0, ColStandardTime) = "Standard Time (min)"
grid.TextMatrix(0, ColTradeCode) = "Trade Code"
grid.TextMatrix(0, ColTradeName) = "Trade Name"
grid.TextMatrix(0, ColCurrency) = "Currency"
grid.TextMatrix(0, ColCostMinute) = "Cost / Minute"

grid.ColWidth(ColS) = 300
grid.ColWidth(ColSeq) = 810
grid.ColWidth(ColProcessCode) = 1530
grid.ColWidth(ColProcessDesc) = 2460
grid.ColWidth(ColStandardTime) = 1530
grid.ColWidth(ColTradeCode) = 1320
grid.ColWidth(ColTradeName) = 3000
grid.ColWidth(ColCurrency) = 960
grid.ColWidth(ColCostMinute) = 2250

grid.ColAlignment(ColS) = flexAlignCenterCenter
grid.ColAlignment(ColSeq) = flexAlignLeftCenter
grid.ColAlignment(ColProcessCode) = flexAlignLeftCenter
grid.ColAlignment(ColProcessDesc) = flexAlignLeftCenter
grid.ColAlignment(ColStandardTime) = flexAlignRightCenter
grid.ColAlignment(ColTradeCode) = flexAlignLeftCenter
grid.ColAlignment(ColTradeName) = flexAlignLeftCenter
grid.ColAlignment(ColCurrency) = flexAlignLeftCenter
grid.ColAlignment(ColCostMinute) = flexAlignRightCenter

grid.Cell(flexcpAlignment, 0, 0, 0, 1) = flexAlignCenterCenter
End Sub

Private Sub CboPCls_KeyPress(KeyAscii As MSForms.ReturnInteger)
If Chr(KeyAscii) = "'" Then KeyAscii = 0
End Sub
Private Sub CboPCode_KeyPress(KeyAscii As MSForms.ReturnInteger)
If Chr(KeyAscii) = "'" Then KeyAscii = 0
End Sub
Private Sub CboPCodeCopy_KeyPress(KeyAscii As MSForms.ReturnInteger)
If Chr(KeyAscii) = "'" Then KeyAscii = 0
End Sub
Private Sub Combobox1_KeyPress(KeyAscii As MSForms.ReturnInteger)
If Chr(KeyAscii) = "'" Then KeyAscii = 0
End Sub
Private Sub Combobox2_KeyPress(KeyAscii As MSForms.ReturnInteger)
If Chr(KeyAscii) = "'" Then KeyAscii = 0
End Sub

Private Sub CboPCls_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 13 Then
    Call uf_ValidateComboData(CboPCls, "8056", LblErrMsg, LblPcls)
End If
End Sub
Private Sub CboPCode_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 13 Then
    Call uf_ValidateComboData(CboPCode, "4003", LblErrMsg, LblPCode)
    Call up_SettingGrid
End If
End Sub
Private Sub CboPCodeCopy_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 13 Then
    If uf_ValidateComboData(CboPCodeCopy, "4003", LblErrMsg, LblPCodeCopy) = True Then
        Call up_CopyData
    End If
End If
End Sub
Private Sub Combobox1_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 13 Then
    Call uf_ValidateComboData(ComboBox1, "8068", LblErrMsg, Text1)
End If
End Sub
Private Sub Combobox2_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 13 Then
    If Trim(ComboBox2) <> "" Then
        Call uf_ValidateComboData(ComboBox2, "4013", LblErrMsg, Text3)
    End If
End If
End Sub

Private Sub cmdCancel_Click()
For i = 1 To grid.Rows - 1
    grid.TextMatrix(i, ColS) = ""
Next
'Sta = ""
up_ClearInputArea
End Sub

Private Sub cmdClear_Click()
up_ClearAll
End Sub

Private Sub up_SettingGrid()
Call up_HeaderGrid
Dim rsisi As New ADODB.Recordset
Dim i As Long
Dim ls_sql As String
ls_sql = " select e.Description CurrDesc,b.Item_Name,c.Trade_Name,d.Description ProcessName,a.* from Process_Master a " & _
            "  left join Item_Master b on a.item_code=b.item_code " & _
            "  left join Trade_Master c on a.Trade_code=c.Trade_code " & _
            "  left join Process_Cls d on a.Process_Cls=d.process_Cls " & _
            "  left join Curr_Cls e on a.Currency_Code=e.Curr_Cls " & _
             " where a.item_code='" & Trim(CboPCode) & "' order by a.Seqno,a.Process_Cls"
            
If rsisi.State <> adStateClosed Then rsisi.Close
rsisi.Open ls_sql, Db, adOpenKeyset, adLockOptimistic
If Not rsisi.EOF And Not rsisi.BOF Then
    rsisi.MoveFirst
    i = 1
    While Not rsisi.EOF
        grid.AddItem ""
        grid.TextMatrix(i, ColS) = ""
        grid.TextMatrix(i, ColSeq) = Trim(rsisi!seqNo)
        grid.TextMatrix(i, ColProcessCode) = Trim(rsisi!Process_Cls)
        grid.TextMatrix(i, ColProcessDesc) = Trim(rsisi!ProcessName)
        grid.TextMatrix(i, ColStandardTime) = Trim(rsisi!Standard_Time)
        grid.TextMatrix(i, ColTradeCode) = IIf(IsNull(rsisi!Trade_Code), "", Trim(rsisi!Trade_Code))
        grid.TextMatrix(i, ColTradeName) = IIf(IsNull(rsisi!trade_name), "", Trim(rsisi!trade_name))
        grid.TextMatrix(i, ColCurrency) = IIf(IsNull(rsisi!CurrDesc), "", Trim(rsisi!CurrDesc))
        grid.TextMatrix(i, ColCostMinute) = IIf(IsNull(rsisi!cost_minute), "", Format(rsisi!cost_minute, gs_formatAmount))

        grid.Cell(flexcpAlignment, 1, 0, i, ColS) = flexAlignCenterCenter
        'grid.Cell(flexcpAlignment, 1, 1, i, 1) = flexAlignLeftCenter
        grid.Cell(flexcpBackColor, grid.Rows - 1, ColS) = vbWhite
        rsisi.MoveNext
        i = i + 1
    Wend
End If
If rsisi.State <> adStateClosed Then rsisi.Close
End Sub

Private Sub CmdSubmit_Click()
Me.MousePointer = vbHourglass
Dim i As Integer
Dim status As String
Dim ls_sql As String
status = "insert"
For i = 1 To grid.Rows - 1
    If grid.TextMatrix(i, ColS) = "D" Then
        status = "delete"
        Exit For
    End If
    If grid.TextMatrix(i, ColS) = "S" Then
        status = "update"
        Exit For
    End If
Next


If status = "insert" Then
    If uf_Validate = False Then
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    Dim RS As New ADODB.Recordset
    If RS.State <> adStateClosed Then RS.Close
    RS.Open " Select * from [Process_Master] " & _
                    "  WHERE [Item_Code] = '" & Trim(CboPCode) & "'  and  " & _
                    " [Process_Cls] = '" & Trim(ComboBox1) & "' ", Db, adOpenKeyset, adLockOptimistic
    If RS.EOF = False Then
        LblErrMsg = DisplayMsg(1023)
        Me.MousePointer = vbDefault
        If RS.State <> adStateClosed Then RS.Close
        Exit Sub
    End If
    If RS.State <> adStateClosed Then RS.Close
    ls_sql = " INSERT INTO [Process_Master] " & _
                  "            ([Item_Code] " & _
                  "            ,[SeqNo] " & _
                  "            ,[Process_Cls] " & _
                  "            ,[Standard_Time] " & _
                  "            ,[Trade_Code] " & _
                  "            ,[Currency_Code] " & _
                  "            ,[Cost_Minute]) " & _
                  "      VALUES " & _
                  "  ( "
    
    ls_sql = ls_sql + " '" & Trim(CboPCode) & "' " & _
                ", '" & Trim(txtseq_no) & "' " & _
                ", '" & Trim(ComboBox1) & "' " & _
                ", " & CDbl(Text2) & " " & _
                ", " & IIf(Trim(ComboBox2) = "", "Null", "'" & Trim(ComboBox2) & "'") & " " & _
                ", " & IIf(Trim(ComboBox2) <> "", "Null", "'" & uf_GetCurrencyCode(Trim(ComboBox3)) & "'") & " " & _
                ", " & IIf(Trim(ComboBox2) <> "", "Null", Text4) & " ) "
    Db.Execute ls_sql
    Call up_ClearInputArea
    LblErrMsg = DisplayMsg(1000)
ElseIf status = "update" Then
    If uf_Validate = False Then
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    Dim Cost As Double
    If Trim(ComboBox2) <> "" Then
        Cost = 0
    Else
        Cost = CDbl(Text4)
    End If
    ls_sql = " UPDATE [Process_Master] " & _
                  "    SET  " & _
                  "       [SeqNo] = '" & Trim(txtseq_no) & "'       " & _
                  "       ,[Standard_Time] =  " & Trim(Text2) & "    " & _
                  "       ,[Trade_Code] =  " & IIf(Trim(ComboBox2) = "", "Null", "'" & Trim(ComboBox2) & "'") & "   " & _
                  "       ,[Currency_Code] =  " & IIf(Trim(ComboBox2) <> "", "Null", "'" & uf_GetCurrencyCode(Trim(ComboBox3)) & "'") & "    " & _
                  "       ,[Cost_Minute] =  " & IIf(Trim(ComboBox2) <> "", "Null", Cost) & "    " & _
                  "  WHERE [Item_Code] = '" & Trim(CboPCode) & "'  and  [Process_Cls] = '" & Trim(ComboBox1) & "' "
    Db.Execute ls_sql
    Call up_ClearInputArea
    LblErrMsg = DisplayMsg(1101)
ElseIf status = "delete" Then
    If MsgBox(DisplayMsg(8071), vbYesNo, "Confirmation") = vbYes Then '#Are you sure want to delete data ?
            For i = 1 To grid.Rows - 1
                If grid.TextMatrix(i, ColS) = "D" Then
                        ls_sql = " delete from [Process_Master] " & _
                                      "  WHERE [Item_Code] = '" & Trim(CboPCode) & "'  and  [Process_Cls] = '" & Trim(grid.TextMatrix(i, ColProcessCode)) & "' "
                        Db.Execute ls_sql
                End If
            Next
            Call up_ClearInputArea
            LblErrMsg = DisplayMsg(1201)
    End If
End If
up_SettingGrid
  Me.MousePointer = vbDefault
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Col = ColS Then
    If grid.Text = "S" Then
        txtseq_no.Text = grid.TextMatrix(Row, ColSeq)
        ComboBox1.Text = grid.TextMatrix(Row, ColProcessCode)
        Text1.Text = grid.TextMatrix(Row, ColProcessDesc)
        ComboBox2.Text = grid.TextMatrix(Row, ColTradeCode)
        Text2.Text = grid.TextMatrix(Row, ColStandardTime)
        Text3.Text = grid.TextMatrix(Row, ColTradeName)
        ComboBox3.Text = grid.TextMatrix(Row, ColCurrency)
        Text4 = grid.TextMatrix(Row, ColCostMinute)
        ComboBox1.Enabled = False
        up_EnableCost
    End If
End If
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
grid.EditMaxLength = 1
If Col <> ColS Then Cancel = True
End Sub

Private Sub grid_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) <> "S" And Chr(KeyAscii) <> "s" And Chr(KeyAscii) <> "D" And Chr(KeyAscii) <> "d" Then
    KeyAscii = 0
    Exit Sub
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
Dim i As Integer
If Chr(KeyAscii) = "D" Then
    For i = 1 To grid.Rows - 1
        If Trim(grid.TextMatrix(i, ColS)) <> "D" Then
            grid.TextMatrix(i, ColS) = ""
        End If
    Next
    Call up_ClearInputArea
End If
If Chr(KeyAscii) = "S" Then
    For i = 1 To grid.Rows - 1
            grid.TextMatrix(i, ColS) = ""
    Next
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If InStr(1, "0123465798.", Chr(KeyAscii)) < 1 And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If InStr(1, "0123465798.", Chr(KeyAscii)) < 1 And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub
