VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmPackingMaster 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Packing Master"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15120
   Icon            =   "FrmPackingMaster.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   10950
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text8 
      Alignment       =   1  'Right Justify
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
      Left            =   13410
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "XX"
      Top             =   8400
      Width           =   1425
   End
   Begin VB.TextBox Text7 
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
      Left            =   11910
      TabIndex        =   7
      Text            =   "XX"
      Top             =   8400
      Width           =   1425
   End
   Begin VB.TextBox Text6 
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
      Left            =   10410
      TabIndex        =   6
      Text            =   "XX"
      Top             =   8400
      Width           =   1425
   End
   Begin VB.TextBox Text5 
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
      Left            =   8910
      TabIndex        =   5
      Text            =   "XX"
      Top             =   8400
      Width           =   1425
   End
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
      Left            =   7410
      TabIndex        =   4
      Text            =   "XX"
      Top             =   8400
      Width           =   1425
   End
   Begin VB.TextBox Text3 
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
      Left            =   5910
      TabIndex        =   3
      Text            =   "XX"
      Top             =   8400
      Width           =   1425
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
      Left            =   4410
      TabIndex        =   2
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
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Text            =   "XX"
      Top             =   8400
      Width           =   2505
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   435
      Left            =   13043
      TabIndex        =   21
      Top             =   270
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   767
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
      TabIndex        =   10
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
      TabIndex        =   11
      Top             =   9750
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   885
      Left            =   300
      TabIndex        =   14
      Top             =   900
      Width           =   14625
      Begin VB.CommandButton cmdBrowser 
         Caption         =   "..."
         Height          =   300
         Left            =   3420
         TabIndex        =   34
         Top             =   322
         Width           =   300
      End
      Begin VB.Label LblDesc 
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
         Height          =   195
         Left            =   9000
         TabIndex        =   32
         Top             =   375
         Width           =   4470
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FDDFE3&
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
         Index           =   2
         Left            =   7710
         TabIndex        =   31
         Top             =   375
         Width           =   960
      End
      Begin VB.Line Line2 
         X1              =   9000
         X2              =   13455
         Y1              =   615
         Y2              =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FDDFE3&
         Caption         =   "Part Number"
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
         Left            =   4110
         TabIndex        =   30
         Top             =   375
         Width           =   1080
      End
      Begin MSForms.ComboBox CboPCode 
         Height          =   315
         Left            =   1530
         TabIndex        =   0
         Top             =   315
         Width           =   1815
         VariousPropertyBits=   612386843
         MaxLength       =   15
         DisplayStyle    =   3
         Size            =   "3201;556"
         MatchEntry      =   1
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
         Height          =   195
         Left            =   5520
         TabIndex        =   18
         Top             =   375
         Width           =   1815
      End
      Begin VB.Line Line1 
         X1              =   5520
         X2              =   7320
         Y1              =   615
         Y2              =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   17
         Top             =   375
         Width           =   915
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid grid 
      Height          =   5610
      Left            =   300
      TabIndex        =   19
      Top             =   1950
      Width           =   14625
      _cx             =   25797
      _cy             =   9895
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
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00A6D2FF&
      Caption         =   "Volume (M3)"
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
      Left            =   13740
      TabIndex        =   33
      Top             =   8010
      Width           =   1095
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00A6D2FF&
      Caption         =   "Qty / Ctn"
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
      Left            =   12540
      TabIndex        =   29
      Top             =   8010
      Width           =   795
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00A6D2FF&
      Caption         =   "Height"
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
      Left            =   11280
      TabIndex        =   28
      Top             =   8010
      Width           =   540
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00A6D2FF&
      Caption         =   "Width"
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
      Left            =   9855
      TabIndex        =   27
      Top             =   8010
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00A6D2FF&
      Caption         =   "Length"
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
      Left            =   8265
      TabIndex        =   26
      Top             =   8010
      Width           =   570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00A6D2FF&
      Caption         =   "Net Weight"
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
      Left            =   6405
      TabIndex        =   25
      Top             =   8010
      Width           =   930
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00A6D2FF&
      Caption         =   "Gross Weight"
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
      Left            =   4695
      TabIndex        =   24
      Top             =   8010
      Width           =   1140
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00A6D2FF&
      Caption         =   "Packing Type Name"
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
      Left            =   1800
      TabIndex        =   23
      Top             =   8010
      Width           =   1695
   End
   Begin MSForms.ComboBox ComboBox1 
      Height          =   315
      Left            =   420
      TabIndex        =   1
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
      AutoSize        =   -1  'True
      BackColor       =   &H00A6D2FF&
      Caption         =   "Packing Style"
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
      Left            =   420
      TabIndex        =   20
      Top             =   8010
      Width           =   1155
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
      Caption         =   "Packing Master"
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
Attribute VB_Name = "FrmPackingMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ColS As Byte
Dim ColPackingStyle As Byte
Dim ColPackingStyleDesc As Byte
Dim ColGrossWeight As Byte
Dim ColNetWeight As Byte
Dim ColLength As Byte
Dim ColWidth As Byte
Dim ColThickness As Byte
Dim ColCtn As Byte
Dim ColVol As Byte

Private Function uf_Validate() As Boolean
uf_Validate = False

If uf_ValidateComboData(CboPCode, "4003", LblErrMsg, lbldesc) = False Then Exit Function
If uf_ValidateComboData(ComboBox1, "0024", LblErrMsg, Text1) = False Then Exit Function

If IsNumeric(Text2) = False Then
    LblErrMsg = DisplayMsg(8102)
    Exit Function
End If

If CDbl(Text2) = 0 Then
    LblErrMsg = DisplayMsg(8102)
    Exit Function
End If

If CDbl(Text2) > gd_MaxWeight Then
    LblErrMsg = DisplayMsg(8104) & " " & gd_MaxWeight
    Exit Function
End If
'---------------------------------------------------------------------------------
If IsNumeric(Text3) = False Then
    LblErrMsg = DisplayMsg(8103)
    Exit Function
End If

If CDbl(Text3) = 0 Then
    LblErrMsg = DisplayMsg(8103)
    Exit Function
End If

If CDbl(Text3) > gd_MaxWeight Then
    LblErrMsg = DisplayMsg(8105) & " " & gd_MaxWeight
    Exit Function
End If
'---------------------------------------------------------------------------------
If IsNumeric(Text4) = False Then
    LblErrMsg = DisplayMsg(8031)
    Exit Function
End If

If CDbl(Text4) = 0 Then
    LblErrMsg = DisplayMsg(8031)
    Exit Function
End If

If CDbl(Text4) > gd_MaxLength Then
    LblErrMsg = DisplayMsg(8032) & " " & gd_MaxLength
    Exit Function
End If
'---------------------------------------------------------------------------------
If IsNumeric(Text5) = False Then
    LblErrMsg = DisplayMsg(8028)
    Exit Function
End If

If CDbl(Text5) = 0 Then
    LblErrMsg = DisplayMsg(8028)
    Exit Function
End If

If CDbl(Text5) > gd_MaxWidth Then
    LblErrMsg = DisplayMsg(8029) & " " & gd_MaxWidth
    Exit Function
End If

'---------------------------------------------------------------------------------
If IsNumeric(Text6) = False Then
    LblErrMsg = DisplayMsg(8026)
    Exit Function
End If

If CDbl(Text6) = 0 Then
    LblErrMsg = DisplayMsg(8026)
    Exit Function
End If

If CDbl(Text6) > gd_MaxThickness Then
    LblErrMsg = DisplayMsg(8027) & " " & gd_MaxThickness
    Exit Function
End If

'---------------------------------------------------------------------------------
If IsNumeric(Text7) = False Then
    LblErrMsg = DisplayMsg(8106)
    Exit Function
End If

If CDbl(Text7) = 0 Then
    LblErrMsg = DisplayMsg(8106)
    Exit Function
End If

If CDbl(Text7) > gd_MaxQty Then
    LblErrMsg = DisplayMsg(4037) & " " & gd_MaxQty
    Exit Function
End If
uf_Validate = True
End Function

Private Sub CboPCode_Change()
Call up_HeaderGrid
If CboPCode.ListIndex < 0 Then Exit Sub
LblPCode = CboPCode.List(CboPCode.ListIndex, 1)
lbldesc = CboPCode.List(CboPCode.ListIndex, 2)
Call up_SettingGrid
End Sub

Private Sub cmdBrowser_Click()
 Me.MousePointer = vbHourglass
 frm_BrowseItem.getItemCode = CboPCode.Text
 frm_BrowseItem.Show 1
 CboPCode.Text = frm_BrowseItem.getItemCode
 Me.MousePointer = vbDefault
End Sub

Private Sub CmdSubMenu_Click()
Unload Me
frmMainMenu.Show
End Sub

Private Sub ComboBox1_Click()
Text1 = ComboBox1.List(ComboBox1.ListIndex, 1)
End Sub

Private Sub Form_Load()
 If gb_Simulation = True Then Call up_InitSimulation(Me)
CtrlMenu1.FormName = Me.Name
Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
ColS = 0
ColPackingStyle = 1
ColPackingStyleDesc = 2
ColGrossWeight = 3
ColNetWeight = 4
ColLength = 5
ColWidth = 6
ColThickness = 7
ColCtn = 8
ColVol = 9

Call up_SettingCombo
Call up_ClearAll
End Sub

Private Sub up_ClearAll()
CboPCode.Text = ""
LblPCode = ""
lbldesc = ""
up_SettingComboItemMaster
up_HeaderGrid
up_ClearInputArea
End Sub
Private Sub up_ClearInputArea()
LblErrMsg = ""
ComboBox1 = ""
ComboBox1.Enabled = True

Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Text7 = ""
Text8 = ""
End Sub

Private Sub up_SettingCombo()
Call up_SettingComboItemMaster
Call up_FillCombo(ComboBox1, "PackingStyle_Cls")
ComboBox1.ColumnWidths = "75 pt;250 pt"
ComboBox1.ListWidth = 325
ComboBox1.ListRows = 15
End Sub

Private Sub up_SettingComboItemMaster()
Dim adoRs As New ADODB.Recordset

'Call up_FillCombo(CboPCode, "Item_Master", "  * ")
'CboPCode.ColumnWidths = "75 pt;250 pt"
'CboPCode.ListWidth = 325
'CboPCode.ListRows = 15

With CboPCode
    
    .clear
    .columnCount = 3
    
    sql = "select item_code, makeritem_code, item_name from item_master where use_endday >= convert(char(8), getdate(), 112)" ' and finishgoodpart_cls = '01'"
    adoRs.Open sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
    While Not adoRs.EOF
        .AddItem ""
        .List(.ListCount - 1, 0) = Trim(adoRs.Fields("item_code"))
        .List(.ListCount - 1, 1) = Trim(adoRs.Fields("makeritem_code"))
        .List(.ListCount - 1, 2) = Trim(adoRs.Fields("item_name"))
        adoRs.MoveNext
    Wend
    adoRs.Close
    
    .ListWidth = 350
    .ColumnWidths = "75 pt;75 pt;200 pt"
    
End With

Set adoRs = Nothing
End Sub

Private Sub up_HeaderGrid()
grid.clear
grid.RowHeightMin = 250
grid.Rows = 1
grid.ColS = 10

grid.TextMatrix(0, ColS) = "S"
grid.TextMatrix(0, ColPackingStyle) = "Packing Style"
grid.TextMatrix(0, ColPackingStyleDesc) = "Packing Type"
grid.TextMatrix(0, ColGrossWeight) = "Gross Weight"
grid.TextMatrix(0, ColNetWeight) = "Net Weight"
grid.TextMatrix(0, ColWidth) = "Width"
grid.TextMatrix(0, ColLength) = "Length"
grid.TextMatrix(0, ColThickness) = "Height"
grid.TextMatrix(0, ColCtn) = "Qty / Ctn"
grid.TextMatrix(0, ColVol) = "Volume (M3)"

grid.ColWidth(ColS) = 300
grid.ColWidth(ColPackingStyle) = 1600
grid.ColWidth(ColPackingStyleDesc) = 1600
grid.ColWidth(ColNetWeight) = 1500
grid.ColWidth(ColGrossWeight) = 1500
grid.ColWidth(ColWidth) = 1500
grid.ColWidth(ColLength) = 1500
grid.ColWidth(ColThickness) = 1500
grid.ColWidth(ColCtn) = 1500
grid.ColWidth(ColVol) = 1500

grid.ColAlignment(ColS) = flexAlignCenterCenter
grid.ColAlignment(ColPackingStyle) = flexAlignLeftCenter
grid.ColAlignment(ColPackingStyleDesc) = flexAlignLeftCenter
grid.ColAlignment(ColNetWeight) = flexAlignRightCenter
grid.ColAlignment(ColGrossWeight) = flexAlignRightCenter
grid.ColAlignment(ColWidth) = flexAlignRightCenter
grid.ColAlignment(ColThickness) = flexAlignRightCenter
grid.ColAlignment(ColLength) = flexAlignRightCenter
grid.ColAlignment(ColCtn) = flexAlignRightCenter
grid.ColAlignment(ColVol) = flexAlignRightCenter
grid.Cell(flexcpAlignment, 0, 0, 0, 1) = flexAlignCenterCenter

End Sub


Private Sub CboPCode_KeyPress(KeyAscii As MSForms.ReturnInteger)
If Chr(KeyAscii) = "'" Then KeyAscii = 0
End Sub

Private Sub Combobox1_KeyPress(KeyAscii As MSForms.ReturnInteger)
If Chr(KeyAscii) = "'" Then KeyAscii = 0
End Sub

Private Sub CboPCode_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 13 Then
    Call uf_ValidateComboData(CboPCode, "4003", LblErrMsg, lbldesc)
    Call up_SettingGrid
End If
End Sub

Private Sub Combobox1_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)

If KeyCode = 13 Then
    Call uf_ValidateComboData(ComboBox1, "0024", LblErrMsg, Text1)
End If
End Sub

Private Sub cmdCancel_Click()
For i = 1 To grid.Rows - 1
    grid.TextMatrix(i, ColS) = ""
Next
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
ls_sql = " select a.*,b.description from PackingItem_Master a " & _
            "  left join PackingStyle_Cls b on a.PackingStyle_cls=b.PackingStyle_cls " & _
             " where a.item_code='" & Trim(CboPCode) & "' order by a.packingstyle_cls"
            
If rsisi.State <> adStateClosed Then rsisi.Close
rsisi.Open ls_sql, Db, adOpenKeyset, adLockOptimistic
If Not rsisi.EOF And Not rsisi.BOF Then
    rsisi.MoveFirst
    i = 1
    While Not rsisi.EOF
        grid.AddItem ""
        grid.TextMatrix(i, ColS) = ""
        
        grid.TextMatrix(i, ColPackingStyle) = Trim(rsisi!PackingStyle_Cls)
        grid.TextMatrix(i, ColPackingStyleDesc) = Trim(rsisi!Description)
        grid.TextMatrix(i, ColGrossWeight) = Format(rsisi!GrossWeight, gs_formatWeight)
        grid.TextMatrix(i, ColNetWeight) = Format(rsisi!NetWeight, gs_formatWeight)
        grid.TextMatrix(i, ColWidth) = Format(rsisi!Width, gs_formatWidth)
        grid.TextMatrix(i, ColLength) = Format(rsisi!Length, gs_formatLength)
        grid.TextMatrix(i, ColThickness) = Format(rsisi!Thickness, gs_formatThickness)
        grid.TextMatrix(i, ColCtn) = Format(rsisi!number_entering, gs_formatQty)
        grid.TextMatrix(i, ColVol) = Format((rsisi!Width / 1000) * (rsisi!Length / 1000) * (rsisi!Thickness / 1000), gs_formatVolume)

        grid.Cell(flexcpAlignment, 1, 0, i, ColS) = flexAlignCenterCenter
        grid.Cell(flexcpBackColor, grid.Rows - 1, ColS) = vbWhite
        rsisi.MoveNext
        i = i + 1
    Wend
End If
If rsisi.State <> adStateClosed Then rsisi.Close
End Sub

Private Sub CmdSubmit_Click()
Me.MousePointer = vbHourglass
Dim i As Long
Dim Status As String
Dim ls_sql As String
Status = "insert"
For i = 1 To grid.Rows - 1
    If grid.TextMatrix(i, ColS) = "D" Then
        Status = "delete"
        Exit For
    End If
    If grid.TextMatrix(i, ColS) = "S" Then
        Status = "update"
        Exit For
    End If
Next

If Status = "insert" Then
    If uf_Validate = False Then
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    Dim RS As New ADODB.Recordset
    If RS.State <> adStateClosed Then RS.Close
    RS.Open " Select * from [PackingItem_Master] " & _
                    "  WHERE [Item_Code] = '" & Trim(CboPCode) & "'  and  " & _
                    " [PackingStyle_Cls] = '" & Trim(ComboBox1) & "' ", Db, adOpenKeyset, adLockOptimistic
    If RS.EOF = False Then
        LblErrMsg = DisplayMsg(1023)
        Me.MousePointer = vbDefault
        If RS.State <> adStateClosed Then RS.Close
        Exit Sub
    End If
    If RS.State <> adStateClosed Then RS.Close
    ls_sql = " INSERT INTO [PackingItem_Master] " & _
                  "            ([Item_Code] " & _
                  "            ,[packingStyle_cls] " & _
                  "            ,[GrossWeight] " & _
                  "            ,[NetWeight] " & _
                  "            ,[Length] " & _
                  "            ,[Width] " & _
                  "            ,[Thickness] " & _
                  "            ,[Number_Entering]) " & _
                  "      VALUES " & _
                  "  ( "
    
    ls_sql = ls_sql + " '" & Trim(CboPCode) & "' " & _
                ", '" & Trim(ComboBox1) & "' " & _
                ", " & CDbl(Text2) & " " & _
                ", " & CDbl(Text3) & " " & _
                ", " & CDbl(Text4) & " " & _
                ", " & CDbl(Text5) & " " & _
                ", " & CDbl(Text6) & " " & _
                ", " & CDbl(Text7) & " )"
                
    Db.Execute ls_sql
    Call up_ClearInputArea
    LblErrMsg = DisplayMsg(1000)
ElseIf Status = "update" Then
    If uf_Validate = False Then
        Me.MousePointer = vbDefault
        Exit Sub
    End If
'    Dim Cost As Double
'    If Trim(ComboBox2) <> "" Then
'        Cost = 0
'    Else
'        Cost = CDbl(Text4)
'    End If
    ls_sql = " UPDATE [PackingItem_Master] " & _
                  "    SET  " & _
                  "       [GrossWeight] =  " & CDbl(Text2) & "    " & _
                  "       ,[NetWeight] =  " & CDbl(Text3) & "    " & _
                  "       ,[Length] =  " & CDbl(Text4) & "    " & _
                  "       ,[Width] =  " & CDbl(Text5) & "    " & _
                  "       ,[Thickness] =  " & CDbl(Text6) & "    " & _
                  "       ,[Number_Entering] =  " & CDbl(Text7) & "    " & _
                  "  WHERE [Item_Code] = '" & Trim(CboPCode) & "'  and  [PackingStyle_Cls] = '" & Trim(ComboBox1) & "' "
    Db.Execute ls_sql
    Call up_ClearInputArea
    LblErrMsg = DisplayMsg(1101)
ElseIf Status = "delete" Then
    If MsgBox(DisplayMsg(8071), vbYesNo, "Confirmation") = vbYes Then '#Are you sure want to delete data ?
            For i = 1 To grid.Rows - 1
                If grid.TextMatrix(i, ColS) = "D" Then
                        ls_sql = " delete from [PackingItem_Master] " & _
                                      "  WHERE [Item_Code] = '" & Trim(CboPCode) & "'  and  [PackingStyle_Cls] = '" & Trim(grid.TextMatrix(i, ColPackingStyle)) & "' "
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
        
        ComboBox1.Text = grid.TextMatrix(Row, ColPackingStyle)
        Text1.Text = grid.TextMatrix(Row, ColPackingStyleDesc)
        
        Text2.Text = grid.TextMatrix(Row, ColGrossWeight)
        Text3.Text = grid.TextMatrix(Row, ColNetWeight)
        Text4.Text = grid.TextMatrix(Row, ColLength)
        Text5.Text = grid.TextMatrix(Row, ColWidth)
        Text6.Text = grid.TextMatrix(Row, ColThickness)
        Text7.Text = grid.TextMatrix(Row, ColCtn)
        Text8.Text = grid.TextMatrix(Row, ColVol)
        
        ComboBox1.Enabled = False
        
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

Private Sub Text3_KeyPress(KeyAscii As Integer)
If InStr(1, "0123465798.", Chr(KeyAscii)) < 1 And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If InStr(1, "0123465798.", Chr(KeyAscii)) < 1 And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub

Private Sub Text4_Change()
    If IsNumeric(Text4) And IsNumeric(Text5) And IsNumeric(Text6) Then
        Text8 = Format((CDbl(Text4) / 1000) * (CDbl(Text5) / 1000) * (CDbl(Text6) / 1000), gs_formatVolume)
    End If
End Sub

Private Sub Text5_Change()
    If IsNumeric(Text4) And IsNumeric(Text5) And IsNumeric(Text6) Then
        Text8 = Format((CDbl(Text4) / 1000) * (CDbl(Text5) / 1000) * (CDbl(Text6) / 1000), gs_formatVolume)
    End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If InStr(1, "0123465798.", Chr(KeyAscii)) < 1 And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If InStr(1, "0123465798.", Chr(KeyAscii)) < 1 And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub

Private Sub Text6_Change()
    If IsNumeric(Text4) And IsNumeric(Text5) And IsNumeric(Text6) Then
        Text8 = Format((CDbl(Text4) / 1000) * (CDbl(Text5) / 1000) * (CDbl(Text6) / 1000), gs_formatVolume)
    End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If InStr(1, "0123465798.", Chr(KeyAscii)) < 1 And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If InStr(1, "0123465798.", Chr(KeyAscii)) < 1 And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub
