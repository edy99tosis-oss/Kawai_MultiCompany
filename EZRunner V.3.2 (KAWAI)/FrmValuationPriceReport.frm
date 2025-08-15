VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmValuationPriceReport 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Valuation Price Report"
   ClientHeight    =   10980
   ClientLeft      =   225
   ClientTop       =   525
   ClientWidth     =   15120
   Icon            =   "FrmValuationPriceReport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10980
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdBrowser 
      Caption         =   "..."
      Height          =   300
      Left            =   4815
      TabIndex        =   28
      Top             =   1702
      Width           =   300
   End
   Begin VB.CommandButton CmdPreview 
      BackColor       =   &H0080FFFF&
      Caption         =   "E&xcel"
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
      Left            =   12660
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9975
      Width           =   1035
   End
   Begin VB.CommandButton CmdSearch 
      BackColor       =   &H0080FFFF&
      Caption         =   "Sea&rch"
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
      Left            =   3900
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2130
      Width           =   1035
   End
   Begin VB.CommandButton CmdSubMenu 
      BackColor       =   &H00A6D2FF&
      Caption         =   "&Sub &Menu"
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
      Left            =   405
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9945
      Width           =   1125
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   405
      TabIndex        =   10
      Top             =   9225
      Width           =   14385
      Begin VB.Label LblErrMsg 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "LblErrMsg"
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
         Left            =   135
         TabIndex        =   11
         Top             =   225
         Width           =   14055
      End
   End
   Begin VB.CommandButton Cmd_Save 
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
      Left            =   13770
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9975
      Width           =   1035
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
      Left            =   11550
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9975
      Width           =   1035
   End
   Begin MSComCtl2.DTPicker DMonth 
      Height          =   315
      Left            =   2310
      TabIndex        =   4
      Top             =   2145
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   556
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
      Format          =   141230083
      UpDown          =   -1  'True
      CurrentDate     =   37798
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   5685
      Left            =   435
      TabIndex        =   12
      Top             =   3165
      Width           =   14355
      _cx             =   25321
      _cy             =   10028
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
      FocusRect       =   5
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
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
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   12930
      TabIndex        =   25
      Top             =   330
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   714
   End
   Begin VB.Label LblRecord 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Record(s)"
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
      Left            =   12720
      TabIndex        =   27
      Top             =   8940
      Width           =   1065
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Record(s)"
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
      Left            =   13740
      TabIndex        =   26
      Top             =   8940
      Width           =   1065
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Base Currency : "
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
      Index           =   3
      Left            =   435
      TabIndex        =   24
      Top             =   2790
      Width           =   1575
   End
   Begin VB.Label lblBase 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
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
      Left            =   2130
      TabIndex        =   23
      Top             =   2790
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Filter"
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
      Left            =   9225
      TabIndex        =   22
      Top             =   1770
      Width           =   1335
   End
   Begin MSForms.ComboBox CboFilter 
      Height          =   315
      Left            =   10815
      TabIndex        =   3
      Top             =   1710
      Width           =   3255
      VariousPropertyBits=   746604571
      MaxLength       =   6
      DisplayStyle    =   3
      Size            =   "5741;556"
      ListRows        =   15
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Line Line8 
      Index           =   1
      X1              =   12315
      X2              =   14130
      Y1              =   1590
      Y2              =   1590
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Index           =   2
      Left            =   435
      TabIndex        =   21
      Top             =   1740
      Width           =   1155
   End
   Begin VB.Label LblProductFilter 
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
      Left            =   5205
      TabIndex        =   20
      Top             =   1740
      Width           =   3375
   End
   Begin MSForms.ComboBox CboProductFilter 
      Height          =   315
      Left            =   2310
      TabIndex        =   2
      Top             =   1695
      Width           =   2460
      VariousPropertyBits=   746604571
      MaxLength       =   25
      DisplayStyle    =   3
      Size            =   "4339;556"
      ListRows        =   15
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Line Line1 
      X1              =   5205
      X2              =   8585
      Y1              =   1980
      Y2              =   1980
   End
   Begin VB.Label Lbl_Make 
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
      Height          =   225
      Left            =   12345
      TabIndex        =   19
      Top             =   1335
      Width           =   1305
   End
   Begin VB.Line Line6 
      X1              =   5220
      X2              =   7950
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label lbl_finish 
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
      Height          =   225
      Left            =   5220
      TabIndex        =   18
      Top             =   1305
      Width           =   1845
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Make Buy Cls"
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
      Left            =   9225
      TabIndex        =   17
      Top             =   1320
      Width           =   1335
   End
   Begin MSForms.ComboBox CboMake 
      Height          =   315
      Left            =   10830
      TabIndex        =   1
      Top             =   1290
      Width           =   1410
      VariousPropertyBits=   612386843
      MaxLength       =   15
      DisplayStyle    =   3
      Size            =   "2487;556"
      ListRows        =   15
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label LblPesan 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "rere"
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
      Left            =   540
      TabIndex        =   16
      Top             =   9540
      Width           =   11940
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Valuation Price Report"
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
      Left            =   495
      TabIndex        =   15
      Top             =   315
      Width           =   14310
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Month Period"
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
      Left            =   435
      TabIndex        =   14
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Finish Good Part Cls"
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
      Left            =   450
      TabIndex        =   13
      Top             =   1290
      Width           =   2565
   End
   Begin MSForms.ComboBox CboFinish 
      Height          =   315
      Left            =   2325
      TabIndex        =   0
      Top             =   1260
      Width           =   1785
      VariousPropertyBits=   612386843
      MaxLength       =   15
      DisplayStyle    =   3
      Size            =   "3149;556"
      ListRows        =   15
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "FrmValuationPriceReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dateUp As Date
Dim ColProd, ColDesc As Byte
Dim ColPremStock, ColPremPrice, ColPremMount, ColIncomeStock, ColIncomePrice, ColIncomeMount, _
     ColIncomeOtherStock, ColIncomeOtherPrice, ColIncomeOtherMount, ColOutgoingStock, ColOutgoingPrice, ColOutgoingMount, _
     ColOutgoingOtherStock, ColOutgoingOtherPrice, ColOutgoingOtherMount As Byte
Dim ColCurrent, ColCurrentPrice, ColCurrentPriceTotal, ColCurrentAmount As Byte
Dim ColInventory, ColInventoryPrice, ColInventoryAmount, ColUnit, ColGroup As Byte

Dim Harga As Integer
Dim tampung() As Integer
Dim Num As Byte


Private Sub SETING()
CboFinish.clear
CboFinish.columnCount = 2

CboFinish.AddItem
CboFinish.List(0, 0) = strAll
CboFinish.List(0, 1) = strAll
CboFinish.AddItem
CboFinish.List(1, 0) = "01"
CboFinish.List(1, 1) = "Finish Goods"
CboFinish.AddItem
CboFinish.List(2, 0) = "02"
CboFinish.List(2, 1) = "Parts/WIP/Material"

CboFinish.ListWidth = 110
CboFinish.ColumnWidths = "20 pt ; 90 pt "
CboFinish.ListIndex = 0
CboFinish.Text = CboFinish.List(0, 0)
'==========================
CboMake.clear
CboMake.columnCount = 2

CboMake.AddItem
CboMake.List(0, 0) = strAll
CboMake.List(0, 1) = strAll
CboMake.AddItem
CboMake.List(1, 0) = "01"
CboMake.List(1, 1) = "Make"
CboMake.AddItem
CboMake.List(2, 0) = "02"
CboMake.List(2, 1) = "Buy"

CboMake.ListWidth = 70
CboMake.ColumnWidths = "20 pt ; 50 pt "
CboMake.ListIndex = 0
CboMake.Text = CboMake.List(0, 0)

'=============FILTER=============
CboFilter.clear
CboFilter.columnCount = 1

CboFilter.AddItem
CboFilter.List(0, 1) = "01"
CboFilter.List(0, 0) = "Show all data"
CboFilter.AddItem
CboFilter.List(1, 1) = "02"
CboFilter.List(1, 0) = "Show negatif stock qty only"
CboFilter.AddItem
CboFilter.List(2, 1) = "03"
CboFilter.List(2, 0) = "Show positive stock qty only"
CboFilter.AddItem
CboFilter.List(3, 1) = "04"
CboFilter.List(3, 0) = "Show zero stock qty only"
CboFilter.ListWidth = 150
CboFilter.ColumnWidths = "20 pt ; 130 pt "
CboFilter.ListIndex = 0
CboFilter.Text = CboFilter.List(0, 0)
 
End Sub

Private Sub CboFilter_Change()
    Call Header
End Sub

Private Sub CboFilter_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
KeyCode = 0
End Sub

Private Sub CboFilter_KeyPress(KeyAscii As MSForms.ReturnInteger)
KeyAscii = 0
End Sub

Private Sub CboFinish_Click()
lbl_finish.Caption = CboFinish.List(CboFinish.ListIndex, 1)
LblErrMsg.Caption = ""
Call Product(CboFinish, CboMake)
Call Header
End Sub

Private Sub CboFinish_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = vbKeyDelete Then lbl_finish.Caption = ""
If KeyCode = vbKeyBack Then lbl_finish.Caption = ""  'KeyCode = 0
If KeyCode = 13 Then
    If CboFinish.ListCount < 1 Then Exit Sub
    For i = 0 To CboFinish.ListCount - 1
        If Trim(CboFinish.Text) = CboFinish.List(i, 0) Then
            lbl_finish.Caption = CboFinish.List(i, 1): LblErrMsg.Caption = "": Exit Sub
        End If
    Next
    LblErrMsg.Caption = DisplayMsg(8096) '"Invalid Make/Buy Clasification !"
    CboFinish.SetFocus
    lbl_finish.Caption = ""
End If
End Sub

Private Sub CboMake_Click()
Lbl_Make.Caption = CboMake.List(CboMake.ListIndex, 1)
LblErrMsg.Caption = ""
Call Product(CboFinish, CboMake)
Call Header
End Sub

Private Sub CboMake_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = vbKeyDelete Then Lbl_Make.Caption = ""
If KeyCode = vbKeyBack Then Lbl_Make.Caption = ""  'KeyCode = 0
If KeyCode = 13 Then
    If CboMake.ListCount < 1 Then Exit Sub
    For i = 0 To CboMake.ListCount - 1
        If Trim(CboMake.Text) = CboMake.List(i, 0) Then
            Lbl_Make.Caption = CboMake.List(i, 1): LblErrMsg.Caption = "": Exit Sub
        End If
    Next
    LblErrMsg.Caption = DisplayMsg(8096)  '"Invalid Make/Buy Clasification !"
    CboMake.SetFocus
    Lbl_Make.Caption = ""
End If
End Sub

Private Sub Cmd_Save_Click()
Dim sql As String
For i = 2 To grid.Rows - 1
        sql = "update inventory_price set inventory_price='" & CDbl(grid.TextMatrix(i, ColInventoryPrice)) & "' " & _
                " where inventory_year='" & DMonth.Year & "' " & _
                " and inventory_month='" & DMonth.Month & "'  " & _
                " and item_code='" & grid.TextMatrix(i, ColProd) & "'"
        Db.Execute sql
Next i
Browse
LblErrMsg = DisplayMsg(1101)
End Sub

Private Sub cmdBrowser_Click()
 Me.MousePointer = vbHourglass
 frm_BrowseItem.getItemCode = CboProductFilter.Text
 frm_BrowseItem.Show 1
 CboProductFilter.Text = frm_BrowseItem.getItemCode
 Me.MousePointer = vbDefault
End Sub

Private Sub cmdCancel_Click()
DMonth = Now()
SETING
grid.clear
Header
LblErrMsg = ""
End Sub

Private Sub cmdSearch_Click()
 
LblErrMsg = ""
Me.MousePointer = vbHourglass
Call up_DropSQLFunctionALL
Call up_CreateSQLFunctionALL
Header
Browse
Call up_DropSQLFunctionALL
Me.MousePointer = vbDefault
End Sub

Private Sub CmdSubMenu_Click()
    frmMainMenu.Show
    Unload Me
End Sub

Private Sub DMonth_Change()
If Format(DMonth.Value, "MM") < Format(dateUp, "MM") And Val(Format(DMonth.Value, "MM")) = 1 And Val(Format(dateUp, "MM")) = 12 Then _
            DMonth.Year = DMonth.Year + 1: GoTo pass
    If Format(DMonth.Value, "MM") > Format(dateUp, "MM") And Val(Format(DMonth.Value, "MM")) = 12 And Val(Format(dateUp, "MM")) = 1 Then _
            DMonth.Year = DMonth.Year - 1
pass:
    dateUp = Format(DMonth.Value, "dd MMM yyyy")
Call Header
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
Dim Ret As String, NC As Long, TempPWD As String
Dim RS As New Recordset
    
CtrlMenu1.FormName = Me.Name
Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
Num = 0
SETING
DMonth = Format(Now(), "MMM yyyy")
dateUp = DMonth.Value
RS.Open "select valuationprice_BaseCurrency from Company_Profile", Db, adOpenForwardOnly, adLockReadOnly
lblBase.Caption = uf_GetCurrencyDescription(Trim$(RS(0) & ""))
RS.Close
End Sub
Sub SettingColumn(lbt_AddCol As Byte)
ColProd = 0
ColDesc = 1
ColUnit = 2
ColGroup = 3
ColPremStock = 4
ColPremPrice = 5
ColPremMount = 6
ColIncomeStock = 7
ColIncomePrice = 8
ColIncomeMount = 9
ColIncomeOtherStock = 10
ColIncomeOtherPrice = 11
ColIncomeOtherMount = 12
ColOutgoingStock = 13
ColOutgoingPrice = 14
ColOutgoingMount = 15
ColOutgoingOtherStock = 16
ColOutgoingOtherPrice = 17
ColOutgoingOtherMount = 18
ColCurrent = 19
ColCurrentPrice = 20
ColCurrentPriceTotal = 21 + lbt_AddCol
ColCurrentAmount = 22 + lbt_AddCol
ColInventory = 23 + lbt_AddCol
ColInventoryPrice = 24 + lbt_AddCol
ColInventoryAmount = 25 + lbt_AddCol


End Sub

Sub Header()
LblRecord = "0"
Dim RS As New ADODB.Recordset
Dim i, z As Integer
sql = "select * from inventorycost_master where additional_cls='0' order by cost_cls"
If RS.State <> adStateClosed Then RS.Close
RS.CursorLocation = adUseClient
RS.Open sql, Db, adOpenKeyset, adLockOptimistic
If Not RS.EOF Then
    Num = RS.RecordCount
End If

    With grid
        .Rows = 2
        
        '#Num Di Set Ke Nol (Header, Browse,Excel) jika tidak ingin menampilkan Additional Cost
        Num = 0
        
        .ColS = 26 + Num
        .FixedRows = 2

        Call SettingColumn(Num)
        
        '=================================================================
        'Set Additional Cost  Header
        
        If Num > 0 Then
            i = 1
            While RS.EOF = False
                .TextMatrix(0, ColCurrentPrice + i) = "Current"
                .TextMatrix(1, ColCurrentPrice + i) = Trim(RS!cost_cls) & "/" & Trim(RS!cost_title)
                i = i + 1
                RS.MoveNext
            Wend
        End If
        If RS.State <> adStateClosed Then RS.Close
        '=================================================================
        
        .TextMatrix(0, ColProd) = "Product Code"
        .TextMatrix(1, ColProd) = "Product Code"
        .TextMatrix(0, ColDesc) = "Description"
        .TextMatrix(1, ColDesc) = "Description"
        .TextMatrix(0, ColUnit) = "Unit"
        .TextMatrix(1, ColUnit) = "Unit"
        .TextMatrix(0, ColGroup) = "Group"
        .TextMatrix(1, ColGroup) = "Group"
        .TextMatrix(0, ColPremStock) = "Premonth"
        .TextMatrix(0, ColPremPrice) = "Premonth"
        .TextMatrix(0, ColPremMount) = "Premonth"
        .TextMatrix(0, ColIncomeStock) = "Incoming"
        .TextMatrix(0, ColIncomePrice) = "Incoming"
        .TextMatrix(0, ColIncomeMount) = "Incoming"
        .TextMatrix(0, ColIncomeOtherStock) = "Incoming Other"
        .TextMatrix(0, ColIncomeOtherPrice) = "Incoming Other"
        .TextMatrix(0, ColIncomeOtherMount) = "Incoming Other"
        .TextMatrix(0, ColOutgoingStock) = "Outgoing"
        .TextMatrix(0, ColOutgoingPrice) = "Outgoing"
        .TextMatrix(0, ColOutgoingMount) = "Outgoing"
        .TextMatrix(0, ColOutgoingOtherStock) = "Outgoing Other"
        .TextMatrix(0, ColOutgoingOtherPrice) = "Outgoing Other"
        .TextMatrix(0, ColOutgoingOtherMount) = "Outgoing Other"
        .TextMatrix(0, ColCurrent) = "Current"
        .TextMatrix(0, ColCurrentPrice) = "Current"
        .TextMatrix(0, ColCurrentPriceTotal) = "Current"
        .TextMatrix(0, ColCurrentAmount) = "Current"
        .TextMatrix(0, ColInventory) = "Inventory"
        .TextMatrix(0, ColInventoryPrice) = "Inventory"
        .TextMatrix(0, ColInventoryAmount) = "Inventory"
        
        .TextMatrix(1, ColPremStock) = "Stock"
        .TextMatrix(1, ColPremPrice) = "Price"
        .TextMatrix(1, ColPremMount) = "Amount"
        .TextMatrix(1, ColIncomeStock) = "Stock"
        .TextMatrix(1, ColIncomePrice) = "Price"
        .TextMatrix(1, ColIncomeMount) = "Amount"
        .TextMatrix(1, ColIncomeOtherStock) = "Stock"
        .TextMatrix(1, ColIncomeOtherPrice) = "Price"
        .TextMatrix(1, ColIncomeOtherMount) = "Amount"
        .TextMatrix(1, ColOutgoingStock) = "Stock"
        .TextMatrix(1, ColOutgoingPrice) = "Price"
        .TextMatrix(1, ColOutgoingMount) = "Amount"
        .TextMatrix(1, ColOutgoingOtherStock) = "Stock"
        .TextMatrix(1, ColOutgoingOtherPrice) = "Price"
        .TextMatrix(1, ColOutgoingOtherMount) = "Amount"
        .TextMatrix(1, ColCurrent) = "Stock"
        .TextMatrix(1, ColCurrentPrice) = "Price"
        .TextMatrix(1, ColCurrentPriceTotal) = "Total Price"
        .TextMatrix(1, ColCurrentAmount) = "Amount"
        .TextMatrix(1, ColInventory) = "Stock"
        .TextMatrix(1, ColInventoryPrice) = "Price"
        .TextMatrix(1, ColInventoryAmount) = "Amount"
    
    
    
        .MergeRow(0) = True
        .MergeCol(ColProd) = True
        .MergeCol(ColDesc) = True
        .MergeCol(ColUnit) = True
        .MergeCol(ColGroup) = True
        .MergeCells = flexMergeFixedOnly
        .ColHidden(ColCurrentPrice) = True

        
        .Cell(flexcpAlignment, 0, ColPremStock, 1, ColInventoryAmount) = flexAlignCenterCenter
        

        .ColWidth(ColProd) = 2750
        .ColWidth(ColDesc) = 3600
        .ColWidth(ColUnit) = 500
        .ColWidth(ColGroup) = 1400
        For i = ColPremStock To ColInventoryAmount
            .ColWidth(i) = 1200
        Next
        .FrozenCols = 2
        
    End With
If RS.State <> adStateClosed Then RS.Close
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Col = ColInventoryPrice Then
    If IsNumeric(grid.TextMatrix(Row, ColInventoryPrice)) = False Then
        grid.TextMatrix(Row, ColInventoryPrice) = 0
    End If
    If CDbl(grid.TextMatrix(Row, ColInventoryPrice)) > gd_MaxPrice Then
        grid.TextMatrix(Row, ColInventoryPrice) = gd_MaxPrice
        LblErrMsg = DisplayMsg(4048) & " " & gd_MaxPrice
    Else
        LblErrMsg = ""
    End If
    grid.TextMatrix(Row, ColInventoryPrice) = Format(grid.TextMatrix(Row, ColInventoryPrice), gs_formatPriceIDR)
    grid.TextMatrix(Row, ColInventoryAmount) = Format(CDbl(grid.TextMatrix(Row, ColInventory)) * CDbl(grid.TextMatrix(Row, ColInventoryPrice)), gs_formatAmountIDR)
End If
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> ColInventoryPrice Then
        Cancel = True
    End If
End Sub

Private Sub grid_KeyPress(KeyAscii As Integer)
    If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 46 And KeyAscii <> 8 Then KeyAscii = 0
End Sub


Sub Browse()
Dim RS As Recordset
Dim rsCek As Recordset
Dim NiL As Integer, z As Integer, k As Integer
ReDim tampung(1 To 20)
LblRecord = Format(0, "#,##0")
If CboProductFilter = "" Then
    LblErrMsg = DisplayMsg(1009)  ' "Please input Product item"
    Exit Sub
End If

sql = " select  IP.item_code, " & _
      " IM.item_name, UC.Description UnitDesc,GC.Description GroupDesc," & _
      " premonth_stock, " & _
      " premonth_price, " & _
      " (premonth_stock*premonth_price)Premonth_amount,       " & _
      "  incoming_stock, " & _
      "  incoming_price,       " & _
      " (incoming_stock*incoming_price) Incoming_Amount," & _
      " incomingother_stock, " & _
      " incomingother_price, " & _
      " (incomingother_stock*incomingother_price) incomingother_amount,       " & _
      " outgoing_stock, " & _
      " outgoing_price, " & _
      " (outgoing_stock*outgoing_price)Outgoing_Amount,             " & _
      "  outgoingother_stock, " & _
      " outgoingother_price, " & _
      " (outgoingother_stock*outgoingother_price)OutgoingOther_amount,       " & _
      " Current_stock," & _
      " Current_Price,   " & _
      " inventory_price       " & _
      " from inventory_price IP " & _
      "  left join item_master IM on IP.item_code=IM.item_code         " & _
      "  left join unit_cls uc on IM.Unit_Cls=uc.Unit_Cls       " & _
      "  left join group_cls gc on IM.group_cls=gc.Group_Cls       "
      
sql = sql + " where IP.inventory_year='" & Format(DMonth, "yyyy") & "'  " & _
      " and IP.inventory_month='" & Format(DMonth, "MM") & "'"
      
If CboFinish <> strAll Then
    sql = sql + " and  IM.finishgoodpart_cls='" & Trim(CboFinish) & "'"
End If

If CboMake <> strAll Then
    sql = sql + "and IM.makebuy_cls='" & Trim(CboMake) & "'"
End If

If CboProductFilter <> strAll Then
    sql = sql + "and IP.item_code='" & Trim(CboProductFilter) & "'"
End If

If CboFilter.List(CboFilter.ListIndex, 1) = "01" Then

ElseIf CboFilter.List(CboFilter.ListIndex, 1) = "04" Then
        sql = sql + " and IP.Current_Stock=0"
ElseIf CboFilter.List(CboFilter.ListIndex, 1) = "02" Then
        sql = sql + " and IP.current_stock < 0 "
ElseIf CboFilter.List(CboFilter.ListIndex, 1) = "03" Then
        sql = sql + " and IP.current_stock > 0 "
End If
Set RS = New ADODB.Recordset
RS.Open sql, Db, adOpenKeyset, adLockOptimistic

i = 1
Dim j As Integer
Call Header
Dim ls_date As String
Dim ld_AdditionalCost As Double
ls_date = DateAdd("d", 1, Format(DateAdd("m", 1, DMonth), "yyyy-MM-01"))
While Not RS.EOF
        i = i + 1
        grid.AddItem ""
        grid.TextMatrix(i, ColProd) = Trim$(RS!Item_Code)
        grid.TextMatrix(i, ColDesc) = Trim$(RS!item_name)
        grid.TextMatrix(i, ColUnit) = Trim$(RS!unitdesc)
        grid.TextMatrix(i, ColGroup) = Trim$(RS!groupdesc & "")
        grid.TextMatrix(i, ColPremStock) = Format(IIf(IsNull(RS("premonth_stock")), "0", Trim(RS("premonth_stock"))), gs_formatQty)
        grid.TextMatrix(i, ColPremPrice) = Format(IIf(IsNull(RS("premonth_price")), "0", Trim(RS("premonth_price"))), gs_formatPriceIDR)
        grid.TextMatrix(i, ColPremMount) = Format(IIf(IsNull(RS("premonth_amount")), "0", Trim(RS("premonth_amount"))), gs_formatAmountIDR)
        grid.TextMatrix(i, ColIncomeStock) = Format(IIf(IsNull(RS("incoming_stock")), "0", Trim(RS("incoming_stock"))), gs_formatQty)
        grid.TextMatrix(i, ColIncomePrice) = Format(IIf(IsNull(RS("incoming_price")), "0", Trim(RS("incoming_price"))), gs_formatPriceIDR)
        grid.TextMatrix(i, ColIncomeMount) = Format(IIf(IsNull(RS("incoming_amount")), "0", Trim(RS("incoming_amount"))), gs_formatAmountIDR)
        grid.TextMatrix(i, ColIncomeOtherStock) = Format(IIf(IsNull(RS("incomingother_stock")), "0", Trim(RS("incomingother_stock"))), gs_formatQty)
        grid.TextMatrix(i, ColIncomeOtherPrice) = Format(IIf(IsNull(RS("incomingother_price")), "0", Trim(RS("incomingother_price"))), gs_formatPriceIDR)
        grid.TextMatrix(i, ColIncomeOtherMount) = Format(IIf(IsNull(RS("incomingother_amount")), "0", Trim(RS("incomingother_amount"))), gs_formatAmountIDR)
        grid.TextMatrix(i, ColOutgoingStock) = Format(IIf(IsNull(RS("outgoing_stock")), "0", Trim(RS("outgoing_stock"))), gs_formatQty)
        grid.TextMatrix(i, ColOutgoingPrice) = Format(IIf(IsNull(RS("outgoing_price")), "0", Trim(RS("outgoing_price"))), gs_formatPriceIDR)
        grid.TextMatrix(i, ColOutgoingMount) = Format(IIf(IsNull(RS("outgoing_amount")), "0", Trim(RS("outgoing_amount"))), gs_formatAmountIDR)
        grid.TextMatrix(i, ColOutgoingOtherStock) = Format(IIf(IsNull(RS("outgoingother_stock")), "0", Trim(RS("outgoingother_stock"))), gs_formatQty)
        grid.TextMatrix(i, ColOutgoingOtherPrice) = Format(IIf(IsNull(RS("outgoingother_price")), "0", Trim(RS("outgoingother_price"))), gs_formatPriceIDR)
        grid.TextMatrix(i, ColOutgoingOtherMount) = Format(IIf(IsNull(RS("outgoingother_amount")), "0", Trim(RS("outgoingother_amount"))), gs_formatAmountIDR)
        grid.TextMatrix(i, ColCurrent) = Format(IIf(IsNull(RS("current_stock")), "0", Trim(RS("current_stock"))), gs_formatQty)
        'Grid.TextMatrix(i, ColCurrentPrice) = Format(IIf(IsNull(rs("current_price")), "0", Trim(rs("current_price"))), gs_formatPriceIDR)
        
        ld_AdditionalCost = 0
        
        '#Num Di Set Ke Nol (Header, Browse,Excel) jika tidak ingin menampilkan Additional Cost
        Num = 0
        
        For j = 1 To Num
            grid.TextMatrix(i, ColCurrentPrice + j) = Format(uf_GetAdditionalCost(Left(grid.TextMatrix(1, ColCurrentPrice + j), 2), Trim$(RS!Item_Code), ls_date), gs_formatPriceIDR)
            ld_AdditionalCost = ld_AdditionalCost + grid.TextMatrix(i, ColCurrentPrice + j)
       Next
       'Grid.TextMatrix(i, ColCurrentPriceTotal) = Format(Grid.TextMatrix(i, ColCurrentPrice) + ld_AdditionalCost, gs_formatPriceIDR)
       grid.TextMatrix(i, ColCurrentPriceTotal) = Format(IIf(IsNull(RS("current_price")), "0", Trim(RS("current_price"))), gs_formatPriceIDR)
       grid.TextMatrix(i, ColCurrentAmount) = Format(grid.TextMatrix(i, ColCurrent) * grid.TextMatrix(i, ColCurrentPriceTotal), gs_formatAmountIDR)
       grid.TextMatrix(i, ColInventory) = Format(IIf(IsNull(RS("current_stock")), "0", Trim(RS("current_stock"))), gs_formatQty)
       grid.TextMatrix(i, ColInventoryPrice) = Format(IIf(IsNull(RS("Inventory_price")), "0", Trim(RS("Inventory_price"))), gs_formatPriceIDR)
       grid.TextMatrix(i, ColInventoryAmount) = Format(grid.TextMatrix(i, ColCurrent) * grid.TextMatrix(i, ColInventoryPrice), gs_formatAmountIDR)
            
        grid.Cell(flexcpBackColor, i, ColInventoryPrice) = vbWhite
        
   RS.MoveNext
Wend
LblRecord = Format(RS.RecordCount, "#,##0")

End Sub

Private Function uf_GetAdditionalCost(ls_CostCls As String, ls_ItemCode As String, ls_date As String) As Double
    Dim ls_sql As String
    Dim RS As New ADODB.Recordset

    If RS.State <> adStateClosed Then RS.Close
    ls_sql = "  " & _
                  " declare @Date as char(10) " & _
                  " declare @CostCls as char(2) " & _
                  " declare @ItemCode as char(15) " & _
                  "  " & _
                  " Set @Date='" & Format(ls_date, "yyyy-MM-dd") & "' " & _
                  " Set @CostCls='" & ls_CostCls & "' " & _
                  " Set @ItemCode='" & ls_ItemCode & "' " & _
                  " select MC.Cost_Cls, " & _
                  " ( " & _
                  " select isnull(MI.Amount,0) * dbo.UF_GetBookExchangeRate(year(@Date),month(@Date),MI.Currency_Code) " & _
                  " from InventoryCost_Item MI " & _
                  " where MI.Item_Code=@ItemCode "

    ls_sql = ls_sql + " and Cost_Cls=MC.Cost_Cls " & _
                      " and Start_Date <=replace(@Date,'-','') " & _
                      " and End_Date>= replace(@Date,'-','') " & _
                      " ) CostItem, " & _
                      " isnull(( " & _
                      " select isnull(MI.Amount,0) * dbo.UF_GetBookExchangeRate(year(@Date),month(@Date),MI.Currency_Code) " & _
                      " from InventoryCost_Group MI " & _
                      " Where Cost_Cls=MC.Cost_Cls " & _
                      " and Start_Date <=replace(@Date,'-','') " & _
                      " and End_Date>= replace(@Date,'-','') " & _
                      " ),0) CostGroup "
    
    ls_sql = ls_sql + " from inventorycost_master MC " & _
                      " where  MC.additional_cls='0' " & _
                      " and MC.Cost_Cls=@CostCls "

    RS.Open ls_sql, Db, adOpenKeyset, adLockOptimistic
    
    
    If RS.EOF = False Then
        If IsNull(RS!CostItem) Then
            uf_GetAdditionalCost = RS!CostGroup
        Else
            uf_GetAdditionalCost = RS!CostItem
        End If
    Else
        uf_GetAdditionalCost = 0
    End If
    If RS.State <> adStateClosed Then RS.Close
End Function

Private Sub grid_Click()
With grid
If grid.Rows <> 1 And grid.Row <> -1 Then
    If grid.Cell(flexcpBackColor, grid.Row, grid.Col, grid.Row, grid.Col) <> vbWhite Then
        grid.FocusRect = flexFocusNone
    Else
        grid.FocusRect = flexFocusInset
    End If
End If
End With
End Sub


Private Sub CmdPreview_Click()

If MsgBox("System need to submit any changes you made " & vbCrLf & "on this screen before show report !" & vbCrLf & "Do you want to submit the data ?", vbYesNo + vbExclamation, "Confirmation") = vbNo Then
    Exit Sub
Else
    Call Cmd_Save_Click
End If
Dim xlapp As New Excel.application

Me.MousePointer = vbHourglass
If grid.Rows > 2 Then

        Dim Idx As Integer
        Dim xlColProd As String
        Dim xlColDesc As String
        Dim xlColUnit As String
        Dim xlColGroup As String
        Dim xlColPremStock As String
        Dim xlColPremPrice As String
        Dim xlColPremMount As String
        Dim xlColIncomeStock As String
        Dim xlColIncomePrice As String
        Dim xlColIncomeMount As String
        Dim xlColIncomeOtherStock As String
        Dim xlColIncomeOtherPrice As String
        Dim xlColIncomeOtherMount As String
        Dim xlColOutgoingStock As String
        Dim xlColOutgoingPrice As String
        Dim xlColOutgoingMount As String
        Dim xlColOutgoingOtherStock As String
        Dim xlColOutgoingOtherPrice As String
        Dim xlColOutgoingOtherMount As String
        Dim xlColCurrent As String
        Dim xlColCurrentPrice As String
        Dim xlColCurrentPriceTotal As String
        Dim xlColCurrentAmount As String
        Dim xlColInventory As String
        Dim xlColInventoryPrice As String
        Dim xlColInventoryAmount  As String
                        
         xlColProd = "a"
         xlColDesc = "b"
         xlColUnit = "c"
         xlColGroup = "d"
         xlColPremStock = "e"
         xlColPremPrice = "f"
         xlColPremMount = "g"
         xlColIncomeStock = "h"
         xlColIncomePrice = "i"
         xlColIncomeMount = "j"
         xlColIncomeOtherStock = "k"
         xlColIncomeOtherPrice = "l"
         xlColIncomeOtherMount = "m"
         xlColOutgoingStock = "n"
         xlColOutgoingPrice = "o"
         xlColOutgoingMount = "p"
         xlColOutgoingOtherStock = "q"
         xlColOutgoingOtherPrice = "r"
         xlColOutgoingOtherMount = "s"
         xlColCurrent = "t"
         
         xlColCurrentPrice = xlColCurrent
         
         Dim i As Long
         Dim ls_prevCol As String
         ls_prevCol = xlColCurrentPrice
         For i = 1 To Num
             ls_prevCol = uf_GetXLColumn(ls_prevCol)
         Next
         
            ls_prevCol = uf_GetXLColumn(ls_prevCol)
         xlColCurrentPriceTotal = ls_prevCol
            ls_prevCol = uf_GetXLColumn(ls_prevCol)
         xlColCurrentAmount = ls_prevCol
            ls_prevCol = uf_GetXLColumn(ls_prevCol)
         xlColInventory = ls_prevCol
            ls_prevCol = uf_GetXLColumn(ls_prevCol)
         xlColInventoryPrice = ls_prevCol
            ls_prevCol = uf_GetXLColumn(ls_prevCol)
         xlColInventoryAmount = ls_prevCol
        
        With xlapp
        
        .Workbooks.Add
        .Range(xlColProd & "2", xlColInventoryAmount & "2").Merge
        .Range(xlColProd & "2") = "Valuation Price Report"
        .Range(xlColProd & "2").horizontalAlignment = xlCenter
        .Range(xlColProd & "2").Font.Bold = True
        
        .Range(xlColProd & "3:" & xlColInventoryAmount & "3").Merge
        .Range(xlColProd & "3") = "Base Currency : " + lblBase
        .Range(xlColProd & "3").horizontalAlignment = xlCenter
        .Range(xlColProd & "3").Font.Bold = False
        
        .Range(xlColInventory & "7", xlColInventoryAmount & "7").Merge
        .Range(xlColInventory & "7") = "Issued Date : " + Format(Now, "dd MMM yyyy  hh:MM:ss")
        .Range(xlColInventory & "7").horizontalAlignment = xlRight
        
        .Range(xlColProd & "4") = "Finish Good Part CLS"
        .Range(xlColDesc & "4") = CboFinish
        .Range(xlColProd & "5") = "MakeBuy CLs"
        .Range(xlColDesc & "5") = CboMake
        .Range(xlColProd & "6") = "Product Code"
        .Range(xlColDesc & "6") = CboProductFilter
        .Range(xlColProd & "7") = "Period"
        .Range(xlColDesc & "7").horizontalAlignment = xlLeft
        .Range(xlColDesc & "7") = "'" & Format(DMonth, "mmmm yyyy")
        .Range(xlColProd & "9") = "Product Code"
        .Range(xlColProd & "9", xlColProd & "10").Merge
        
        .Range(xlColUnit & "9") = "Unit"
        .Range(xlColUnit & "9", xlColUnit & "10").Merge
        .Range(xlColGroup & "9") = "Group"
        .Range(xlColGroup & "9", xlColGroup & "10").Merge
        
        .Range(xlColDesc & "9") = "Description"
        .Range(xlColDesc & "9", xlColDesc & "10").Merge
        .Range(xlColPremStock & "9", xlColPremMount & "9").Merge
        .Range(xlColPremStock & "9") = "Premonth"
        .Range(xlColIncomeStock & "9", xlColIncomeMount & "9").Merge
        .Range(xlColIncomeStock & "9") = "Incoming"
        .Range(xlColIncomeOtherStock & "9", xlColIncomeOtherMount & "9").Merge
        .Range(xlColIncomeOtherStock & "9") = "Incoming Other"
        .Range(xlColOutgoingStock & "9", xlColOutgoingMount & "9").Merge
        .Range(xlColOutgoingStock & "9") = "Outgoing"
        .Range(xlColOutgoingOtherStock & "9", xlColOutgoingOtherMount & "9").Merge
        .Range(xlColOutgoingOtherStock & "9") = "Outgoing Other"
        .Range(xlColCurrent & "9", xlColCurrentAmount & "9").Merge
        .Range(xlColCurrent & "9") = "Current"
        .Range(xlColInventory & "9", xlColInventoryAmount & "9").Merge
        .Range(xlColInventory & "9") = "Inventory"
        .Range(xlColProd & "9:" & xlColInventoryAmount & "10").horizontalAlignment = xlCenter
        .Range(xlColProd & "9:" & xlColInventoryAmount & "10").verticalAlignment = xlCenter
        
        Idx = 10
        
        '#Grid Header
        .Range(xlColPremStock & Idx) = "Stock"
        .Range(xlColPremPrice & Idx) = "Price"
        .Range(xlColPremMount & Idx) = "Amount"
        .Range(xlColIncomeStock & Idx) = "Stock"
        .Range(xlColIncomePrice & Idx) = "Price"
        .Range(xlColIncomeMount & Idx) = "Amount"
        .Range(xlColIncomeOtherStock & Idx) = "Stock"
        .Range(xlColIncomeOtherPrice & Idx) = "Price"
        .Range(xlColIncomeOtherMount & Idx) = "Amount"
        .Range(xlColOutgoingStock & Idx) = "Stock"
        .Range(xlColOutgoingPrice & Idx) = "Price"
        .Range(xlColOutgoingMount & Idx) = "Amount"
        .Range(xlColOutgoingOtherStock & Idx) = "Stock"
        .Range(xlColOutgoingOtherPrice & Idx) = "Price"
        .Range(xlColOutgoingOtherMount & Idx) = "Amount"
        .Range(xlColCurrent & Idx) = "Stock"
        '.Range(xlColCurrentPrice & idx) = "Price"
        
        ls_prevCol = xlColCurrentPrice
        For i = ColCurrentPrice + 1 To ColCurrentPriceTotal - 1
            ls_prevCol = uf_GetXLColumn(ls_prevCol)
            .Range(ls_prevCol & Idx) = Trim(grid.TextMatrix(1, i))
        Next
        
        '.Range(xlColCurrentPriceTotal & idx) = "Price Total"
        .Range(xlColCurrentPriceTotal & Idx) = "Price"
        .Range(xlColCurrentAmount & Idx) = "Amount"
        .Range(xlColInventory & Idx) = "Stock"
        .Range(xlColInventoryPrice & Idx) = "Price"
        .Range(xlColInventoryAmount & Idx) = "Amount"
            
        '#Fill Grid
        Dim j As Integer
        For i = 2 To grid.Rows - 1
            Idx = Idx + 1
            .Range(xlColProd & Idx) = grid.TextMatrix(i, ColProd)
            .Range(xlColDesc & Idx) = grid.TextMatrix(i, ColDesc)
            .Range(xlColUnit & Idx) = grid.TextMatrix(i, ColUnit)
            .Range(xlColGroup & Idx) = grid.TextMatrix(i, ColGroup)
            .Range(xlColPremStock & Idx) = Format(grid.TextMatrix(i, ColPremStock), gs_formatQty)
            .Range(xlColPremPrice & Idx) = Format(grid.TextMatrix(i, ColPremPrice), gs_formatPriceIDR)
            .Range(xlColPremMount & Idx) = Format(grid.TextMatrix(i, ColPremMount), gs_formatAmountIDR)
            .Range(xlColIncomeStock & Idx) = Format(grid.TextMatrix(i, ColIncomeStock), gs_formatQty)
            .Range(xlColIncomePrice & Idx) = Format(grid.TextMatrix(i, ColIncomePrice), gs_formatPriceIDR)
            .Range(xlColIncomeMount & Idx) = Format(grid.TextMatrix(i, ColIncomeMount), gs_formatAmountIDR)
            .Range(xlColIncomeOtherStock & Idx) = Format(grid.TextMatrix(i, ColIncomeOtherStock), gs_formatQty)
            .Range(xlColIncomeOtherPrice & Idx) = Format(grid.TextMatrix(i, ColIncomeOtherPrice), gs_formatPriceIDR)
            .Range(xlColIncomeOtherMount & Idx) = Format(grid.TextMatrix(i, ColIncomeOtherMount), gs_formatAmountIDR)
            .Range(xlColOutgoingStock & Idx) = Format(grid.TextMatrix(i, ColOutgoingStock), gs_formatQty)
            .Range(xlColOutgoingPrice & Idx) = Format(grid.TextMatrix(i, ColOutgoingPrice), gs_formatPriceIDR)
            .Range(xlColOutgoingMount & Idx) = Format(grid.TextMatrix(i, ColOutgoingMount), gs_formatAmountIDR)
            .Range(xlColOutgoingOtherStock & Idx) = Format(grid.TextMatrix(i, ColOutgoingOtherStock), gs_formatQty)
            .Range(xlColOutgoingOtherPrice & Idx) = Format(grid.TextMatrix(i, ColOutgoingOtherPrice), gs_formatPriceIDR)
            .Range(xlColOutgoingOtherMount & Idx) = Format(grid.TextMatrix(i, ColOutgoingOtherMount), gs_formatAmountIDR)
            .Range(xlColCurrent & Idx) = Format(grid.TextMatrix(i, ColCurrent), gs_formatQty)
            '.Range(xlColCurrentPrice & idx) = Format(Grid.TextMatrix(i, ColCurrentPrice), gs_formatPriceIDR)
            

            ls_prevCol = xlColCurrentPrice
            
            '#Num Di Set Ke Nol (Header, Browse,Excel) jika tidak ingin menampilkan Additional Cost
            Num = 0
            
            For j = 1 To Num
                ls_prevCol = uf_GetXLColumn(ls_prevCol)
                .Range(ls_prevCol & Idx) = Format(grid.TextMatrix(i, ColCurrentPrice + j), gs_formatPriceIDR)
            Next
            
            .Range(xlColCurrentPriceTotal & Idx) = Format(grid.TextMatrix(i, ColCurrentPriceTotal), gs_formatPriceIDR)
            .Range(xlColCurrentAmount & Idx) = Format(grid.TextMatrix(i, ColCurrentAmount), gs_formatAmountIDR)
            .Range(xlColInventory & Idx) = Format(grid.TextMatrix(i, ColInventory), gs_formatQty)
            .Range(xlColInventoryPrice & Idx) = Format(grid.TextMatrix(i, ColInventoryPrice), gs_formatPriceIDR)
            .Range(xlColInventoryAmount & Idx) = Format(grid.TextMatrix(i, ColInventoryAmount), gs_formatAmountIDR)
        Next
        
        '#Run Macro
        .Columns(xlColProd & ":" & xlColProd).columnWidth = 22.14
        .Columns(xlColDesc & ":" & xlColDesc).columnWidth = 43.29
        .Columns(xlColUnit & ":" & xlColUnit).columnWidth = 5
        .Columns(xlColGroup & ":" & xlColGroup).columnWidth = 22.14
        .Columns(xlColPremStock & ":" & xlColInventoryAmount).columnWidth = 14
        .Range(xlColProd & "2:" & xlColInventoryAmount & "2").Select
        .Selection.Font.Size = 18
        .Range(xlColProd & "9:" & xlColInventoryAmount & Idx).Select
        .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        
        With .Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        With .Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        With .Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        With .Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        With .Selection.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        With .Selection.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
                        
        .Range(xlColPremStock & "11:" & xlColPremStock & Idx).Select
        .Selection.NumberFormat = gs_formatQty
        
        .Range(xlColPremPrice & "11:" & xlColPremPrice & Idx).Select
        .Selection.NumberFormat = gs_formatPriceIDR

        .Range(xlColPremMount & "11:" & xlColPremMount & Idx).Select
        .Selection.NumberFormat = gs_formatAmountIDR

        .Range(xlColIncomeStock & "11:" & xlColIncomeStock & Idx).Select
        .Selection.NumberFormat = gs_formatQty

        .Range(xlColIncomePrice & "11:" & xlColIncomePrice & Idx).Select
        .Selection.NumberFormat = gs_formatPriceIDR

        .Range(xlColIncomeMount & "11:" & xlColIncomeMount & Idx).Select
        .Selection.NumberFormat = gs_formatAmountIDR

        .Range(xlColIncomeOtherStock & "11:" & xlColIncomeOtherStock & Idx).Select
        .Selection.NumberFormat = gs_formatQty

        .Range(xlColIncomeOtherPrice & "11:" & xlColIncomeOtherPrice & Idx).Select
        .Selection.NumberFormat = gs_formatPriceIDR

        .Range(xlColIncomeOtherMount & "11:" & xlColIncomeOtherMount & Idx).Select
        .Selection.NumberFormat = gs_formatAmountIDR

        .Range(xlColOutgoingStock & "11:" & xlColOutgoingStock & Idx).Select
        .Selection.NumberFormat = gs_formatQty

        .Range(xlColOutgoingPrice & "11:" & xlColOutgoingPrice & Idx).Select
        .Selection.NumberFormat = gs_formatPriceIDR

        .Range(xlColOutgoingMount & "11:" & xlColOutgoingMount & Idx).Select
        .Selection.NumberFormat = gs_formatAmountIDR

        .Range(xlColOutgoingOtherStock & "11:" & xlColOutgoingOtherStock & Idx).Select
        .Selection.NumberFormat = gs_formatQty

        .Range(xlColOutgoingOtherPrice & "11:" & xlColOutgoingOtherPrice & Idx).Select
        .Selection.NumberFormat = gs_formatPriceIDR

        .Range(xlColOutgoingOtherMount & "11:" & xlColOutgoingOtherMount & Idx).Select
        .Selection.NumberFormat = gs_formatAmountIDR

        .Range(xlColCurrent & "11:" & xlColCurrent & Idx).Select
        .Selection.NumberFormat = gs_formatQty

        .Range(xlColCurrentPrice & "11:" & xlColCurrentPrice & Idx).Select
        .Selection.NumberFormat = gs_formatPriceIDR

        ls_prevCol = xlColCurrentPrice
        For j = 1 To Num
            ls_prevCol = uf_GetXLColumn(ls_prevCol)
            .Range(ls_prevCol & "11:" & ls_prevCol & Idx).Select
            .Selection.NumberFormat = gs_formatPriceIDR
        Next

        .Range(xlColCurrentPriceTotal & "11:" & xlColCurrentPriceTotal & Idx).Select
        .Selection.NumberFormat = gs_formatPriceIDR

        .Range(xlColCurrentAmount & "11:" & xlColCurrentAmount & Idx).Select
        .Selection.NumberFormat = gs_formatAmountIDR

        .Range(xlColInventory & "11:" & xlColInventory & Idx).Select
        .Selection.NumberFormat = gs_formatQty

        .Range(xlColInventoryPrice & "11:" & xlColInventoryPrice & Idx).Select
        .Selection.NumberFormat = gs_formatPriceIDR

        .Range(xlColInventoryAmount & "11:" & xlColInventoryAmount & Idx).Select
        .Selection.NumberFormat = gs_formatAmountIDR


        .ActiveWindow.ScrollColumn = 1
        .Range("A1").Select
        .Visible = True

        
        .WindowState = xlMaximized
        .ActiveWindow.Zoom = 80
End With
End If

Me.MousePointer = vbDefault
End Sub

Private Function uf_GetXLColumn(ls_PreviousColumn) As String
    If Len(ls_PreviousColumn) = 1 Then
        If Asc(ls_PreviousColumn) + 1 > 122 Then
            uf_GetXLColumn = "a" & Chr(Asc(ls_PreviousColumn) + 1 - 26)
        Else
            uf_GetXLColumn = Chr(Asc(ls_PreviousColumn) + 1)
        End If
    Else
        If Asc(Right(ls_PreviousColumn, 1)) + 1 > 122 Then
            uf_GetXLColumn = Left(ls_PreviousColumn, 1) & Chr(Asc(Right(ls_PreviousColumn, 1)) + 1 - 26)
        Else
            uf_GetXLColumn = Left(ls_PreviousColumn, 1) & Chr(Asc(Right(ls_PreviousColumn, 1)) + 1)
        End If
    End If
End Function

Private Sub CboProductFilter_Click()
    If CboProductFilter.ListIndex <> -1 Then _
        LblProductFilter.Caption = CboProductFilter.Column(1)
    If CboProductFilter = strAll Then CboFilter.Visible = True: Label1(1).Visible = True Else CboFilter.Visible = False: Label1(1).Visible = False
    Call Header
End Sub

Private Sub Product(finish As String, make As String)
    Dim RsProduct As Recordset
    Dim sqlProduct As String, SqlFilter As String
    
    LblProductFilter = ""
    sqlProduct = "Select Item_Code, Item_Name From Item_Master where item_code=item_code "
    
    If CboFinish <> strAll Then
        sqlProduct = sqlProduct + " and finishgoodpart_cls = '" & CboFinish & "'"
    End If
    If CboMake <> strAll Then
        sqlProduct = sqlProduct + " and finishgoodpart_cls = '" & CboMake & "'"
    End If

    Set RsProduct = Db.Execute(sqlProduct)

    With CboProductFilter
        .clear
        .columnCount = 2
        .ColumnWidths = "130pt;200pt"
        .ListWidth = 330
        .ListRows = 15
'        If CboFinish = strAll And CboMake = strAll Then
            .AddItem
            .List(0, 0) = strAll
            .List(0, 1) = strAll
            i = 1
'        Else
'            i = 0
'        End If
        Do While Not RsProduct.EOF
            .AddItem
            .List(i, 0) = Trim(RsProduct("Item_Code"))
            .List(i, 1) = Trim(RsProduct("Item_Name"))
            RsProduct.MoveNext
            i = i + 1
        Loop
    End With
    Set RsProduct = Nothing
    If CboFinish = strAll And CboMake = strAll Then
        CboProductFilter.ListIndex = 0
        CboProductFilter.Text = CboProductFilter.List(0, 0)
    End If

End Sub

'Private Sub Grid_CellChanged(ByVal Row As Long, ByVal Col As Long)
'If Grid.Row >= 1 Then
'        Grid.Cell(flexcpBackColor, Grid.Row, Grid.Col) = &H80000005
'        Grid.Cell(flexcpAlignment, Grid.Row, Grid.Col) = flexAlignRightCenter
'End If
'End Sub
Private Sub Label19_Click()

End Sub
