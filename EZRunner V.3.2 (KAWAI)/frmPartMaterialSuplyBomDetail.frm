VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPartMaterialSuplyBomDetail 
   BackColor       =   &H00FDDFE3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form Parts (Material) Supply [By BOM] Detail"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10530
   Icon            =   "frmPartMaterialSuplyBomDetail.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   10530
   StartUpPosition =   2  'CenterScreen
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   555
      Left            =   8430
      TabIndex        =   25
      Top             =   90
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   979
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Print"
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
      Index           =   4
      Left            =   7830
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   8370
      Width           =   1080
   End
   Begin VB.CommandButton cmdAction 
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
      Index           =   3
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   8370
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.CommandButton cmdAction 
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
      Index           =   2
      Left            =   6630
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   8370
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.CommandButton cmdAction 
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
      Left            =   7830
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   8370
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   615
      Left            =   270
      TabIndex        =   15
      Top             =   7620
      Width           =   10005
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
         Height          =   300
         Left            =   150
         TabIndex        =   16
         Top             =   180
         Width           =   9600
      End
   End
   Begin VB.CommandButton cmdAction 
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
      Left            =   9060
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8370
      Width           =   1080
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
      Left            =   270
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8370
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   2745
      Left            =   270
      TabIndex        =   0
      Top             =   690
      Width           =   10005
      Begin VB.TextBox txtbcno 
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
         Left            =   7740
         TabIndex        =   26
         Top             =   1800
         Width           =   2085
      End
      Begin VB.TextBox txt_set 
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
         Left            =   4230
         TabIndex        =   23
         Top             =   2280
         Width           =   2385
      End
      Begin MSComCtl2.DTPicker dt_supply 
         Height          =   330
         Left            =   2220
         TabIndex        =   1
         Top             =   2265
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   582
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
         CustomFormat    =   "dd MMM yyyy"
         Format          =   141230083
         CurrentDate     =   39287
      End
      Begin MSComCtl2.DTPicker DTPbcdate 
         Height          =   345
         Left            =   7740
         TabIndex        =   27
         Top             =   2280
         Width           =   2085
         _ExtentX        =   3678
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
         CustomFormat    =   "dd MMM yyyy"
         Format          =   141230083
         CurrentDate     =   41080
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "BC No."
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
         Left            =   6720
         TabIndex        =   31
         Top             =   1860
         Width           =   735
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "BC Type"
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
         Left            =   6720
         TabIndex        =   30
         Top             =   1350
         Width           =   855
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "BC Date"
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
         Left            =   6720
         TabIndex        =   29
         Top             =   2310
         Width           =   885
      End
      Begin MSForms.ComboBox cbobctype 
         Height          =   315
         Left            =   7740
         TabIndex        =   28
         Top             =   1290
         Width           =   2085
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "3678;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Set"
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
         Left            =   3840
         TabIndex        =   24
         Top             =   2370
         Width           =   285
      End
      Begin VB.Label lbl_parent 
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
         Left            =   10620
         TabIndex        =   12
         Top             =   1785
         Width           =   60
      End
      Begin MSForms.ComboBox cbo_status 
         Height          =   315
         Left            =   90
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   780
         Width           =   1125
         VariousPropertyBits=   612386841
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "1984;556"
         ShowDropButtonWhen=   2
         Value           =   "cbo_status"
         FontName        =   "Verdana"
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cbo_supply 
         Height          =   315
         Left            =   2220
         TabIndex        =   10
         Top             =   780
         Width           =   2475
         VariousPropertyBits=   612386841
         MaxLength       =   25
         DisplayStyle    =   3
         Size            =   "4366;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supply No"
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
         Left            =   1290
         TabIndex        =   9
         Top             =   840
         Width           =   870
      End
      Begin VB.Line Line2 
         X1              =   3750
         X2              =   6600
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line1 
         X1              =   3750
         X2              =   6570
         Y1              =   1530
         Y2              =   1530
      End
      Begin VB.Label lbl_location 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lbl_location"
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
         Left            =   3750
         TabIndex        =   8
         Top             =   1800
         Width           =   2490
      End
      Begin VB.Label lbl_warehouse 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lbl_warehouse"
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
         Left            =   3750
         TabIndex        =   7
         Top             =   1305
         Width           =   3210
      End
      Begin MSForms.ComboBox cbo_location 
         Height          =   330
         Left            =   2220
         TabIndex        =   6
         Top             =   1710
         Width           =   1500
         VariousPropertyBits=   746604569
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2646;582"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         Value           =   "cbo_location"
         FontName        =   "Verdana"
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cbo_warehouse 
         Height          =   330
         Left            =   2220
         TabIndex        =   5
         Top             =   1230
         Width           =   1500
         VariousPropertyBits=   746604569
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2646;582"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         Value           =   "cbo_warehouse"
         FontName        =   "Verdana"
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supply Date"
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
         Left            =   105
         TabIndex        =   4
         Top             =   2340
         Width           =   1050
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Location CD"
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
         Left            =   105
         TabIndex        =   3
         Top             =   1785
         Width           =   1305
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Warehouse CD"
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
         Left            =   105
         TabIndex        =   2
         Top             =   1305
         Width           =   1785
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid grid 
      Height          =   3900
      Left            =   240
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3510
      Width           =   10035
      _cx             =   17701
      _cy             =   6879
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
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   3
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   275
      RowHeightMax    =   275
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
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
      ExplorerBar     =   1
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Parts (Material) Supply [By BOM] Detail"
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
      Index           =   0
      Left            =   330
      TabIndex        =   18
      Top             =   60
      Width           =   9915
   End
End
Attribute VB_Name = "frmPartMaterialSuplyBomDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db2 As New ADODB.Connection
Dim ColMaterialCode, ColDescription, ColQty, ColUnit, colUnitCls, colQtyRev As Byte
Dim colQtyLama As Byte
Public PODate As Date
Sub Header()
    ColMaterialCode = 1
    ColDescription = 2
    ColQty = 3
    colQtyRev = 4
    ColUnit = 5
    colUnitCls = 6
    colQtyLama = 7
    grid.ColS = 8
    grid.Rows = 1
    grid.TextMatrix(0, ColMaterialCode) = "Material Code"
    grid.TextMatrix(0, ColDescription) = "Description"
    grid.TextMatrix(0, ColQty) = "Qty BOM"
    grid.TextMatrix(0, colQtyRev) = "Set Qty"
    grid.TextMatrix(0, ColUnit) = "Unit"
    grid.TextMatrix(0, colUnitCls) = "ColUnitCls"
    grid.TextMatrix(0, colQtyLama) = "ColQtyLama"
    
    grid.ColHidden(colUnitCls) = True
    grid.ColWidth(0) = 400
    grid.ColHidden(0) = True
    grid.ColWidth(ColMaterialCode) = 1500
    
    grid.ColWidth(ColDescription) = 3500
    grid.ColWidth(ColUnit) = 800
    grid.ColWidth(ColQty) = 1000
    grid.EditMaxLength = 1
    grid.ColHidden(colQtyLama) = True
End Sub
'ambil  item code yang di pake di
Function getItemCode() As String
Dim i As Long
getItemCode = ""
 With frmPartMaterialSupplyByBom.grid
 For i = 1 To .Rows - 1
 
  If .Cell(flexcpChecked, i, 0) = flexChecked Then
    If getItemCode = "" Then
        getItemCode = "'" & .TextMatrix(i, 5) & "'"
    Else
        getItemCode = getItemCode & ",'" & .TextMatrix(i, 5) & "'"
    End If
  End If
 Next
 End With
End Function
Private Function GetQty(Item_Code As String) As Integer
Dim s As String
Dim jmlParent As Integer
Dim i As Long
Dim Rget As New ADODB.Recordset
Set Rget = Nothing
s = "SELECT  qty,parent_ItemCode FROM Bom_Master WHERE Parent_ItemCode IN( " & getItemCode & ") and Item_Code='" & Item_Code & "'"
If Rget.State <> adStateClosed Then Rget.Close
Rget.Open s, Db, adOpenStatic, adLockReadOnly
GetQty = 0
While Not Rget.EOF
'ambil jml parent dari grid
    For i = 1 To frmPartMaterialSupplyByBom.grid.Rows - 1
        If RTrim(frmPartMaterialSupplyByBom.grid.TextMatrix(i, 5)) = RTrim(Rget!parent_itemcode) Then
         jmlParent = frmPartMaterialSupplyByBom.grid.TextMatrix(i, 8) 'jumlah qtyset
         Exit For
        End If
    Next
GetQty = GetQty + (CDbl(Rget!Qty) * CDbl(jmlParent))
Rget.MoveNext
Wend



End Function

Sub BrowseGrid(Optional StrQueeing As String)
Dim rd As New ADODB.Recordset
Set rd = Nothing
Dim sbom As String
Dim SupplyBom As Boolean
Dim Test As String

SupplyBom = False
'cek apakah ada data di PartBom detail
If rd.State <> adStateClosed Then rd.Close

sbom = " select distinct ChildItem_Code Item_Code,(Select Item_name From Item_Master WHERE item_Code=ChildItem_Code)ItemDesc,"
sbom = sbom & " remarks  as TotQty_BOM,childRequirement_Qty TotQty_Supply,Childunit_Cls unit_cls,(select Description FROM Unit_Cls WHERE Unit_Cls=Childunit_cls)Unit_Desc"
sbom = sbom & " , BC_Type, BC40_No,BC40_Date FROM part_Supply where SupplyRec_NO='" & cbo_supply & "'"
rd.Open sbom, Db, adOpenDynamic, adLockReadOnly
Test = "1"
If rd.EOF Then
    cboBCType.Text = ""
    txtBCNo.Text = ""
    dtpBCDate.Value = Now()
    SupplyBom = True
    sbom = "select item_code,(SELECT Item_Name FROM Item_Master WHERE Item_Code=Bm.Item_Code)ItemDesc,sum(qty) TotQty_BOM,TotQty_Supply=sum(qty),(Select Description FROM Unit_Cls where Unit_Cls=Bm.Unit_cLS)Unit_Desc "
    sbom = sbom & vbCrLf & ",Unit_Cls FROM bom_Master BM where Parent_ItemCode IN ( " & getItemCode & ")"
    sbom = sbom & vbCrLf & " group by Item_Code,Unit_Cls"
    If rd.State <> adStateClosed Then rd.Close
    Test = "2"
Else
    cboBCType.Text = IIf(IsNull(rd!BC_Type), "", (rd!BC_Type))
    txtBCNo.Text = rd!BC40_No & ""
    dtpBCDate.Value = IIf(IsNull(rd!BC40_Date), Now(), rd!BC40_Date)
End If

If rd.State <> adStateClosed Then rd.Close

If StrQueeing = "" Then
    rd.Open sbom, Db, adOpenStatic, adLockReadOnly
Else
    rd.Open StrQueeing, Db, adOpenStatic, adLockReadOnly
End If

Call Header
While Not rd.EOF

    grid.Rows = grid.Rows + 1
    grid.Cell(flexcpBackColor, grid.Rows - 1, 0) = vbWhite
    grid.TextMatrix(grid.Rows - 1, ColMaterialCode) = rd!Item_Code
    grid.TextMatrix(grid.Rows - 1, ColDescription) = rd!ItemDesc
    
    grid.TextMatrix(grid.Rows - 1, ColQty) = Format(rd!TotQty_BOM, gs_formatQty) ' GetQty(rd!Item_Code) ' Format(rd!jml, gs_formatQty)
    grid.TextMatrix(grid.Rows - 1, colQtyRev) = Format(rd!TotQty_Supply, gs_formatQty) 'Format(grid.TextMatrix(grid.Rows - 1, ColQty), gs_formatQty)
    
    grid.TextMatrix(grid.Rows - 1, ColUnit) = uf_GetUnitDescription(IIf(IsNull(rd!Unit_cls), "", rd!Unit_cls))
    grid.TextMatrix(grid.Rows - 1, colUnitCls) = IIf(IsNull(rd!Unit_cls), "", rd!Unit_cls)
    grid.Cell(flexcpBackColor, grid.Rows - 1, colQtyRev) = vbWhite
    rd.MoveNext
Wend

End Sub
Private Sub comboBCtype()
Dim ls_sql As String
Dim rs_combo As New ADODB.Recordset
Dim i As Long


cboBCType.columnCount = 1
cboBCType.clear

ls_sql = "select bc_type from BC_master"
rs_combo.Open ls_sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
i = 0

Do While Not rs_combo.EOF
cboBCType.AddItem rs_combo("Bc_type")
rs_combo.MoveNext
Loop

cboBCType.ColumnWidths = "90"
cboBCType.ListWidth = 90
cboBCType.ListRows = 7


End Sub
Private Sub cmdAction_Click(Index As Integer)
On Error GoTo ss
Select Case Index
Case 0 'submit
Dim RS As New ADODB.Recordset
Dim i, j As Integer
Dim TempQty As Double

Set RS = Nothing
Dim s, Y As String
MousePointer = vbHourglass
lblerror = ""
'update master Supply by bom
s = "SELECT * FROM SupplyBom_Master WHERE Supply_No='" & cbo_supply & "'"
If RS.State <> adStateClosed Then RS.Close
RS.Open s, Db, adOpenDynamic, adLockOptimistic
If RS.EOF Then
    RS.AddNew
    RS!Supply_no = cbo_supply
    RS!Register_Date = Now()
End If
RS!FromWH_Code = cbo_warehouse
RS!ToWH_Code = cbo_location
RS!Supply_Date = dt_supply
RS!Last_Update = Now()
RS.update
RS.Close
DoEvents
'Updaet Detail Supply by bom
With frmPartMaterialSupplyByBom.grid

    For i = 1 To .Rows - 1
      If .Cell(flexcpChecked, i, 0) = flexChecked Then
          s = "SELECT * FROM SupplyBom_Detail WHERE Supply_No='" & cbo_supply & "'"
          s = s & "AND Item_Code='" & .TextMatrix(i, 5) & "' AND Po_No='" & .TextMatrix(i, 3) & "'"
          If RS.State <> adStateClosed Then RS.Close
          RS.Open s, Db, adOpenDynamic, adLockOptimistic
          If RS.EOF Then
            RS.AddNew
            RS!Supply_no = cbo_supply
            RS!Item_Code = .TextMatrix(i, 5)
            RS!po_no = .TextMatrix(i, 3)
            RS!Register_Date = Now()
          End If
           RS!po_date = PODate
          RS!Supplier_Code = .TextMatrix(i, 1)
          RS!Qty = .TextMatrix(i, 8)
          RS!Unit_cls = Get_Record("SELECT Unit_Cls FROM Unit_Cls WHERE description='" & .TextMatrix(i, 10) & "'")
          RS!Last_Update = Now()
          RS!last_user = userLogin
          RS.update
      End If
    Next

End With

DoEvents
Dim ParentCode() As String
Dim insert_Status As Boolean
TempQty = 0
    ParentCode = Split(getItemCode, ",")
    For i = 1 To grid.Rows - 1
    s = "SELECT * FROM part_Supply WHERE SupplyRec_NO='" & cbo_supply & "'"
    s = s & " AND ChildItem_Code='" & grid.TextMatrix(i, ColMaterialCode) & "'"
        For j = 0 To UBound(ParentCode)
            Y = s & " AND ParentItem_Code=" & ParentCode(j) & ""
            If RS.State <> adStateClosed Then RS.Close
            TempQty = 0
            RS.Open Y, Db, adOpenDynamic, adLockOptimistic
            If RS.EOF Then
                RS.AddNew
                RS!Register_Date = Now()
                RS!supplyRec_No = cbo_supply
                insert_Status = True
            Else
                insert_Status = False
            End If
            RS!from_address = ""
            RS!FromWarehouse_Code = cbo_warehouse
            RS!towarehouse_code = cbo_location
            RS!childsupply_date = Format(dt_supply, "yyyy-mm-dd")
            RS!childitem_code = grid.TextMatrix(i, ColMaterialCode)
            TempQty = RS!ChildRequirement_qty
            RS!ChildRequirement_qty = grid.TextMatrix(i, colQtyRev)
            RS!childunit_cls = grid.TextMatrix(i, colUnitCls)
            RS!parentItem_code = Replace(ParentCode(j), "'", "")
            RS!Remarks = grid.TextMatrix(i, ColQty)
            RS!last_user = userLogin
            RS!Last_Update = Now()
            RS!supply_cls = "S1"
            RS!do_no = ""
            RS!BC40_No = txtBCNo
            RS!BC_Type = cboBCType
            RS!BC40_Date = dtpBCDate
        
            RS.update
            
            
            
            
            If insert_Status Then
                'insert
            FromControlCls = "01"
            Call up_UpdateStockMaster(Format(Trim(dt_supply.Value), "yyyy-MM-dd"), _
            uf_GetLastClosing("month"), uf_GetLastClosing("year"), _
            Trim(cbo_warehouse.Text), Trim(cbo_location.Text), _
            Trim(grid.TextMatrix(i, ColMaterialCode)), _
             CDbl(grid.TextMatrix(i, colQtyRev)), "S1", _
            "01", "", "I", "", "", False, False, True, Db)
             FromControlCls = ""
'
'            Call up_UpdateStockMaster(Format(Trim(dt_supply.Value), "yyyy-MM-dd"), _
'            uf_GetLastClosing("month"), uf_GetLastClosing("year"), _
'            Trim(cbo_warehouse.Text), Trim(cbo_location.Text), _
'            Trim(grid.TextMatrix(i, ColMaterialCode)), _
'            CDbl(grid.TextMatrix(i, colQtyRev)), "R", _
'            "01", "", "I", "", "", False, False, True, Db)
'
            
            'Trim(.TextMatrix(i, bteColStockControlCls)), "", "I", "", "", False, False, True, Db_ps)
            
            Else
             FromControlCls = "01"
             'update
             'Call up_UpdateStockMaster(Format(Trim(dt_supply.Value), "yyyy-MM-dd"), _
             uf_GetLastClosing("month"), uf_GetLastClosing("year"), Trim(cbo_warehouse), Trim(cbo_location), _
             Trim(l_item_code_update), 0 - CDbl(l_update_stock), Trim(cbo_supply), Trim(l_stock_location), "", "U", "", "", False, False, True, db2)
             TempQty = TempQty - CDbl(grid.TextMatrix(i, colQtyRev))
             Call up_UpdateStockMaster(Format(Trim(dt_supply.Value), "yyyy-MM-dd"), _
             uf_GetLastClosing("month"), uf_GetLastClosing("year"), _
             Trim(cbo_warehouse.Text), Trim(cbo_location.Text), _
             Trim(grid.TextMatrix(i, ColMaterialCode)), _
              0 - TempQty, "S1", _
             "01", "", "U", "", "", False, False, True, Db)
             
             FromControlCls = ""
             
            End If
        Next

  Next
   DoEvents
   lblerror = DisplayMsg(8004)
   MousePointer = vbDefault
   Call BrowseGrid
   cbo_status = "Update"
Case 4 'sub menu
    'Print
    'Rpt_supply_By_Bom
    Call PrintSupplyByBom
    
Case 1 'clear
    Call Kosong
Case 2 'cancel
Call BrowseGrid
Case 3
If (MsgBox("Are you sure want to delete?", vbQuestion + vbDefaultButton2 + vbYesNo, "Confirmation") = vbNo) Then Exit Sub
        If Get_Record("SELECT * FROM part_Supply where SupplyRec_no='" & cbo_supply & "'") = "" Then
            lblerror = DisplayMsg(4047)
            Exit Sub
        End If
        Db.Execute "DELETE FROM Part_Supply WHERE SupplyRec_no='" & cbo_supply & "'"
    
        Dim X As Byte
        For X = 1 To grid.Rows - 1
             '# Delete data from stock Master
            Call up_UpdateStockMaster(Format(dt_supply, "yyyy-MM-dd"), uf_GetLastClosing("month"), uf_GetLastClosing("year"), Trim(cbo_warehouse), Trim(cbo_location), Trim(grid.TextMatrix(X, ColMaterialCode)), (CDbl(grid.TextMatrix(X, colQtyLama)) * -1), "S1", "01", "", "D", "", "", False, False, True, Db)
        
            '# Erase data from stockMaster ( base on From WareHouse code)
            Call up_EraseBlankDataInStockMaster(Trim(cbo_warehouse), Trim(grid.TextMatrix(i, ColMaterialCode)), Trim(grid.TextMatrix(X, ColDescription)))
        
            '# Erase data from stockMaster ( base on To WareHouse code)
            Call up_EraseBlankDataInStockMaster(Trim(cbo_location), Trim(grid.TextMatrix(X, ColMaterialCode)), Trim(grid.TextMatrix(X, ColDescription)))
        
        Next
        Db.Execute "DELETE FROM SupplyBom_Detail WHERE Supply_no='" & cbo_supply & "'"
        Db.Execute "DELETE FROM SupplyBom_Master WHERE Supply_no='" & cbo_supply & "'"
        grid.Rows = 1
        lblerror = DisplayMsg(1201)

End Select
Exit Sub
ss:
lblerror = err.number & "-" & err.Description
MousePointer = vbDefault
End Sub
'
    Private Sub PrintSupplyByBom()
    
    Dim sbom As String
    Dim rsRpt As New ADODB.Recordset
    Set rsRpt = Nothing
    Dim application As New CRAXDDRT.application
    Dim report As New CRAXDDRT.report
    
    Dim SqlRpt As String
    Dim Rpt As New FrmRpt3
    lblerror = ""
   sbom = " select distinct SupplyRec_NO,fromWarehouse_code,"
   sbom = sbom & vbCrLf & "isnull((select wh_name  from warehouse_master where wh_code=fromWarehouse_Code),'')FromWhName,"
   sbom = sbom & vbCrLf & " ToWarehouse_code,(select wh_name  from warehouse_master where wh_code=ToWarehouse_Code)ToWHCode,"
   sbom = sbom & vbCrLf & " ChildSupply_date,"
   sbom = sbom & vbCrLf & " ChildItem_Code Item_Code,(Select Item_name From Item_Master"
   sbom = sbom & vbCrLf & " WHERE item_Code=ChildItem_Code)ItemDesc, remarks  as jml,childRequirement_Qty JmlRev,Childunit_Cls unit_cls,"
   sbom = sbom & vbCrLf & " (select Description FROM Unit_Cls WHERE Unit_Cls=Childunit_cls)Unit_Desc, ParentItem_code FROM part_Supply PS"
   sbom = sbom & " where SupplyRec_NO='" & cbo_supply & "'"
    
    If rsRpt.State <> adStateClosed Then rsRpt.Close
    rsRpt.Open sbom, Db, adOpenDynamic, adLockOptimistic
    sqlprint = sbom
    If rsRpt.EOF Then
        lblerror = DisplayMsg(13)
        MousePointer = vbDefault
    Exit Sub
    End If
    Set report = application.OpenReport(App.path & "\Reports\Rpt_supply_By_Bom.rpt")
    reportcode = "RptSupplyByBom"
    printorient = "2"
    report.Database.Tables(1).SetDataSource rsRpt
    report.FormulaFields(1).Text = "" & gi_decimalDigitQtyBOM & ""
    Rpt.CRViewer1.ReportSource = report
    Rpt.CRViewer1.ViewReport
    Rpt.CRViewer1.Zoom 1
    Rpt.WindowState = 2
    Rpt.Show 1
    End Sub
Private Sub CmdSubMenu_Click()
If Get_Record("select SupplyRec_NO  FROM Part_supply WHERE SupplyRec_NO ='" & cbo_supply & "'") <> "" Then
    frmPartMaterialSupplyByBom.cbo_supply = cbo_supply
    frmPartMaterialSupplyByBom.cbo_status = "Update"
    frmPartMaterialSupplyByBom.cbo_supply = cbo_supply
    'frmPartMaterialSupplyByBom.BrowseGrid
'    frmPartMaterialSupplyByBom.Show
 '   frmPartMaterialSupplyByBom.BrowseGrid
End If
Unload Me
'frmPartMaterialSuplyBomDetail.Refresh
End Sub

Private Sub Form_Load()
'Call BrowseGrid
Call comboBCtype
End Sub
Sub Kosong()
cbo_supply.Text = ""
cbo_warehouse.Text = ""
lbl_warehouse.Caption = ""
cbo_location.Text = ""
lbl_location.Caption = ""


lblerror.Caption = ""
Call Header
End Sub



Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Col = colQtyRev Then
        grid.TextMatrix(Row, colQtyRev) = Format(CDbl(grid.TextMatrix(Row, colQtyRev)), gs_formatAmountIDR)
        
End If
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col <> colQtyRev And Col <> 0 Then Cancel = True
If Col = 0 Then
grid.EditMaxLength = 1
Else
grid.EditMaxLength = 6
End If




End Sub

Private Sub Grid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
If Col <> colQtyRev And Col <> 0 Then KeyAscii = 0
If grid.Col = 0 Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii <> Asc("D") And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then _
            KeyAscii = 0
        If KeyAscii = Asc(".") Then KeyAscii = 0
End If
If Col = colQtyRev Then
        If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then _
            KeyAscii = 0
End If
End Sub

