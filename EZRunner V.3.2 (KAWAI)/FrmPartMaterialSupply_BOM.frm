VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmPartMaterialSupply_BOM 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Parts (Material) Supply [By BOM]"
   ClientHeight    =   10530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15120
   Icon            =   "FrmPartMaterialSupply_BOM.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10530
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd_preview 
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
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   9930
      Width           =   1125
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   615
      Left            =   300
      TabIndex        =   27
      Top             =   8970
      Width           =   14655
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
         Left            =   60
         TabIndex        =   28
         Top             =   210
         Width           =   14490
      End
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
      Index           =   3
      Left            =   9690
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   9930
      Width           =   1200
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
      Left            =   13755
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   9930
      Width           =   1200
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
      Left            =   11010
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9930
      Width           =   1200
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
      Index           =   1
      Left            =   12390
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9930
      Width           =   1200
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
      Left            =   300
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   9930
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   2745
      Left            =   300
      TabIndex        =   15
      Top             =   870
      Width           =   14655
      Begin VB.CommandButton cmdBrowser 
         Caption         =   "..."
         Height          =   315
         Left            =   10200
         TabIndex        =   33
         Top             =   1740
         Width           =   300
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
         Left            =   8160
         TabIndex        =   8
         Top             =   2280
         Width           =   1965
      End
      Begin VB.CommandButton cmd_search 
         BackColor       =   &H0080FFFF&
         Caption         =   "Searc&h"
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
         Left            =   10260
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2250
         Width           =   1125
      End
      Begin MSComCtl2.DTPicker dt_supply 
         Height          =   330
         Left            =   2430
         TabIndex        =   6
         Top             =   2272
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
         Format          =   334823427
         CurrentDate     =   39287
      End
      Begin MSComCtl2.DTPicker dt_from 
         Height          =   330
         Left            =   2430
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   232
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
         Format          =   334823427
         CurrentDate     =   39105
      End
      Begin MSComCtl2.DTPicker dt_to 
         Height          =   330
         Left            =   4440
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   232
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
         Format          =   334823427
         CurrentDate     =   39289
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
         TabIndex        =   32
         Top             =   1785
         Width           =   60
      End
      Begin VB.Line Line3 
         X1              =   10560
         X2              =   14400
         Y1              =   2055
         Y2              =   2055
      End
      Begin MSForms.ComboBox cbo_status 
         Height          =   315
         Left            =   180
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   780
         Width           =   1125
         VariousPropertyBits=   612386843
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "1984;556"
         ShowDropButtonWhen=   2
         Value           =   "cbo_status"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cbo_parent 
         Height          =   315
         Left            =   8160
         TabIndex        =   7
         Top             =   1725
         Width           =   1965
         VariousPropertyBits=   612386843
         MaxLength       =   15
         DisplayStyle    =   3
         Size            =   "3466;556"
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
         Left            =   7710
         TabIndex        =   26
         Top             =   2340
         Width           =   285
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Parent"
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
         Left            =   7440
         TabIndex        =   25
         Top             =   1785
         Width           =   555
      End
      Begin MSForms.ComboBox cbo_supply 
         Height          =   315
         Left            =   2430
         TabIndex        =   3
         Top             =   780
         Width           =   2475
         VariousPropertyBits=   612386843
         MaxLength       =   25
         DisplayStyle    =   3
         Size            =   "4366;556"
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
         Left            =   1410
         TabIndex        =   24
         Top             =   840
         Width           =   870
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
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
         Left            =   4080
         TabIndex        =   23
         Top             =   300
         Width           =   210
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supply Date From"
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
         Left            =   210
         TabIndex        =   22
         Top             =   300
         Width           =   1545
      End
      Begin VB.Line Line2 
         X1              =   4050
         X2              =   7110
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line1 
         X1              =   4050
         X2              =   7110
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
         Left            =   4050
         TabIndex        =   20
         Top             =   1785
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
         Left            =   4050
         TabIndex        =   19
         Top             =   1305
         Width           =   3210
      End
      Begin MSForms.ComboBox cbo_location 
         Height          =   330
         Left            =   2430
         TabIndex        =   5
         Top             =   1710
         Width           =   1500
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2646;582"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         Value           =   "cbo_location"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cbo_warehouse 
         Height          =   330
         Left            =   2430
         TabIndex        =   4
         Top             =   1237
         Width           =   1500
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2646;582"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         Value           =   "cbo_warehouse"
         FontName        =   "Verdana"
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
         Left            =   225
         TabIndex        =   18
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
         Left            =   225
         TabIndex        =   17
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
         Left            =   225
         TabIndex        =   16
         Top             =   1305
         Width           =   1785
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   4920
      Left            =   300
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   3840
      Width           =   14655
      _cx             =   25850
      _cy             =   8678
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
      Cols            =   5
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
      Left            =   13095
      TabIndex        =   30
      Top             =   217
      Width           =   1860
      _extentx        =   3281
      _extenty        =   741
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Parts (Material) Supply [By BOM]"
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
      Left            =   360
      TabIndex        =   21
      Top             =   240
      Width           =   14505
   End
End
Attribute VB_Name = "FrmPartMaterialSupply_BOM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bteColMaterialCode As Byte
Dim bteColDescription As Byte
Dim bteColQtyAwal As Byte
Dim bteColQtyAkhir As Byte
Dim bteColUnit As Byte
Dim bteColUnitCls As Byte
Dim bteColQtyLama As Byte 'Hidden utk QtyAkhir yg lama
Dim bteColSeqNo As Byte
Dim bteColStockControlCls As Byte
Dim sql As String
Dim Month, Year
Dim Db_ps As New ADODB.Connection
Dim rsRpt As New ADODB.Recordset
Dim PODate
Sub Header()
    bteColMaterialCode = 0
    bteColDescription = 1
    bteColQtyAwal = 2
    bteColQtyAkhir = 3
    bteColUnit = 4
    bteColUnitCls = 5
    bteColQtyLama = 6
    bteColSeqNo = 7
    bteColStockControlCls = 8
                
    With grid
        .clear
        .Rows = 1
        .ColS = 9
        
        .TextMatrix(0, bteColMaterialCode) = "Material Code"
        .TextMatrix(0, bteColDescription) = "Description"
        .TextMatrix(0, bteColQtyAwal) = "Qty"
        .TextMatrix(0, bteColQtyAkhir) = "Qty"
        .TextMatrix(0, bteColUnit) = "Unit"
        .TextMatrix(0, bteColUnitCls) = "Unit Cls"
        .TextMatrix(0, bteColQtyLama) = "QtyLama"
        .TextMatrix(0, bteColSeqNo) = "SeqNo"
        .TextMatrix(0, bteColStockControlCls) = "CtrlCls"
                        
        .ColWidth(bteColMaterialCode) = 1500
        .ColWidth(bteColDescription) = 4500
        .ColWidth(bteColQtyAwal) = 1500
        .ColWidth(bteColQtyAkhir) = 1500
        .ColWidth(bteColUnit) = 1000
        .ColWidth(bteColUnitCls) = 1000
        .ColWidth(bteColQtyLama) = 1500
        .ColWidth(bteColSeqNo) = 1000
        .ColWidth(bteColStockControlCls) = 1000
                
        .Cell(flexcpAlignment, 0, 0, 0, grid.ColS - 1) = flexAlignCenterCenter
        .ColAlignment(bteColMaterialCode) = flexAlignLeftCenter
        .ColAlignment(bteColDescription) = flexAlignLeftCenter
        .ColAlignment(bteColQtyAwal) = flexAlignRightCenter
        .ColAlignment(bteColQtyAkhir) = flexAlignRightCenter
        .ColAlignment(bteColUnit) = flexAlignLeftCenter
        .ColAlignment(bteColUnitCls) = flexAlignLeftCenter
        .ColAlignment(bteColQtyLama) = flexAlignRightCenter
        .ColAlignment(bteColSeqNo) = flexAlignLeftCenter
        .ColAlignment(bteColStockControlCls) = flexAlignLeftCenter
        
        .ColHidden(bteColUnitCls) = True
        .ColHidden(bteColQtyLama) = True
        .ColHidden(bteColSeqNo) = True
        .ColHidden(bteColStockControlCls) = True
        .MergeCells = flexMergeRestrictColumns
        .MergeRow(0) = True
        
    End With
End Sub

Private Sub cmd_preview_Click()
If cbo_status.ListIndex = 1 Then
 Exit Sub
ElseIf cbo_supply.Text = "" Then
 lblerror = DisplayMsg(8107) 'Please Select Supply No!
 cbo_supply.SetFocus
 Exit Sub
Else
 Me.MousePointer = vbHourglass
  
 sql = "select " & _
       "rtrim(ps.SupplyRec_No) as [Supply No], " & _
       "rtrim(bm.Parent_ItemCode) as [Parent Item Code],  rtrim(im.makeritem_code) MakerItem_Code, " & _
       "(select rtrim(item_name) from item_master where item_code = bm.parent_itemcode) as ParentDesc, " & _
       "rtrim(ps.Remarks) as [Qty Set], " & _
       "ps.ChildSupply_date as [Supply Date], " & _
       "rtrim(ps.FromWarehouse_Code) as [From Location], " & _
       "case when (Select rtrim(trade_name) from trade_master where trade_code = ps.fromwarehouse_code) is null " & _
       "then (Select rtrim(wh_name) from warehouse_master where wh_code = ps.fromwarehouse_code) " & _
       "else (Select rtrim(trade_name) from trade_master where trade_code = ps.fromwarehouse_code) end From_name, " & _
       "rtrim(ps.ToWarehouse_Code) as [To Location], " & _
       "case when (Select rtrim(trade_name) from trade_master where trade_code = ps.towarehouse_code) is null " & _
       "then (Select rtrim(wh_name) from warehouse_master where wh_code = ps.towarehouse_code) " & _
       "else (Select rtrim(trade_name) from trade_master where trade_code = ps.towarehouse_code) end To_name, " & _
       "rtrim(ps.childitem_code) as [Material Code], " & _
       "rtrim(im.item_name) as Description, " & _
       "isnull(bm.qty,0) as [Qty BOM], " & _
       "ps.childrequirement_qty as [Qty Supply], " & _
       "rtrim(ps.childunit_cls) as unit_cls, " & _
       "rtrim(uc.description) as Unit, " & _
       "ps.seq_no, " & _
       "rtrim(company_name) As company_name " & _
       "from part_supply ps " & _
       "left join bom_master bm " & _
       "on ps.childitem_code = bm.item_code "
 sql = sql & _
       "and ps.ParentItem_Code = bm.Parent_ItemCode " & _
       "left join Item_Master im " & _
       "on ps.ChildItem_Code = im.Item_Code " & _
       "left join unit_cls uc " & _
       "on ps.childunit_cls= uc.unit_cls, " & _
       "company_profile " & _
       "where ps.supplyrec_no = '" & cbo_supply.Text & "' "
       
 sqlprint = sql
     
 If rsRpt.State <> adStateClosed Then rsRpt.Close
 rsRpt.Open sql, Db, adOpenDynamic, adLockOptimistic
 If rsRpt.EOF Then lblerror.Caption = DisplayMsg(4006): Me.MousePointer = vbDefault: Exit Sub

 Call Preview

 Me.MousePointer = vbDefault
End If
End Sub
Sub Preview()
Dim application As New CRAXDDRT.application
Dim report As New CRAXDDRT.report

Dim Rpt As New FrmRpt3
          
Set report = application.OpenReport(App.path & "\Reports\rptPartMaterialSupplyBOM.rpt")
report.Database.Tables(1).SetDataSource rsRpt
report.FormulaFields(1).Text = gi_decimalDigitQtyBOM
report.FormulaFields(4).Text = gi_decimalDigitQty
         
report.ReportTitle = "Part Material Supply BOM"
reportcode = "pmsBOM"
printorient = 1
            
Rpt.CRViewer1.ReportSource = report
Rpt.CRViewer1.ViewReport
Rpt.CRViewer1.Zoom (75)
        
Rpt.WindowState = 2
Rpt.Show 1
End Sub
              
Private Sub cmdAction_Click(Index As Integer)
Dim i As Long, rsseqno As New ADODB.Recordset, VSeqNo As Double

On Error GoTo errHandler
lblerror.Caption = ""

With grid
Select Case (Index)
    Case 0: 'Submit
    
    If hakUpdate(Me.Name) = 0 Then
        lblerror = DisplayMsg(3008)
        Exit Sub
    End If
    
    lblerror = up_ValidateDateRange(dt_supply.Value, True)
    If Trim(lblerror) <> "" Then Exit Sub

     If cbo_status.ListIndex = 0 Then
         If cbo_supply.Text = "" Then Exit Sub
         If cbo_supply.MatchFound = False Then
            Header
            lblerror = DisplayMsg(8093)
            cbo_supply.SetFocus
            Exit Sub
         End If
      End If
         
     If cbo_supply.Text = "" Then
      lblerror = DisplayMsg(8107) 'Please Select Supply No!
      cbo_supply.SetFocus
     ElseIf cbo_warehouse.Text = "" Then
      lblerror = DisplayMsg(8067)
      cbo_warehouse.SetFocus
     ElseIf cbo_location.Text = "" Then
      lblerror = DisplayMsg(8069)
      cbo_location.SetFocus
     ElseIf cbo_parent.Text = "" Then
      lblerror = DisplayMsg(1008)
      cbo_parent.SetFocus
     ElseIf txt_set.Text = "" Then
      lblerror = DisplayMsg(1012)
      txt_set.SetFocus
     Else
      If cbo_warehouse.MatchFound = False Then
       lblerror = DisplayMsg(4018)
       cbo_warehouse.SetFocus
      ElseIf cbo_location.MatchFound = False Then
       lblerror = DisplayMsg(4018)
       cbo_location.SetFocus
      ElseIf cbo_parent.MatchFound = False Then
       lblerror = DisplayMsg(4002)
       cbo_parent.SetFocus
'      ElseIf Grid.Rows < 2 Then
'       lblerror = DisplayMsg(8108) 'Please press search button
'       cmd_search.SetFocus
      Else
        If .Rows > 1 Then
         Me.MousePointer = vbHourglass
         Db_ps.BeginTrans
          If cbo_status.ListIndex = 1 Then
           
           For i = 1 To grid.Rows - 1
           
           rsseqno.Open "Select Max(Seq_No) From Part_Supply", Db_ps, adOpenForwardOnly, adLockReadOnly
           If IsNull(rsseqno(0)) Then
                VSeqNo = 1
            Else
                VSeqNo = rsseqno(0) + 1
            End If
            rsseqno.Close
           
            sql = "Insert Into Part_Supply (" & _
                  "SupplyRec_No, " & _
                  "RecSeq_No, " & _
                  "FromWarehouse_Code, " & _
                  "From_Address, " & _
                  "ToWarehouse_Code, " & _
                  "ChildSupply_Date, " & _
                  "ChildItem_Code, " & _
                  "Supply_Cls, " & _
                  "ChildRequirement_Qty, " & _
                  "Consumption_Qty, " & _
                  "ChildUnit_Cls, " & _
                  "Currency_Code, " & _
                  "Price, " & _
                  "Amount, " & _
                  "ParentItem_Code, " & _
                  "Lot_No, " & _
                  "Production_Date, " & _
                  "DO_No, " & _
                  "Remarks, " & _
                  "SJNo, "
            sql = sql + "MaterialConsump_Cls, " & _
                  "SubConPartReceipt_SeqNo, " & _
                  "Last_Update, " & _
                  "Last_User, " & _
                  "Register_Date) "
            sql = sql + "Values ('" & _
                  cbo_supply.Text & "', " & _
                  "NULL, '" & _
                  cbo_warehouse.Text & "', '', '" & _
                  cbo_location.Text & "', '" & _
                  Format(Trim(dt_supply.Value), "yyyy-MM-dd") & "', '" & _
                  Trim(.TextMatrix(i, bteColMaterialCode)) & "', '" & _
                  "S1', " & _
                  CDbl(.TextMatrix(i, bteColQtyAkhir)) & ", " & _
                  "NULL, '" & _
                  Trim(.TextMatrix(i, bteColUnitCls)) & "', " & _
                  "NULL, NULL, NULL,'" & _
                  cbo_parent.Text & "', '', NULL, " & _
                  "'', '" & Trim(txt_set.Text) & "', '', NULL, NULL, " & _
                  "getdate(), '" & _
                  userLogin & "', " & _
                  "getdate() ) "
           
            Db_ps.Execute (sql)
            
            '********Set ControlCls********
            FromControlCls = Trim(cbo_warehouse.Column(2))
            ItemControlCls = Trim(cbo_location.Column(2))
            '******Update Stock***********
            Call up_UpdateStockMaster(Format(Trim(dt_supply.Value), "yyyy-MM-dd"), _
            uf_GetLastClosing("month"), uf_GetLastClosing("year"), _
            Trim(cbo_warehouse.Text), Trim(cbo_location.Text), _
            Trim(.TextMatrix(i, bteColMaterialCode)), _
            CDbl(.TextMatrix(i, bteColQtyAkhir)), "S1", _
            Trim(.TextMatrix(i, bteColStockControlCls)), "", "I", "", "", False, False, True, Db_ps)
            
           Next i
           cbo_status.ListIndex = 0
          Else
            For i = 1 To grid.Rows - 1
             sql = "Update Part_Supply " & _
                   "Set FromWarehouse_Code = '" & cbo_warehouse.Text & "', " & _
                   "ToWarehouse_Code = '" & cbo_location.Text & "', " & _
                   "ChildSupply_Date = '" & Format(Trim(dt_supply.Value), "yyyy-MM-dd") & "', " & _
                   "ChildRequirement_Qty = " & CDbl(.TextMatrix(i, bteColQtyAkhir)) & ", " & _
                   "Last_Update = getdate(), " & _
                   "Last_User = '" & userLogin & "', " & _
                   "Register_Date = getdate() " & _
                   "where seq_no = '" & Trim(.TextMatrix(i, bteColSeqNo)) & "' "
             
             Db_ps.Execute (sql)
             
             '********Set ControlCls********
             FromControlCls = Trim(cbo_warehouse.Column(2))
             ItemControlCls = Trim(cbo_location.Column(2))
             '******Update Stock***********
             Call up_UpdateStockMaster(Format(Trim(dt_supply.Value), "yyyy-MM-dd"), _
             uf_GetLastClosing("month"), uf_GetLastClosing("year"), _
             Trim(cbo_warehouse.Text), Trim(cbo_location.Text), _
             Trim(.TextMatrix(i, bteColMaterialCode)), _
             ((CDbl(.TextMatrix(i, bteColQtyLama)) - CDbl(.TextMatrix(i, bteColQtyAkhir))) * -1), "S1", _
             Trim(.TextMatrix(i, bteColStockControlCls)), "", "U", "", "", False, False, True, Db_ps)
             Next i
           End If
           
            Db_ps.CommitTrans
            lblerror.Caption = DisplayMsg(1000) '"data saved success !"
           
           Call BrowseGrid
        End If
     End If
    End If
    
    Case 1: 'Delete
     If .Rows > 1 Then
       If cbo_status.ListIndex = 0 Then
        
        If hakUpdate(Me.Name) = 0 Then
            lblerror = DisplayMsg(3008)
            Exit Sub
        End If
        
        lblerror = up_ValidateDateRange(dt_supply.Value, True)
        If Trim(lblerror) <> "" Then Exit Sub
        
        If (MsgBox("Are you sure want to delete this Supply No ?", vbQuestion + vbYesNo, "Confirmation") = vbYes) Then
         Me.MousePointer = vbHourglass
         Db_ps.BeginTrans
          
           For i = 1 To grid.Rows - 1
           sql = "Delete from part_supply where " & _
                 "seq_no = " & CDbl(Trim(.TextMatrix(i, bteColSeqNo)))
           
           Db_ps.Execute (sql)
           
           '********Set ControlCls********
           FromControlCls = Trim(cbo_warehouse.Column(2))
           ItemControlCls = Trim(cbo_location.Column(2))
           '******Update Stock***********
           Call up_UpdateStockMaster(Format(Trim(dt_supply.Value), "yyyy-MM-dd"), _
           uf_GetLastClosing("month"), uf_GetLastClosing("year"), _
           Trim(cbo_warehouse.Text), Trim(cbo_location.Text), _
           Trim(.TextMatrix(i, bteColMaterialCode)), _
           (CDbl(.TextMatrix(i, bteColQtyLama)) * -1), "S1", _
           Trim(.TextMatrix(i, bteColStockControlCls)), "", "D", "", "", False, False, True, Db_ps)
           
           Next i
         
            Db_ps.CommitTrans
            Call IsiDefaultValue
            Call Kosong
            Call isi_cbo_supply
            lblerror.Caption = DisplayMsg(1201) '"delete success !"
       End If
      End If
     End If
    
    Case 2: 'Cancel
     Me.MousePointer = vbHourglass
     If .Rows > 1 Then Call BrowseGrid
    
    Case 3: 'Clear
     Me.MousePointer = vbHourglass
     Call IsiDefaultValue
     Call Kosong
     cbo_status.ListIndex = 0

 End Select
End With

ErrExit:
    Me.MousePointer = vbDefault
    Exit Sub

errHandler:
    Db_ps.RollbackTrans
    lblerror.Caption = "[" & err.number & "] " & err.Description
    err.clear
    Resume ErrExit

End Sub

Private Sub cmdBrowser_Click()
 If cbo_parent.Enabled = True Then
  Me.MousePointer = vbHourglass
  frm_BrowseItem.getPartNumber = cbo_parent.Text
  frm_BrowseItem.Show 1
  cbo_parent.Text = frm_BrowseItem.getPartNumber
  Me.MousePointer = vbDefault
 End If

End Sub

Private Sub dt_from_Change()
If cbo_status.Text = "Update" Then
 Call isi_cbo_supply
Else
 Call Kosong
End If
End Sub
Private Sub dt_to_Change()
If cbo_status.Text = "Update" Then
 Call isi_cbo_supply
Else
 Call Kosong
End If
End Sub

Private Sub Form_Load()
If gb_Simulation = True Then Call up_InitSimulation(Me)
CtrlMenu1.FormName = Me.Name
Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"

Db_ps.Open Db.ConnectionString
Call IsiDefaultValue

Call Kosong
Call adtocombo

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    lblerror.Caption = ErrMsg
End If
End Sub
Private Sub CmdSubMenu_Click()
DoEvents
frmMainMenu.Show
DoEvents
Unload Me
End Sub
Sub Kosong()
cbo_supply.Text = ""
cbo_warehouse.Text = ""
lbl_warehouse.Caption = ""
cbo_location.Text = ""
lbl_location.Caption = ""
cbo_parent.Text = ""
txt_set.Text = ""
lblerror.Caption = ""
Call Header
End Sub

Sub adtocombo()
'****cbo_status*****
With cbo_status
.clear
.AddItem "Update"
.AddItem "Create"
.ListIndex = 0
End With

'*******Parent Item Code**********
Dim rs_parent As New ADODB.Recordset
With cbo_parent
    .clear
    .columnCount = 3
    
    sql = "select item_code, makeritem_code, item_name from item_master " & _
        "order by item_code"
    
    If rs_parent.State <> adStateClosed Then rs_parent.Close
    rs_parent.Open sql, Db_ps, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    While Not rs_parent.EOF
        .AddItem ""
        .List(.ListCount - 1, 0) = Trim(rs_parent!Item_Code)
        .List(.ListCount - 1, 1) = Trim(rs_parent!MakerItem_Code)
        .List(.ListCount - 1, 2) = Trim(rs_parent!item_name)
        rs_parent.MoveNext
    Wend
    
    If rs_parent.State <> adStateClosed Then rs_parent.Close
    .ListWidth = 340
    .ColumnWidths = "75pt;75pt;90pt"
End With

'****************************From n To Warehouse*************************************
 Dim rs_warehouse As New ADODB.Recordset
 Dim sql_warehouse As String
 
 sql_warehouse = "select * from (select wh_code,wh_name,isnull(stockControl_cls,'02') stockControl_cls from warehouse_master " & _
                 "union all " & _
                 "select distinct(manufacture_line.manufacture_code)wh_code,trade_name wh_name,stockControl_Cls='01' from manufacture_line join trade_master on manufacture_line.manufacture_code=trade_master.trade_code)tbJ order by wh_code "
 
 Set rs_warehouse = Db_ps.Execute(sql_warehouse)
    
 'From
 With cbo_warehouse
    .clear
    .columnCount = 3
    .TextColumn = 1
    
    If rs_warehouse.EOF = False Or rs_warehouse.BOF = False Then
        rs_warehouse.MoveFirst
        While rs_warehouse.EOF = False
            .AddItem ""
            .List(.ListCount - 1, 0) = Trim(rs_warehouse!wh_code)
            .List(.ListCount - 1, 1) = Trim(rs_warehouse!WH_Name)
            .List(.ListCount - 1, 2) = Trim(rs_warehouse!stockcontrol_cls)
            rs_warehouse.MoveNext
        Wend
        .ColumnWidths = "50 pt; 175 pt; 0 pt"
        .ListWidth = 225
    End If
 End With
    
 'To
 With cbo_location
    .clear
    .columnCount = 3
    .TextColumn = 1
       
    If rs_warehouse.EOF = False Or rs_warehouse.BOF = False Then
       rs_warehouse.MoveFirst
       While rs_warehouse.EOF = False
           .AddItem ""
           .List(.ListCount - 1, 0) = Trim(rs_warehouse!wh_code)
           .List(.ListCount - 1, 1) = Trim(rs_warehouse!WH_Name)
           .List(.ListCount - 1, 2) = Trim(rs_warehouse!stockcontrol_cls)
           rs_warehouse.MoveNext
       Wend
       .ColumnWidths = "50 pt; 175 pt;0 pt"
       .ListWidth = 225
    End If
 End With

Set rs_warehouse = Nothing
End Sub

Private Sub cbo_status_Click()
cbo_parent.locked = (cbo_status.Text = "Update")
txt_set.locked = (cbo_status.Text = "Update")
        
If cbo_status.Text = "Create" Then
    Header
    Call Kosong
    cbo_supply.clear
    Call GenerateSupplyNo
    Call IsiDefaultValue
    cbo_supply.locked = True
Else
    cbo_supply.locked = False
    Call isi_cbo_supply
End If
    
End Sub
Private Sub cbo_status_Change()
Call cbo_status_Click
End Sub

Private Sub cbo_warehouse_Click()
If cbo_warehouse.MatchFound Then
  lbl_warehouse = cbo_warehouse.Column(1)
  lblerror = ""
Else
  lbl_warehouse = ""
  lblerror = DisplayMsg(4018)
End If
End Sub

Private Sub cbo_warehouse_Change()
lbl_warehouse = ""
lblerror = ""
End Sub

Private Sub cbo_warehouse_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 13 Then Call cbo_warehouse_Click
End Sub

Private Sub cbo_location_Click()
If cbo_location.MatchFound Then
  lbl_location = cbo_location.Column(1)
  lblerror = ""
Else
  lbl_location = ""
  lblerror = DisplayMsg(4018)
End If
End Sub

Private Sub cbo_location_Change()
lbl_location = ""
lblerror.Caption = ""
End Sub

Private Sub cbo_location_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 13 Then Call cbo_location_Click
End Sub
Private Sub cbo_parent_Click()
If cbo_parent.MatchFound Then
  lblerror.Caption = ""
Else
  lblerror.Caption = DisplayMsg(4002)
End If
Header
End Sub
Private Sub cbo_parent_Change()
lblerror.Caption = ""
If cbo_parent.MatchFound Then
    If cbo_parent.Column(0) = cbo_parent.Column(1) Then
        lbl_parent.Caption = cbo_parent.Column(2)
    Else
        lbl_parent.Caption = cbo_parent.Column(1) & " " & cbo_parent.Column(2)
    End If
Else
    lbl_parent.Caption = ""
End If
End Sub
Private Sub cbo_parent_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 13 Then Call cbo_parent_Click
End Sub
Private Sub cbo_supply_Click()
If cbo_supply.MatchFound Then
 Header
 Call IsiDataPartSupply
 lblerror.Caption = ""
Else
 lblerror.Caption = DisplayMsg(8093)
End If
End Sub
Private Sub cbo_supply_Change()
lblerror.Caption = ""
End Sub
Private Sub cbo_supply_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 13 Then Call cbo_supply_Click
End Sub
Sub GenerateSupplyNo()
Dim rsSupplyNo As New ADODB.Recordset
    
    sql = "Select Isnull(Max(SubString(SupplyRec_No,1,5)), 0) + 1 As SupplyNo from Part_Supply"
    
    If rsSupplyNo.State <> adStateClosed Then rsSupplyNo.Close
    rsSupplyNo.Open sql, Db_ps, adOpenKeyset, adLockOptimistic
    If rsSupplyNo.EOF Then
        cbo_supply.Text = "00001" & Format(dt_supply, "MM") & "/" & "/" & Format(dt_supply, "YYYY") & "/" & "BOM"
    Else
        cbo_supply.Text = Format(rsSupplyNo!SupplyNo, "00000") & "/" & Format(dt_supply, "MM") & "/" & Format(dt_supply, "YYYY") & "/" & "BOM"
    End If
    If rsSupplyNo.State <> adStateClosed Then rsSupplyNo.Close
End Sub
Private Sub dt_supply_Change()
If cbo_status.Text = "Create" Then
    Call GenerateSupplyNo
Else
 If cbo_supply.MatchFound Then
  dt_supply.Month = Month
  dt_supply.Year = Year
  dt_supply.Value = Format(dt_supply.Value, "dd MMM YYYY")
 End If
End If
End Sub

Private Sub cmd_search_Click()
If cbo_parent.Text = "" Then
 cbo_parent.SetFocus
 lblerror = DisplayMsg(1008) '"Please Input Parent Code"
 Me.MousePointer = vbDefault
 Exit Sub
ElseIf txt_set.Text = "" Then
 txt_set.SetFocus
 lblerror = DisplayMsg(1012) '"Please Input Quantity"
 Me.MousePointer = vbDefault
 Exit Sub
Else
 Call BrowseGrid
End If
End Sub

Sub BrowseGrid()
Dim rs_grid As New ADODB.Recordset
Dim sql_grid As String
Dim i As Long
Dim Qty As Double

Call Header

If cbo_status.Text = "Update" Then
 sql_grid = "select ps.seq_no, ps.childitem_code item_code, im.item_name, ps.childrequirement_qty qty_akhir, isnull(bm.qty,0) qty, ps.childunit_cls unit_cls, uc.description, isnull(im.stockcontrol_cls,'02') stockcontrol_cls " & _
            "from part_supply ps " & _
            "left join bom_master bm " & _
            "on ps.childitem_code = bm.item_code " & _
            "and ps.ParentItem_Code = bm.Parent_ItemCode " & _
            "left join Item_Master im " & _
            "on ps.ChildItem_Code = im.Item_Code " & _
            "left join unit_cls uc " & _
            "on ps.childunit_cls= uc.unit_cls " & _
            "where ps.supplyrec_no = '" & cbo_supply.Text & "' "
            
Else
 sql_grid = "select seq_no = '', bm.item_code, im.item_name, isnull(bm.qty,0) qty, bm.unit_cls, uc.description, isnull(im.stockcontrol_cls,'02') stockcontrol_cls " & _
            "From " & _
            "BOM_Master bm left join Item_Master im " & _
            "On bm.item_code = im.item_code " & _
            "left join unit_cls uc " & _
            "On bm.unit_cls = uc.unit_cls " & _
            "where bm.parent_itemcode = '" & cbo_parent.Text & "' "
End If

Set rs_grid = Db_ps.Execute(sql_grid)

i = 1
With grid
    Do While Not rs_grid.EOF
        .Rows = .Rows + 1
        Qty = rs_grid!Qty * CDbl(txt_set.Text)
        .TextMatrix(i, bteColMaterialCode) = Trim(rs_grid!Item_Code)
        .TextMatrix(i, bteColDescription) = Trim(rs_grid!item_name)
        .TextMatrix(i, bteColQtyAwal) = Format(Qty, gs_formatQtyBOM)
        
        If cbo_status.Text = "Update" Then
         .TextMatrix(i, bteColQtyAkhir) = Format(rs_grid!qty_akhir, gs_formatQtyBOM)
        Else
         .TextMatrix(i, bteColQtyAkhir) = .TextMatrix(i, bteColQtyAwal)
        End If
        
        .TextMatrix(i, bteColUnit) = Trim(rs_grid!Description)
        .TextMatrix(i, bteColUnitCls) = Trim(rs_grid!Unit_cls)
        .TextMatrix(i, bteColQtyLama) = .TextMatrix(i, bteColQtyAkhir)
        .TextMatrix(i, bteColSeqNo) = Trim(rs_grid!Seq_no)
        .TextMatrix(i, bteColStockControlCls) = Trim(rs_grid!stockcontrol_cls)
        grid.Cell(flexcpBackColor, i, bteColQtyAkhir) = &HFFFFFF
        i = i + 1
        rs_grid.MoveNext
    Loop
End With

Set rs_grid = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
Db_ps.Close
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
lblerror = ""
If grid.Col = bteColQtyAkhir Then
    If Trim(grid.Text) = "" Then grid.Text = "0"
    grid.TextMatrix(Row, bteColQtyAkhir) = Format(grid.TextMatrix(Row, bteColQtyAkhir), gs_formatQtyBOM)
End If
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
grid.EditMaxLength = 18
If grid.Col = bteColQtyAkhir Then Exit Sub
Cancel = True


    
  
End Sub

Private Sub grid_Click()
If grid.Col = bteColQtyAkhir Then
    grid.FocusRect = flexFocusInset
Else
    grid.FocusRect = flexFocusNone
End If
End Sub

Private Sub Grid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyEscape Then
    KeyAscii = 0
End If
End Sub

Sub isi_cbo_supply()
Dim rs_No As New ADODB.Recordset

sql = "select distinct ps.supplyrec_no, im.makeritem_code, im.item_name " & _
    "from part_supply ps " & _
    "inner join item_master im on ps.parentitem_code = im.item_code " & _
    "where ps.childsupply_date >= '" & Format(dt_from.Value, "yyyy-MM-dd") & "' " & _
    "and ps.childsupply_date <= '" & Format(dt_to.Value, "yyyy-MM-dd") & "' " & _
    "and ps.supplyrec_no like '%BOM' " & _
    "order by ps.supplyrec_no"
If rs_No.State <> adStateClosed Then rs_No.Close
rs_No.Open sql, Db_ps, adOpenKeyset, adLockOptimistic
    cbo_supply.clear
    cbo_supply.columnCount = 3
    If Not rs_No.EOF Then
        While Not rs_No.EOF
            cbo_supply.AddItem ""
            cbo_supply.List(cbo_supply.ListCount - 1, 0) = Trim(rs_No!supplyRec_No)
            cbo_supply.List(cbo_supply.ListCount - 1, 1) = Trim(rs_No!MakerItem_Code)
            cbo_supply.List(cbo_supply.ListCount - 1, 2) = Trim(rs_No!item_name)
            rs_No.MoveNext
        Wend
    End If
    cbo_supply.ListWidth = 375
    cbo_supply.ColumnWidths = "100pt;75pt;200pt"
If rs_No.State <> adStateClosed Then rs_No.Close
End Sub

Sub IsiDefaultValue()
    dt_from = Format(Now, "dd MMM YYYY")
    dt_to = Format(Now, "dd MMM YYYY")
    dt_supply = Format(Now, "dd MMM YYYY")
End Sub

Sub IsiDataPartSupply()
Dim rs_isi As New ADODB.Recordset

sql = "select top 1 fromwarehouse_code, towarehouse_code, childsupply_date, parentitem_code, isnull(remarks,1) qty " & _
      "from part_supply where supplyrec_no = '" & cbo_supply.Text & "'"
Set rs_isi = Db_ps.Execute(sql)

If Not rs_isi.EOF Then
 cbo_warehouse.Text = Trim(rs_isi!FromWarehouse_Code)
 cbo_location.Text = Trim(rs_isi!towarehouse_code)
 dt_supply.Value = Format(Trim(rs_isi!childsupply_date), "dd MMM YYYY")
 Month = dt_supply.Month
 Year = dt_supply.Year
 cbo_parent.Text = Trim(rs_isi!parentItem_code)
 txt_set.Text = Format(Trim(rs_isi!Qty), gs_formatQty)
End If

Set rs_isi = Nothing
End Sub

Private Sub txt_set_Change()
 If InStr(1, txt_set.Text, ",") = 1 Then txt_set.Text = Right(txt_set, Len(txt_set) - 1)
End Sub
Private Sub txt_set_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then
   KeyAscii = 0
End If

If KeyAscii = Asc(".") Then KeyAscii = 0

If Val(txt_set.Text) > gd_MaxQty And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub txt_set_LostFocus()
 txt_set.Text = Format(txt_set.Text, gs_formatQty)
 If txt_set.Text = "" Then
  grid.Rows = 1
 Else
  If grid.Rows > 1 Then Call BrowseGrid
 End If
End Sub
