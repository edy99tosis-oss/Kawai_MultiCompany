VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPartMaterialSupplyByBom 
   BackColor       =   &H00FDDFE3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form Part (Material) Supply [By BOM]"
   ClientHeight    =   10230
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15120
   Icon            =   "frmPartMaterialSupplyByBom.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   10230
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H0080FFFF&
      Caption         =   "Delete"
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
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   9630
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   2745
      Left            =   240
      TabIndex        =   4
      Top             =   690
      Width           =   14655
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
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Format          =   141230083
         CurrentDate     =   39287
      End
      Begin MSComCtl2.DTPicker dt_from 
         Height          =   330
         Left            =   2430
         TabIndex        =   7
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
         Format          =   141230083
         CurrentDate     =   39105
      End
      Begin MSComCtl2.DTPicker dt_to 
         Height          =   330
         Left            =   4440
         TabIndex        =   8
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
         Format          =   141230083
         CurrentDate     =   39289
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remaining"
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
         Left            =   5100
         TabIndex        =   25
         Top             =   2280
         Width           =   900
      End
      Begin MSForms.ComboBox ComboBox1 
         Height          =   330
         Left            =   6195
         TabIndex        =   24
         Top             =   2220
         Width           =   870
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "1535;582"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         Value           =   "cbo_location"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
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
         TabIndex        =   21
         Top             =   1305
         Width           =   1785
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
         TabIndex        =   20
         Top             =   1785
         Width           =   1305
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
         TabIndex        =   19
         Top             =   2340
         Width           =   1050
      End
      Begin MSForms.ComboBox cbo_warehouse 
         Height          =   330
         Left            =   2430
         TabIndex        =   18
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
      Begin MSForms.ComboBox cbo_location 
         Height          =   330
         Left            =   2430
         TabIndex        =   17
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
         TabIndex        =   16
         Top             =   1305
         Width           =   3210
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
         TabIndex        =   15
         Top             =   1785
         Width           =   2490
      End
      Begin VB.Line Line1 
         X1              =   4050
         X2              =   7110
         Y1              =   1530
         Y2              =   1530
      End
      Begin VB.Line Line2 
         X1              =   4050
         X2              =   7110
         Y1              =   2040
         Y2              =   2040
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
         TabIndex        =   14
         Top             =   300
         Width           =   1545
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
         TabIndex        =   13
         Top             =   300
         Width           =   210
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
         TabIndex        =   12
         Top             =   840
         Width           =   870
      End
      Begin MSForms.ComboBox cbo_supply 
         Height          =   315
         Left            =   2430
         TabIndex        =   11
         Top             =   780
         Width           =   2670
         VariousPropertyBits=   612386843
         MaxLength       =   25
         DisplayStyle    =   3
         Size            =   "4710;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cbo_status 
         Height          =   315
         Left            =   180
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   1785
         Width           =   60
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9630
      Width           =   1200
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H0080FFFF&
      Caption         =   "To Material &Supply"
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
      Left            =   12825
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9630
      Width           =   2010
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   615
      Left            =   210
      TabIndex        =   0
      Top             =   8700
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
         TabIndex        =   1
         Top             =   210
         Width           =   14490
      End
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   420
      Left            =   13035
      TabIndex        =   22
      Top             =   120
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   741
   End
   Begin VSFlex8Ctl.VSFlexGrid grid 
      Height          =   5070
      Left            =   210
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   3510
      Width           =   14715
      _cx             =   25956
      _cy             =   8943
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
      Left            =   60
      TabIndex        =   23
      Top             =   23
      Width           =   14505
   End
End
Attribute VB_Name = "frmPartMaterialSupplyByBom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ColCustCd, ColCustDesc, ColNoPo, ColDatePO, ColProdCode, colprodname, ColQty, ColSetQty, ColRemaining, ColUnit, ColSetQtyOri
Dim sql As String
Dim Month, Year
Dim Db_ps As New ADODB.Connection
Dim rsRpt As New ADODB.Recordset
Sub Header()
    
    ColCustCd = 1
    ColCustDesc = 2
    ColNoPo = 3
    ColDatePO = 4
     ColProdCode = 5
    colprodname = 6
    ColQty = 7
    ColSetQty = 8
    ColRemaining = 9
    ColUnit = 10
    ColSetQtyOri = 11
                
    With grid
        .clear
        .Rows = 1
        .ColS = 12
        
        .TextMatrix(0, 0) = "S"
        .TextMatrix(0, ColCustCd) = "Cust CD"
        .TextMatrix(0, ColCustDesc) = "Description"
        .TextMatrix(0, ColNoPo) = "No Po"
        .TextMatrix(0, ColDatePO) = "Date PO"
        .TextMatrix(0, ColProdCode) = "Product CD"
        .TextMatrix(0, colprodname) = "Name"
        .TextMatrix(0, ColQty) = "Qty"
        .TextMatrix(0, ColSetQty) = "Set Qty"
        .TextMatrix(0, ColRemaining) = "Remaining"
        .TextMatrix(0, ColUnit) = "Unit"
        .TextMatrix(0, ColSetQtyOri) = "Set Qty Original" 'Hide
                        
        .ColWidth(0) = 250
        .ColWidth(ColCustCd) = 900
        .ColWidth(ColCustDesc) = 2600
        .ColWidth(ColNoPo) = 1700
        .ColWidth(ColDatePO) = 1200
        .ColWidth(ColProdCode) = 1100
        .ColWidth(colprodname) = 3000
        .ColWidth(ColQty) = 950
        .ColWidth(ColSetQty) = 950
        .ColWidth(ColRemaining) = 1150
        .ColWidth(ColUnit) = 700
        .ColWidth(ColSetQtyOri) = 0 'hide
                
        .Cell(flexcpAlignment, 0, 0, 0, grid.ColS - 1) = flexAlignCenterCenter
        .ColAlignment(ColCustCd) = flexAlignLeftCenter
        .ColAlignment(ColCustDesc) = flexAlignLeftCenter
        .ColAlignment(ColNoPo) = flexAlignLeftCenter
        .ColAlignment(ColProdCode) = flexAlignLeftCenter
        '.Editable = flexEDKbdMouse
        'EditMaxLength = 1
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
              
Private Sub cmdBrowser_Click()
 

End Sub

Private Sub cmdAction_Click(Index As Integer)
On Error GoTo xxx
lblerror = ""
If cbo_supply = "" Then
    lblerror = DisplayMsg(8107)
    cbo_supply.SetFocus
    Exit Sub
ElseIf cbo_warehouse = "" Then
    lblerror = DisplayMsg(31)
    cbo_warehouse.SetFocus
    Exit Sub

ElseIf cbo_location = "" Then
    lblerror = DisplayMsg(31)
    cbo_location.SetFocus
    Exit Sub
ElseIf cbo_location = cbo_warehouse Then
    lblerror = DisplayMsg(4053)
    cbo_location.SetFocus
    Exit Sub
ElseIf grid.Rows <= 1 Then
    lblerror = DisplayMsg(8011)
    Exit Sub
End If
'pengecekan data
Dim GaLengkap As Boolean
GaLengkap = True
For i = 1 To grid.Rows - 1
 If grid.Cell(flexcpChecked, i, 0) = flexChecked Then
    GaLengkap = False
 End If
Next
If GaLengkap Then
    lblerror = DisplayMsg(8011)
    Exit Sub
End If


For i = 1 To grid.Rows - 1
 If grid.Cell(flexcpChecked, i, 0) = flexChecked And CDbl(grid.TextMatrix(i, ColSetQty) = 0) Then
    GaLengkap = True
 End If
Next
If GaLengkap Then
    lblerror = DisplayMsg(1012)
    Exit Sub
End If


Dim StrLemparQuery As String
Dim HitUnion As Integer
Dim TempSetQTY As Double
Dim TGLPO
StrLemparQuery = ""
HitUnion = 0
TempSetQTY = 0
For i = 1 To grid.Rows - 1
    If grid.Cell(flexcpChecked, i, 0) = flexChecked Then
    HitUnion = HitUnion + 1
    If HitUnion > 1 Then
        StrLemparQuery = StrLemparQuery & " Union ALL "
    End If
    StrLemparQuery = StrLemparQuery & " select item_code,(SELECT Item_Name FROM Item_Master WHERE Item_Code=Bm.Item_Code)ItemDesc,sum(qty)Jml,JmlRev=sum(qty)* " & CDbl(grid.TextMatrix(i, ColSetQty)) & ",(Select Description FROM Unit_Cls where Unit_Cls=Bm.Unit_cLS)Unit_Desc " & _
                    " ,Unit_Cls FROM bom_Master BM where Parent_ItemCode IN ( '" & grid.TextMatrix(i, ColProdCode) & "') " & _
                    " group by Item_Code,Unit_Cls "
    TempSetQTY = CDbl(grid.TextMatrix(i, ColSetQty))
    End If
    TGLPO = grid.TextMatrix(i, ColDatePO)
Next i
 StrLemparQuery = " sELECT Item_Code , ItemDesc,Sum(JML)TotQty_BOM , sUM(JmlRev) TotQty_Supply,Unit_Cls FROM ( " & StrLemparQuery & " )ccc  group by Item_Code,ItemDesc,Unit_Cls"

'MsgBox "" & StrLemparQuery


MousePointer = vbHourglass

Load frmPartMaterialSuplyBomDetail
    With frmPartMaterialSuplyBomDetail
    .Header
    .cbo_supply = cbo_supply
    .cbo_warehouse = cbo_warehouse
    .cbo_location = cbo_location
    .dt_supply = dt_supply
    .lbl_warehouse = lbl_warehouse
    .lbl_location = lbl_location
    .cbo_status = cbo_status
    DoEvents
   .PODate = TGLPO
    .txt_set = Format(TempSetQTY, gs_formatAmountIDR)
    
    Dim JmlSetQty As Double
    Dim JmlSetQtyOri As Double
    
    JmlSetQty = 0
    JmlSetQtyOri = 0
    
    For i = 1 To grid.Rows - 1
        If grid.Cell(flexcpChecked, i, 0) = flexChecked Then
            JmlSetQty = JmlSetQty + CDbl(grid.TextMatrix(i, ColSetQty))
            JmlSetQtyOri = JmlSetQtyOri + CDbl(grid.TextMatrix(i, ColSetQtyOri))
        End If
    Next i
    
    If UCase(Trim(cbo_status)) = "UPDATE" And JmlSetQty = JmlSetQtyOri Then
    .BrowseGrid
    Else
    .BrowseGrid (StrLemparQuery)
    End If
    '.txt_set = Format(TempSetQTY, gs_formatAmountIDR)
    .txt_set.Enabled = False
    .cmdsubmenu.Caption = "&Back"
    .Show 1
    End With
MousePointer = vbDefault
Exit Sub
xxx:
lblerror = err.number & ":" & err.Description
MousePointer = vbDefault

End Sub

Private Sub cmdDelete_Click(Index As Integer)
Dim sql As String
Dim RS As New Recordset
Dim db1 As New Connection

lblerror.Caption = ""

'If hakUpdate(Me.Name) = 0 Then lblerror = DisplayMsg(3008): Exit Sub
    
    lblerror = up_ValidateDateRange(dt_supply.Value, True)
    If lblerror.Caption <> "" Then Exit Sub
    
    '#Get Last Closing Info
    Dim ls_ClosingMonth As String
    Dim ls_ClosingYear As String
    ls_ClosingMonth = uf_GetLastClosing("month")
    ls_ClosingYear = uf_GetLastClosing("year")
    
    '#Validate date Range
    lblerror.Caption = up_ValidateDateRange(dt_supply.Value, True)
    If lblerror <> "" Then Exit Sub
    
    If (MsgBox("Are you sure want to delete?", vbQuestion + vbDefaultButton2 + vbYesNo, "Confirmation") = vbYes) Then
        sql = "Select * from SupplyBOM_Master WHERE Supply_no = '" & Trim(cbo_supply.Text) & "'"
        If RS.State = 1 Then RS.Close
        RS.Open sql, Db, adOpenKeyset, adLockOptimistic
        If RS.EOF Then
            lblerror.Caption = DisplayMsg(4024)
            Exit Sub
        End If
 
MousePointer = vbHourglass
 
On Error GoTo errHandler
        
        db1.ConnectionString = Db.ConnectionString
        db1.Open
        db1.BeginTrans
        
        
        sql = "Delete SupplyBOM_Master where Supply_no = '" & Trim(cbo_supply.Text) & "'" & vbCrLf & _
              "Delete SupplyBOM_Detail where Supply_no = '" & Trim(cbo_supply.Text) & "'"
        If RS.State = 1 Then RS.Close
        RS.Open sql, db1, adOpenKeyset, adLockOptimistic
        
        sql = "select * from part_supply where SupplyRec_no = '" & Trim(cbo_supply.Text) & "'"
        If RS.State = 1 Then RS.Close
        RS.Open sql, db1, adOpenKeyset, adLockOptimistic
        
        If Not RS.EOF Then
            Do While Not RS.EOF
                FromControlCls = "01"
                Call up_UpdateStockMaster(Format(Trim(dt_supply.Value), "yyyy-MM-dd"), _
                    uf_GetLastClosing("month"), uf_GetLastClosing("year"), _
                    Trim(cbo_warehouse.Text), Trim(cbo_location.Text), _
                    Trim(RS!childitem_code), _
                    0 - (RS!ChildRequirement_qty), "S1", _
                    "01", "", "D", "", "", False, False, True, db1)
                FromControlCls = ""
                    
                RS.MoveNext
            Loop
            
            sql = "Delete Part_supply Where SupplyRec_no = '" & Trim(cbo_supply.Text) & "'"
            If RS.State = 1 Then RS.Close
            RS.Open sql, db1, adOpenKeyset, adLockOptimistic
            
        End If
        
        db1.CommitTrans
        
        lblerror.Caption = DisplayMsg(1000)
        
    End If
    
ExitError:
Set RS = Nothing
db1.Close

BrowseGrid
MousePointer = vbDefault

Exit Sub
    
errHandler:
    lblerror.Caption = "[ " & err.number & " ] " & err.Description
    db1.RollbackTrans
    Resume ExitError
    
   
End Sub

Private Sub dt_from_Change()
If cbo_status.Text = "Update" Then
 Call isi_cbo_supply
Else
 'Call kosong
End If
End Sub
Private Sub dt_to_Change()
If cbo_status.Text = "Update" Then
 Call isi_cbo_supply
Else
 'Call kosong
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
Unload frmPartMaterialSuplyBomDetail
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

Sub adtocombo()
'****cbo_status*****
With cbo_status
.clear
.AddItem "Update"
.AddItem "Create"
.ListIndex = 0
End With

'*******Parent Item Code**********
'Dim rs_parent As New ADODB.Recordset
'With cbo_parent
'    .clear
'    .ColumnCount = 3
'
'    Sql = "select item_code, makeritem_code, item_name from item_master " & _
'        "order by item_code"
'
'    If rs_parent.State <> adStateClosed Then rs_parent.Close
'    rs_parent.Open Sql, Db_ps, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'    While Not rs_parent.EOF
'        .AddItem ""
'        .List(.ListCount - 1, 0) = Trim(rs_parent!Item_Code)
'        .List(.ListCount - 1, 1) = Trim(rs_parent!makeritem_code)
'        .List(.ListCount - 1, 2) = Trim(rs_parent!item_name)
'        rs_parent.MoveNext
'    Wend
'
'    If rs_parent.State <> adStateClosed Then rs_parent.Close
'    .ListWidth = 340
'    .ColumnWidths = "75pt;75pt;90pt"
'End With

'****************************From n To Warehouse*************************************
 Dim rs_warehouse As New ADODB.Recordset
 Dim sql_warehouse As String
 
 sql_warehouse = "select wh_code,wh_name,isnull(stockControl_cls,'02') stockControl_cls from warehouse_master where adm_group in (select Trade_code FROM trade_master where trade_Cls='3') "
 
 Set rs_warehouse = Db_ps.Execute(sql_warehouse)
    
 'From
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
        .ColumnWidths = "50 pt; 175 pt; 0 pt"
        .ListWidth = 225
    End If
 End With
    
 'To
' sql_warehouse = "select * from (select wh_code,wh_name,isnull(stockControl_cls,'02') stockControl_cls from warehouse_master where adm_group in (select Trade_code FROM trade_master where trade_Cls='3') " & _
                 "union all " & _
                 "select distinct(manufacture_line.manufacture_code)wh_code,trade_name wh_name,stockControl_Cls='01' from manufacture_line join trade_master on manufacture_line.manufacture_code=trade_master.trade_code)tbJ order by wh_code "
 sql_warehouse = "select * from (select wh_code,wh_name,isnull(stockControl_cls,'02') stockControl_cls from warehouse_master where adm_group in (select Trade_code FROM trade_master where trade_Cls<>'3') " & _
                 "union all " & _
                 "select distinct(manufacture_line.manufacture_code)wh_code,trade_name wh_name,stockControl_Cls='01' from manufacture_line join trade_master on manufacture_line.manufacture_code=trade_master.trade_code)tbJ order by wh_code "
 Set rs_warehouse = Db_ps.Execute(sql_warehouse)
 
 
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
       .ColumnWidths = "50 pt; 175 pt;0 pt"
       .ListWidth = 225
    End If
 End With
 Set rs_warehouse = Nothing
 
 With ComboBox1
 .AddItem
 .List(.ListCount - 1, 0) = "YES"
 .AddItem
 .List(.ListCount - 1, 0) = "NO"
 .ListIndex = 0
 End With
End Sub

Private Sub cbo_status_Click()
'cbo_parent.locked = (cbo_status.Text = "Update")
'txt_set.locked = (cbo_status.Text = "Update")
        
If cbo_status.Text = "Create" Then
    Header
    Call Kosong
    cbo_supply.clear
    Call GenerateSupplyNo
    'Call IsiDefaultValue
    cbo_supply.locked = True
Else
    'Db.Execute "DELETE  FROM SupplyBOm_Master WHERE Supply_No NOT IN (SELECT Supply_No From SupplyBom_Detail)" ' AND Supply_no='" & cbo_supply & "'"
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

Private Sub cbo_parent_Change()

End Sub
Private Sub cbo_parent_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)

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
    
    sql = "Select Isnull(Max(SubString(Supply_No,1,5)), 0) + 1 As SupplyNo from SupplyBOM_Master"
    
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
On Error GoTo xxx
lblerror = ""
If cbo_supply = "" Then
    lblerror = DisplayMsg(8107)
    cbo_supply.SetFocus
    Exit Sub
ElseIf cbo_warehouse = "" Then
    lblerror = DisplayMsg(31)
    cbo_warehouse.SetFocus
    Exit Sub

ElseIf cbo_location = "" Then
    lblerror = DisplayMsg(31)
    cbo_location.SetFocus
    Exit Sub
ElseIf cbo_location = cbo_warehouse Then
    lblerror = DisplayMsg(4053)
    cbo_location.SetFocus
    Exit Sub

End If
MousePointer = vbHourglass
Call BrowseGrid
MousePointer = vbDefault
If grid.Rows = 1 Then
    lblerror = DisplayMsg(13)
End If
Exit Sub
xxx:
lblerror = err.number & "-" & err.Description

End Sub

Sub BrowseGrid()
Dim rs_grid As New ADODB.Recordset
Dim sql_grid As String
Dim i As Long
Dim Qty As Double

Call Header

If cbo_status.Text = "Update" Then
    sql_grid = "SELECT Supplier_Code,(SELECT Trade_Name FROM Trade_Master WHERE Trade_Code=Supplier_Code)Supplier_name,"
    sql_grid = sql_grid & " PO_NO,Coalesce(PO_Date,getdate()) PO_Date,Item_Code,(SELECT Item_Name From Item_Master WHERE Item_Code=Sd.Item_Code)Item_Name"
    sql_grid = sql_grid & ",Qty,(SELECT Description From Unit_cls WHERE Unit_Cls=SD.Unit_Cls)UnitDesc "
    sql_grid = sql_grid & ",Coalesce((select top 1 Qty FROM PurchaseORder_Detail WHERE PO_NO=sd.PO_NO and item_code=sd.item_Code),0) QtyPO,"
    sql_grid = sql_grid & "Coalesce(((select Qty FROM PurchaseORder_Detail WHERE PO_NO=sd.PO_NO and item_code=sd.item_Code )"
    sql_grid = sql_grid & " - (SELECT isnull(Sum(Qty),0) From SupplyBom_Detail WHERE Item_code=sd.item_code AND PO_NO=sd.po_no)),0) remaining"
    sql_grid = sql_grid & " FROM SupplyBom_Detail SD WHERE Supply_No='" & cbo_supply & "'"
Else
    sql_grid = " select supplier_Code,"
    sql_grid = sql_grid & vbCrLf & "(SELECT Trade_Name FROM trade_master WHERE Trade_Code=Supplier_Code)Supplier_Name,"
    sql_grid = sql_grid & vbCrLf & " Pd.PO_No,Coalesce(Po_date,getdate()) Po_Date ,Item_Code,(SELECT Item_name FROM Item_master WHERE Item_code=pd.item_code)Item_Name,Qty,"
    sql_grid = sql_grid & vbCrLf & " (SELECT Description From Unit_cls WHERE Unit_Cls=PD.Unit_Cls)UnitDesc"
    sql_grid = sql_grid & vbCrLf & " ,isnull((SELECT Sum(Qty) From SupplyBom_Detail WHERE Item_code=PD.item_code AND PO_NO=pd.po_no),0)remaining"
    sql_grid = sql_grid & vbCrLf & " FROM PurchaseOrder_Master PM"
    sql_grid = sql_grid & vbCrLf & " INNER JOIN PurchaseOrder_Detail PD ON PM.PO_No=pd.po_no"
    sql_grid = sql_grid & vbCrLf & " WHERE PO_Date Between '" & dt_from & "' AND '" & dt_to & "'"
    sql_grid = sql_grid & vbCrLf & " and whto in (" & _
    "       select code from (select trade_code code from Trade_master union select wh_code code from warehouse_master) a "
    sql_grid = sql_grid & vbCrLf & " ) and Supplier_Code IN (SELECT Trade_Code FROM Trade_Master WHERE Trade_Cls='3') AND isnull(FIx_Cls,0)=1 "
    sql_grid = sql_grid & vbCrLf & " AND Item_Code  IN (SELECT Parent_ItemCode  FROM BOM_Master) " 'untuk mengecek apakah item di po tersebut punya anak atau tidak
    sql_grid = sql_grid & vbCrLf & " AND item_Code in "
    sql_grid = sql_grid & vbCrLf & " (Select Item_Code From Item_master wHERE Supplier_Code"
    sql_grid = sql_grid & vbCrLf & "in (select Adm_Group FROM Warehouse_Master where wh_Code='" & cbo_location & "'))"
    sql_grid = sql_grid & " Order BY pd.Po_No,Supplier_Code"
End If

Set rs_grid = Db_ps.Execute(sql_grid)

i = 1
With grid
    Do While Not rs_grid.EOF
        .Rows = .Rows + 1
        
        .Cell(flexcpChecked, i, 0) = flexUnchecked
        .TextMatrix(i, ColCustCd) = Trim(rs_grid!Supplier_Code)
        .TextMatrix(i, ColCustDesc) = Trim(IIf(IsNull(rs_grid!Supplier_Name), "", rs_grid!Supplier_Name))
        .TextMatrix(i, ColNoPo) = Trim(rs_grid!po_no)
        .TextMatrix(i, ColDatePO) = Format(rs_grid!po_date, "dd-MMM-yy")
        .TextMatrix(i, ColProdCode) = Trim(rs_grid!Item_Code)
        .TextMatrix(i, colprodname) = Trim(rs_grid!item_name)
        
        If cbo_status.Text = "Update" Then
            .TextMatrix(i, ColSetQty) = Format(Trim(rs_grid!Qty), gs_formatQty)
            .TextMatrix(i, ColQty) = Format(rs_grid!QtyPo, gs_formatQty)
            .TextMatrix(i, ColRemaining) = Format(CDbl(rs_grid!Remaining), gs_formatQty) 'CDbl(.TextMatrix(i, ColQty)) - CDbl(.TextMatrix(i, ColSetQty)) 'Trim(rs_grid!remaining)
            .Cell(flexcpChecked, i, 0) = flexChecked
        Else
            
            .TextMatrix(i, ColQty) = Format(rs_grid!Qty, gs_formatQty)
            .TextMatrix(i, ColRemaining) = Format(CDbl(.TextMatrix(i, ColQty)) - CDbl(rs_grid!Remaining), gs_formatQty)
            .TextMatrix(i, ColSetQty) = Format(.TextMatrix(i, ColRemaining), gs_formatQty)
        End If
        If ComboBox1.Text = "YES" Then
            If CDbl(.TextMatrix(i, ColRemaining)) = 0 Then
                .RowHidden(i) = True
            Else
                .RowHidden(i) = False
            End If
        Else
          If CDbl(.TextMatrix(i, ColRemaining)) = 0 Then
                .RowHidden(i) = False
            Else
                .RowHidden(i) = True
            End If
        End If
        
        .TextMatrix(i, ColUnit) = IIf(IsNull(rs_grid!unitdesc), "", Trim(rs_grid!unitdesc))
         grid.Cell(flexcpBackColor, i, ColSetQty) = &HFFFFFF
         
         .TextMatrix(i, ColSetQtyOri) = rs_grid!Qty
      
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
If grid.Col = 0 Then
'grid.Cell(flexcpChecked, Row, 0) = flexChecked
End If
lblerror = ""

Dim i As Long
For i = 1 To grid.Rows - 1
    If grid.Cell(flexcpChecked, i, 0) = flexChecked Then
        If i <> Row Then
            lblerror.Caption = "Please Select One Record"
            grid.Cell(flexcpChecked, Row, 0) = flexUnchecked
        End If
    End If
Next i


If CDbl(grid.TextMatrix(Row, ColSetQty)) > CDbl(grid.TextMatrix(Row, ColRemaining)) Then
    lblerror = DisplayMsg(4045) & grid.TextMatrix(Row, ColRemaining)
    grid.TextMatrix(Row, ColSetQty) = 0
End If
If grid.Col = ColSetQty Then
    grid.TextMatrix(Row, Col) = Format(grid.TextMatrix(Row, Col), gs_formatAmountIDR)
End If

End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
grid.Editable = flexEDKbdMouse
If Col = 0 Then
Cancel = False

End If
If grid.Cell(flexcpChecked, Row, 0) <> flexChecked Then
    If grid.Col <> 0 Then
        Cancel = True
    End If
  Else
    If grid.Col <> 0 And grid.Col <> ColSetQty Then  'And Grid.Col <> bteColCurr And Grid.Col <> bteColPrice Then
      Cancel = True
    End If
    'If Grid.Col = bteColOrder Then orderawal = CDbl(Grid.TextMatrix(Row, bteColOrder))
  End If

End Sub

Private Sub grid_Click()

'Grid.Cell(flexcpChecked, Grid.RowSel, 0) = flexChecked

End Sub

Private Sub Grid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)

If Col <> ColSetQty Then KeyAscii = 0
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyEscape Then
    KeyAscii = 0
End If
End Sub

Sub isi_cbo_supply()
Dim rs_No As New ADODB.Recordset

sql = "select Supply_no " & _
    "from SupplyBOm_Master WHERE Supply_Date BETWEEN '" & Format(dt_from.Value, "yyyy-mm-dd") & "' AND '" & Format(dt_to.Value, "yyyy-mm-dd") & "' "
    
If rs_No.State <> adStateClosed Then rs_No.Close
'rs_No.Open sql, Db_ps, adOpenKeyset, adLockOptimistic

Set rs_No = Db_ps.Execute(sql)
    cbo_supply.clear
    cbo_supply.columnCount = 1
    If Not rs_No.EOF Then
        While Not rs_No.EOF
            cbo_supply.AddItem ""
            cbo_supply.List(cbo_supply.ListCount - 1, 0) = Trim(rs_No!Supply_no)
            rs_No.MoveNext
        Wend
    End If
    cbo_supply = ""
    cbo_supply.ListWidth = 150
    cbo_supply.ColumnWidths = "150pt"
If rs_No.State <> adStateClosed Then rs_No.Close
End Sub

Sub IsiDefaultValue()
    dt_from = Format(Now, "dd MMM YYYY")
    dt_to = Format(Now, "dd MMM YYYY")
    dt_supply = Format(Now, "dd MMM YYYY")
End Sub

Sub IsiDataPartSupply()
Dim rs_isi As New ADODB.Recordset

sql = "select fromWH_code, toWH_code, supply_date " & _
      "from SupplyBom_Master where Supply_No= '" & cbo_supply.Text & "'"
Set rs_isi = Db_ps.Execute(sql)

If Not rs_isi.EOF Then
 cbo_warehouse.Text = Trim(rs_isi!FromWH_Code)
 cbo_location.Text = Trim(rs_isi!ToWH_Code)
 dt_supply.Value = Format(Trim(rs_isi!Supply_Date), "dd MMM YYYY")
 Month = dt_supply.Month
 Year = dt_supply.Year
 Call BrowseGrid
End If

Set rs_isi = Nothing
End Sub

Private Sub txt_set_Change()
 
End Sub
Private Sub txt_set_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then
   KeyAscii = 0
End If

If KeyAscii = Asc(".") Then KeyAscii = 0


End Sub

Private Sub txt_set_LostFocus()
 
End Sub


