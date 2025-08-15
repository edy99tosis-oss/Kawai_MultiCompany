VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmPartSupplyUpload 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Part Supply Unschedule Upload"
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15120
   Icon            =   "frmPartSupplyUpload.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdTemplate 
      BackColor       =   &H0080FFFF&
      Caption         =   "Template"
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
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3000
      Width           =   1005
   End
   Begin VB.CommandButton cmdUpload 
      BackColor       =   &H00E0E0E0&
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
      Height          =   315
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3000
      Width           =   405
   End
   Begin VB.TextBox txtUpload 
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
      Left            =   240
      TabIndex        =   20
      Top             =   3000
      Width           =   7275
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   585
      Left            =   255
      TabIndex        =   17
      Top             =   9000
      Width           =   14595
      Begin VB.Label lbl_pesan 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "lbl_pesan"
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
         Left            =   90
         TabIndex        =   18
         Top             =   240
         Width           =   14235
      End
   End
   Begin VB.CommandButton cmd_submit 
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
      Left            =   13710
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   9840
      Width           =   1125
   End
   Begin VB.CommandButton cmd_sub_menu 
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
      Left            =   255
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   9810
      Width           =   1125
   End
   Begin VB.CommandButton cmd_clear 
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
      Left            =   12510
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   9810
      Width           =   1125
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   12930
      TabIndex        =   12
      Top             =   120
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   714
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid1 
      Height          =   5475
      Left            =   240
      TabIndex        =   19
      Top             =   3360
      Width           =   14610
      _cx             =   25770
      _cy             =   9657
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
   Begin MSComDlg.CommonDialog cdg 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1950
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   14595
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   330
         Left            =   2340
         TabIndex        =   2
         Top             =   1485
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
         Format          =   115474435
         CurrentDate     =   37867
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000FFFF&
         FillColor       =   &H0000FFFF&
         Height          =   300
         Left            =   10920
         Shape           =   5  'Rounded Square
         Top             =   1065
         Width           =   300
      End
      Begin VB.Label lblStockQty 
         AutoSize        =   -1  'True
         BackColor       =   &H00FDDFE3&
         BackStyle       =   0  'Transparent
         Caption         =   "Validasi Stock Qty (0)"
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
         Left            =   11400
         TabIndex        =   26
         Top             =   1095
         Width           =   1875
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         FillColor       =   &H00808080&
         Height          =   300
         Left            =   10920
         Shape           =   5  'Rounded Square
         Top             =   667
         Width           =   300
      End
      Begin VB.Label lblQtyFormat 
         AutoSize        =   -1  'True
         BackColor       =   &H00FDDFE3&
         BackStyle       =   0  'Transparent
         Caption         =   "Invalid Qty Format (0)"
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
         Left            =   11400
         TabIndex        =   25
         Top             =   700
         Width           =   1920
      End
      Begin VB.Shape shapeSkip 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000C000&
         FillColor       =   &H0000C000&
         Height          =   300
         Left            =   10920
         Shape           =   5  'Rounded Square
         Top             =   225
         Width           =   300
      End
      Begin VB.Label lblProdCode 
         AutoSize        =   -1  'True
         BackColor       =   &H00FDDFE3&
         BackStyle       =   0  'Transparent
         Caption         =   "Invalid Product Code (0)"
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
         Left            =   11400
         TabIndex        =   24
         Top             =   280
         Width           =   2115
      End
      Begin VB.Label lbl_supply 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lbl_supply"
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
         Left            =   3930
         TabIndex        =   11
         Top             =   1125
         Width           =   855
      End
      Begin VB.Line Line3 
         X1              =   3930
         X2              =   6990
         Y1              =   1395
         Y2              =   1395
      End
      Begin VB.Line Line2 
         X1              =   3930
         X2              =   6990
         Y1              =   990
         Y2              =   990
      End
      Begin VB.Line Line1 
         X1              =   3930
         X2              =   6990
         Y1              =   540
         Y2              =   540
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
         Left            =   3930
         TabIndex        =   10
         Top             =   735
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
         Left            =   3930
         TabIndex        =   9
         Top             =   315
         Width           =   3210
      End
      Begin MSForms.ComboBox cbo_supply 
         Height          =   330
         Left            =   2340
         TabIndex        =   8
         Top             =   1065
         Width           =   780
         VariousPropertyBits=   746604571
         MaxLength       =   2
         DisplayStyle    =   3
         Size            =   "1376;582"
         ShowDropButtonWhen=   2
         Value           =   "cbo_supply"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cbo_location 
         Height          =   330
         Left            =   2340
         TabIndex        =   7
         Top             =   667
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
         Left            =   2340
         TabIndex        =   1
         Top             =   225
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
         TabIndex        =   6
         Top             =   1553
         Width           =   1050
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supply Cls"
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
         TabIndex        =   5
         Top             =   1133
         Width           =   915
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
         TabIndex        =   4
         Top             =   735
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
         TabIndex        =   3
         Top             =   293
         Width           =   1785
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Upload File"
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
      Left            =   7680
      TabIndex        =   22
      Top             =   3045
      Width           =   930
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Part/Material Supply Unscheduled Upload"
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
      Left            =   270
      TabIndex        =   13
      Top             =   360
      Width           =   14505
   End
End
Attribute VB_Name = "FrmPartSupplyUpload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim db2 As New ADODB.Connection
Dim rs_part_supply As New ADODB.Recordset
Dim rs_warehouse As New ADODB.Recordset
Dim rs_trade_master As New ADODB.Recordset
Dim rs_item As New ADODB.Recordset
Dim l_update_stock As Double
Dim l_tambah_stock As Double
Dim l_item_code As String, l_supply_cls As String, l_stock_warehouse As String
Dim stockcontrol_cls   As String, l_stock_location As String, l_seqNo As Double
Dim Status As String

Dim fromWHCode As String, FromAddres As String, toWHCode As String, SupplyDate As Date, ProductionCode As String
Dim SupplyCls As String, Qty As Double, UnitCls As String, Curr As String, Price As Double, Amount As Double, CurrStock As Double
Dim SuratJalan As String, bcType As String, bcNo As String, bcDate As Date, Remarks As String, LastUser As String, RegisterDate As Date
Dim l_prod_code As Double, l_qty_format As Double, l_stock_qty As Double

Dim rsDB As New ADODB.Recordset
Dim rs_Unit As New ADODB.Recordset
Dim rs_PR As New ADODB.Recordset
Dim rs_Curr As New ADODB.Recordset
                            
Dim bteColProdCode As Byte
Dim bteColPartNo As Byte
Dim bteColDesc As Byte
Dim bteColCurr As Byte
Dim bteColQty As Byte
Dim bteColUnitCls As Byte
Dim bteColCurrency As Byte
Dim bteColPrice As Byte
Dim bteColAmount As Byte
Dim bteColSJNo As Byte
Dim bteColBctype As Byte
Dim bteColBcNo As Byte
Dim bteColBCDate As Byte
Dim bteColRemark As Byte
Dim bteColNote As Byte

Dim bteHakPrice As Byte
Dim ls_ReplacementWarehouseCode As String
Dim ls_FromWarehouseCode As String
Dim ls_ToWarehouseCode As String
Dim ls_SupplySeqNo As Double
Dim ls_SupplyDate As String
Dim ls_SupplyCls As String
Dim ls_PathExcel As String

'====================================================================================================================================================================
' 1. Functions (Start)
'====================================================================================================================================================================

Private Function uf_ValidateInput() As Boolean
    If cbo_location = "" Then
        cbo_location.SetFocus
        lbl_pesan = "Please Select Location Code To!"
        uf_ValidateInput = False
        Exit Function
    ElseIf cbo_warehouse = "" Then
        cbo_warehouse.SetFocus
        lbl_pesan = "Please Select Warehouse Code From!"
        uf_ValidateInput = False
        Exit Function
    End If
    uf_ValidateInput = True
End Function

Sub Header()
    
    With Grid1
    
        bteColProdCode = 0
        bteColPartNo = 1
        bteColDesc = 2
        bteColCurr = 3
        bteColQty = 4
        bteColUnitCls = 5
        bteColCurrency = 6
        bteColPrice = 7
        bteColAmount = 8
        bteColSJNo = 9
        bteColBctype = 10
        bteColBcNo = 11
        bteColBCDate = 12
        bteColRemark = 13
        bteColNote = 14
       
        .clear
        .ColS = 15
        .Rows = 1
        
        .TextMatrix(0, bteColProdCode) = "Product Code"
        .TextMatrix(0, bteColPartNo) = "Part Number"
        .TextMatrix(0, bteColDesc) = "Description"
        .TextMatrix(0, bteColCurr) = "Current Stock"
        .TextMatrix(0, bteColQty) = "Qty"
        .TextMatrix(0, bteColUnitCls) = "Unit Cls"
        .TextMatrix(0, bteColCurrency) = "Currency"
        .TextMatrix(0, bteColPrice) = "Price"
        .TextMatrix(0, bteColAmount) = "Amount"
        .TextMatrix(0, bteColSJNo) = "SJ No."
        .TextMatrix(0, bteColBctype) = "BC Type"
        .TextMatrix(0, bteColBcNo) = "BC No"
        .TextMatrix(0, bteColBCDate) = "BC Date"
        .TextMatrix(0, bteColRemark) = "Remarks"
        .TextMatrix(0, bteColNote) = "Remarks"
                
        .ColWidth(bteColProdCode) = 1500
        .ColWidth(bteColPartNo) = 1500
        .ColWidth(bteColDesc) = 2250
        .ColWidth(bteColCurr) = 1300
        .ColWidth(bteColQty) = 1100
        .ColWidth(bteColUnitCls) = 1100
        .ColWidth(bteColCurrency) = 1100
        .ColWidth(bteColPrice) = 1500
        .ColWidth(bteColAmount) = 1850
        .ColWidth(bteColSJNo) = 2250
        .ColWidth(bteColBctype) = 1000
        .ColWidth(bteColBcNo) = 1000
        .ColWidth(bteColBCDate) = 1300
        .ColWidth(bteColRemark) = 2000
        .ColWidth(bteColNote) = 4500
        
        .Cell(flexcpAlignment, 0, 0, 0, bteColNote) = flexAlignCenterCenter

        .EditMaxLength = 1
    
    End With

End Sub

Private Sub up_ExportOffline()
Dim objExcel As New Excel.application
Dim RS As New ADODB.Recordset
Dim strSQL As String
Dim i As Double
    
    If G_CekExcelApp = False Then lbl_pesan.Caption = "Excel Application is not found !": Exit Sub
    
    lbl_pesan.Caption = ""
    cdg.filter = "Excel Worksheets 2003 (*.xls)|*.xls|"
    cdg.filename = "Upload Part Material Supply Unschedule"
    cdg.CancelError = True
    
    On Error GoTo errCancel
    cdg.ShowSave
    
   On Error GoTo errHandler
    If Len(cdg.filename) = 0 Then Exit Sub
    If Dir(cdg.filename) <> "" Then
        If MsgBox("Overwrite existing file?", vbExclamation + vbYesNo, "Overwrite") = vbNo Then Exit Sub
    End If
    ls_PathExcel = Mid(cdg.filename, 1, Len(cdg.filename) - Len(cdg.FileTitle))
    
    MousePointer = MousePointerConstants.vbHourglass
    
    Set objExcel = New Excel.application
    With objExcel
        .Workbooks.Add
        .Visible = True
        .Cells.Select
        .Cells.EntireColumn.delete

        .Range("A1").EntireColumn.delete xlDown
        .Range("A1:G1").Borders.Weight = xlThin
        .Range("A2:G2").Borders.Weight = xlThin
        .Selection.Interior.Pattern = xlNone
        
        .Range("A1:G1").Select
        .Selection.Font.Bold = True
        
        .Range("A2:G2").Select
        .Selection.Font.color = &HFF00&
        
        .Range("A1").Select
        .Selection.MergeCells = True
        .Selection.horizontalAlignment = xlCenter
        .Selection.verticalAlignment = xlCenter
        .Range("A1").Value = "Product Code"


        .Range("B1").Select
        .Selection.MergeCells = True
        .Selection.horizontalAlignment = xlCenter
        .Selection.verticalAlignment = xlCenter
        .Range("B1").Value = "Qty"


        .Range("C1").Select
        .Selection.MergeCells = True
        .Selection.horizontalAlignment = xlCenter
        .Selection.verticalAlignment = xlCenter
        .Range("C1").Value = "Surat Jalan No"


        .Range("D1").Select
        .Selection.MergeCells = True
        .Selection.horizontalAlignment = xlCenter
        .Selection.verticalAlignment = xlCenter
        .Range("D1").Value = "BC Type"


        .Range("E1").Select
        .Selection.MergeCells = True
        .Selection.horizontalAlignment = xlCenter
        .Selection.verticalAlignment = xlCenter
        .Range("E1").Value = "BC No"
        
        .Range("F1").Select
        .Selection.MergeCells = True
        .Selection.horizontalAlignment = xlCenter
        .Selection.verticalAlignment = xlCenter
        .Range("F1").Value = "BC Date"
        
        .Range("G1").Select
        .Selection.MergeCells = True
        .Selection.horizontalAlignment = xlCenter
        .Selection.verticalAlignment = xlCenter
        .Range("G1").Value = "Remark"
        
        .Range("A2").Value = "Char (25)"
        .Range("B2").Value = "Numeric (18, 2)"
        .Range("C2").Value = "Char (20)"
        .Range("D2").Value = "Varchar (15)"
        .Range("E2").Value = "Varchar (30)"
        .Range("F2").Value = "Date (30) (dd-MM-yyyy)"
        .Range("G2").Value = "Varchar (20)"
        

        .Cells.Select
        .Cells.EntireColumn.AutoFit
    
        .ActiveWorkbook.SaveAs filename:= _
        cdg.filename, FileFormat:= _
                               xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False _
                               , CreateBackup:=False
    End With
        
    MousePointer = MousePointerConstants.vbDefault
    Exit Sub
    
errHandler:
    If err.number <> 0 Then
        MousePointer = MousePointerConstants.vbDefault
        lbl_pesan.Caption = err.Description
        Grid1.FixedRows = 1
    End If
    If RS.State = adStateOpen Then
        RS.Close
        Set RS = Nothing
    End If
errCancel:

End Sub

Private Sub setting_grid()
    
    Dim sql_join As String, l_curr As String, l_item_name As String, L_price As String, l_price2 As String
    Dim rs_join As New ADODB.Recordset, rs_Replacement As New ADODB.Recordset
    
    Dim Bo_Replacement As Boolean
    
    Bo_Replacement = False
    Me.MousePointer = vbHourglass
    
    Header
    
    With Grid1
    DoEvents
        
        sql_join = "select * from (select part_supply.*,stockcontrol_cls,makeritem_code,item_name,sheetcoil_cls,width,length,thickness,unit_cls from (Select * From part_supply where fromwarehouse_code='" & Trim(cbo_warehouse) & "' and towarehouse_code='" & Trim(cbo_location) & "' and supply_cls='" & Trim(cbo_supply) & "' and childsupply_date ='" & Format(DTPicker3.Value, "yyyy-MM-dd") & "' and COALESCE(supplyrec_no,'')='')part_supply join item_master on part_supply.childitem_code=item_master.item_code ) xxx order by Register_Date" & vbCrLf & _
            "  " & vbCrLf & _
            "" & vbCrLf & _
            "" & vbCrLf & _
            " "
        
        
        rs_join.Open sql_join, Db, adOpenKeyset, adLockOptimistic
   DoEvents

    End With
    
    Me.MousePointer = vbDefault
    cmd_submit.Enabled = True

End Sub

Private Sub cbo_location_Change()
    If cbo_location.Text = "" Then
        lbl_location.Caption = ""
    End If
End Sub

Private Sub cbo_location_Click()
    If cbo_location.DataChanged = False Then Exit Sub
    If cbo_location.ListIndex <> -1 Then
        lbl_location.Caption = cbo_location.List(cbo_location.ListIndex, 1)
        l_stock_location = Trim(cbo_location.List(cbo_location.ListIndex, 2))
    Else
        lbl_location.Caption = ""
    End If
    lbl_pesan = validCombo
    If Trim(lbl_pesan) <> "" Then clearGrid: Exit Sub
    lbl_pesan = ""
    Call cbo_location_Change
End Sub

Private Sub cbo_location_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        lbl_pesan = ""
        cbo_location.DataChanged = False
        lbl_pesan = validCombo
        cbo_location.DataChanged = True
        If Trim(lbl_pesan) <> "" Then clearGrid: Exit Sub
        Call setting_grid
    End If
End Sub

Private Sub cbo_location_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub cbo_price_KeyPress(KeyAscii As MSForms.ReturnInteger)
     If InStr(1, "0123456789.", Chr(KeyAscii)) <= 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub cbo_supply_Change()
    lbl_pesan = ""
    If cbo_supply.Text = "" Then
        lbl_supply.Caption = ""
    End If
    lbl_pesan = validCombo
    If Trim(lbl_pesan) <> "" Then clearGrid: Exit Sub
    
    Select Case Trim(cbo_supply)
    Case "S1":
        If Trim(cbo_warehouse) = Trim(cbo_location) Then
            lbl_pesan.Caption = DisplayMsg(4053) '"Can't supply to same warehouse !"
            cmd_submit.Enabled = False
            Exit Sub
        End If
    Case "S":
    Case "L":
        If Trim(cbo_warehouse) <> Trim(cbo_location) Then
            lbl_pesan.Caption = DisplayMsg(4054) '"Can't input loss to different warehouse !"
            cmd_submit.Enabled = False
            Exit Sub
        End If
    Case "RJ":
        If Trim(cbo_warehouse) <> Trim(cbo_location) Then
            lbl_pesan.Caption = DisplayMsg(4055) '"Can't input reject to different warehouse !"
            cmd_submit.Enabled = False
            Exit Sub
        End If
    End Select
End Sub

Private Sub cbo_supply_Click()
    lbl_supply.Caption = cbo_supply.List(cbo_supply.ListIndex, 1)
    'Call setting_grid
End Sub

Private Sub cbo_supply_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        lbl_pesan = ""
        lbl_pesan = validCombo
        If Trim(lbl_pesan) <> "" Then clearGrid: Exit Sub
        Call setting_grid
    End If
End Sub

Private Sub cbo_supply_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub cbo_warehouse_Click()
    If cbo_warehouse.DataChanged = False Then Exit Sub
    lbl_warehouse.Caption = cbo_warehouse.List(cbo_warehouse.ListIndex, 1)
    l_stock_warehouse = Trim(cbo_warehouse.List(cbo_warehouse.ListIndex, 2))
    cbo_location.Text = ""
    txtUpload = ""
    lbl_pesan = ""
    Call setting_grid
End Sub

Private Sub set_item(Optional ls_SortBy As String)
    
    Dim sqlitem As String
    Dim RsItem As New Recordset
    
    sqlitem = "select item_code, makeritem_code, item_name, address ,a.unit_cls,b.Description from item_master a inner join unit_cls b on a.Unit_Cls=b.Unit_Cls where use_endday > convert(char(8), getdate(), 112)  "
    If Trim(ls_SortBy) = "" Then
        sqlitem = sqlitem & " order by item_code asc "
    Else
        sqlitem = sqlitem & " order by " & ls_SortBy & " asc "
    End If
    Set RsItem = Db.Execute(sqlitem)

End Sub

Private Sub cbo_warehouse_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    
    If KeyCode = 13 Then
        lbl_pesan = ""
        cbo_warehouse.DataChanged = False
        lbl_pesan = validCombo
        cbo_warehouse.DataChanged = True
        If Trim(lbl_pesan) <> "" Then clearGrid: Exit Sub
        Call setting_grid
        lbl_warehouse = cbo_warehouse.List(cbo_warehouse.ListIndex, 1)
    End If
End Sub

Sub clearGrid()
    Call Header
End Sub

Private Sub cbo_warehouse_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub cmd_Cancel_Click()
    lbl_pesan = ""
    Call setting_grid
End Sub

Private Sub cmd_clear_Click()
cmd_submit.Enabled = True
    DTPicker3.Value = Format(Date, "dd MMM yyyy")
    cbo_location = ""
    cbo_supply = "S1"
    cbo_warehouse = ""
    Call Header
    lbl_pesan.Caption = ""
    lbl_warehouse.Caption = ""
    lbl_location.Caption = ""
    txtUpload = ""
        
End Sub

Private Sub cmd_sub_menu_Click()
    frmMainMenu.Show
    Unload Me
End Sub

Private Sub Cmd_Submit_Click()
    Dim s As Integer, d As Integer, j  As Integer
    Dim l_curr As String, sql_del As String, l_amount As String, l_qty As String, L_price As String, l_unit_cls As String, sql_prod As String
    Dim RS As New ADODB.Recordset, ls_sql As String
    Dim X As Double
    
   If validasi = False Then cmd_submit.Enabled = True: Exit Sub
   
   Me.MousePointer = vbHourglass

    cmd_submit.Enabled = False
    If hakUpdate(Me.Name) = 0 Then lbl_pesan = DisplayMsg(3008): cmd_submit.Enabled = True: Exit Sub

    lbl_pesan = up_ValidateDateRange(DTPicker3.Value, True)
    If lbl_pesan.Caption <> "" Then cmd_submit.Enabled = True: cmd_submit.Enabled = True: Exit Sub
        
    s = 0
    d = 0

    '#Get Last Closing Info
    Dim ls_ClosingMonth As String
    Dim ls_ClosingYear As String
    ls_ClosingMonth = uf_GetLastClosing("month")
    ls_ClosingYear = uf_GetLastClosing("year")

    '#Validate date Range
    lbl_pesan = up_ValidateDateRange(DTPicker3.Value, True)
    If lbl_pesan <> "" Then cmd_submit.Enabled = True: cmd_submit.Enabled = True:   Exit Sub
    'validasi
    
    
    For X = 1 To Grid1.Rows - 1
    
    fromWHCode = Trim(cbo_warehouse.Text)
    FromAddres = "addres"
    toWHCode = Trim(cbo_location.Text)
    SupplyDate = DTPicker3.Value
    ProductionCode = Grid1.TextMatrix(X, bteColProdCode)
    SupplyCls = Trim(cbo_supply)
    Qty = Grid1.TextMatrix(X, bteColQty)
    UnitCls = Grid1.TextMatrix(X, bteColUnitCls)
    Curr = Grid1.TextMatrix(X, bteColCurrency)
    CurrStock = Grid1.TextMatrix(X, bteColCurr)
    Price = Grid1.TextMatrix(X, bteColPrice)
    Amount = Grid1.TextMatrix(X, bteColAmount)
    SuratJalan = Grid1.TextMatrix(X, bteColSJNo)
    bcType = Grid1.TextMatrix(X, bteColBctype)
    bcNo = Grid1.TextMatrix(X, bteColBcNo)
    bcDate = Grid1.TextMatrix(X, bteColBCDate)
    Remarks = Grid1.TextMatrix(X, bteColRemark)
    LastUser = userLogin

    Me.MousePointer = vbHourglass

    rs_part_supply.AddNew
    rs_part_supply!FromWarehouse_Code = fromWHCode
    rs_part_supply!from_address = FromAddres
    rs_part_supply!towarehouse_code = toWHCode
    rs_part_supply!childsupply_date = Format(Trim(DTPicker3.Value), "yyyy-MM-dd")
    rs_part_supply!childitem_code = ProductionCode
    rs_part_supply!supply_cls = SupplyCls
    rs_part_supply!ChildRequirement_qty = Qty
    rs_part_supply!Remarks = Remarks

    Dim rs_UnitCls As New ADODB.Recordset
    rs_UnitCls.Open "Select Unit_Cls From Unit_Cls where Description='" & UnitCls & "'", Db, adOpenKeyset, adLockOptimistic
    If rs_UnitCls.EOF = False Then
        l_unit_cls = Trim(rs_UnitCls!Unit_cls)
    Else
        l_unit_cls = ""
    End If
    rs_UnitCls.Close
    
    Dim rs_CurrCls As New ADODB.Recordset
    rs_CurrCls.Open "Select Curr_Cls From Curr_Cls where Description='" & Curr & "'", Db, adOpenKeyset, adLockOptimistic
    If rs_CurrCls.EOF = False Then
        l_curr = Trim(rs_CurrCls!curr_cls)
    Else
        l_curr = ""
    End If
    rs_CurrCls.Close

    '#Insert data into part Supply
    rs_part_supply!childunit_cls = Trim(l_unit_cls)
    rs_part_supply!currency_code = Trim(l_curr)
    rs_part_supply!Price = Price
    rs_part_supply!Amount = Amount
    rs_part_supply!SJNo = SuratJalan
    rs_part_supply!BC40_No = bcNo
    rs_part_supply!BC_Type = bcType
    rs_part_supply!BC40_Date = Format(bcDate, "yyyy-mm-dd")
    rs_part_supply!Remarks = IIf(Trim(Remarks) = "", Null, Trim(Remarks))
    rs_part_supply!parentItem_code = ""
    rs_part_supply!Lot_no = ""
    rs_part_supply!production_date = Null
    rs_part_supply!do_no = ""
    rs_part_supply!SupplySeq_No = IIf(Trim(ls_SupplySeqNo) = 0, Null, Trim(ls_SupplySeqNo))
    rs_part_supply!Last_Update = Now
    rs_part_supply!last_user = userLogin
    rs_part_supply.update
    
    If cbo_supply = "S1" Then
        uf_UpdateSupplyStockMaster
        uf_UpdateReceiptStockMaster
    Else
        uf_UpdateStockMaster
    End If
    
    sql = "EXEC dbo.sp_PartSupplyNoRegister_Upd '" & SupplyDate & "', '" & SuratJalan & "', '" & userLogin & "' "
    
    Db.Execute sql
    

Next X

lbl_pesan.Caption = DisplayMsg(1000)
Me.MousePointer = vbDefault

End Sub

Public Function SSeqNo()
Dim rsmax As New ADODB.Recordset
Dim strSQL As String

strSQL = "Select Max(seq_No) from part_supply"

Set rsmax = Db.Execute(strSQL)

SSeqNo = IIf(IsNull(rsmax(0)), 1, rsmax(0) + 1)

End Function

Private Sub cmdtemplate_Click()
Call Header
If uf_ValidateInput = False Then Exit Sub

Call up_ExportOffline
End Sub

Private Sub cmdUpload_Click()
Call Header
If uf_ValidateInput = False Then Exit Sub

lblQtyFormat = "Invalid Qty Format (0)"
lblProdCode = "Invalid Product Code (0)"
lblStockQty = "Validasi Stock Qty (0)"
l_prod_code = 0
l_qty_format = 0
l_stock_qty = 0
cdg.filename = ""
lbl_pesan.Caption = ""
txtUpload = ""

Call up_ImportOffline
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
    If ErrMsg = "" Then
        Unload Me
    Else
        lbl_pesan.Caption = ErrMsg
    End If
End Sub

Private Sub DTPicker3_Change()
    lbl_pesan = ""
    lbl_pesan = validCombo
    If Trim(lbl_pesan) <> "" Then clearGrid: Exit Sub
    
End Sub

Function validCombo() As String
    
    Dim j As Integer
    
    j = 0
    For i = 0 To cbo_warehouse.ListCount - 1
        If UCase(Trim(cbo_warehouse)) = UCase(Trim(cbo_warehouse.List(i, 0))) Then
            cbo_warehouse.Text = cbo_warehouse.List(i, 0)
          
                lbl_warehouse.Caption = cbo_warehouse.List(i, 1)
          
            lbl_pesan.Caption = ""
            j = 1
            Exit For
        End If
    Next
    
    If j = 0 Then
        lbl_warehouse.Caption = "": validCombo = DisplayMsg(4018) ' "Invalid warehouse code !"
        Exit Function
    End If
    
    j = 0
    For i = 0 To cbo_location.ListCount - 1
        If UCase(Trim(cbo_location)) = UCase(Trim(cbo_location.List(i, 0))) Then
            cbo_location.Text = cbo_location.List(i, 0)
            lbl_location.Caption = cbo_location.List(i, 1)
            lbl_pesan.Caption = ""
            j = 1
            Exit For
        End If
    Next
    
    If j = 0 Then
        lbl_location.Caption = "":  validCombo = DisplayMsg(4014) '"Invalid location code !"
        Exit Function
    End If
    
    j = 0
    For i = 0 To cbo_supply.ListCount - 1
        If UCase(Trim(cbo_supply)) = UCase(Trim(cbo_supply.List(i, 0))) Then
            cbo_supply.Text = cbo_supply.List(i, 0)
            lbl_pesan.Caption = ""
            j = 1
            Exit For
        End If
    Next
    
    If j = 0 Then
        validCombo = DisplayMsg(4056) '"Invalid supply clasification !"
        Exit Function
    End If

End Function

Function validasi() As Boolean
    
    Dim j As Integer
    
    validasi = True
    
    j = 0
    For i = 0 To cbo_warehouse.ListCount - 1
        If UCase(Trim(cbo_warehouse)) = UCase(Trim(cbo_warehouse.List(i, 0))) Then
            cbo_warehouse.Text = cbo_warehouse.List(i, 0)
            lbl_warehouse.Caption = cbo_warehouse.List(i, 1)
            lbl_pesan.Caption = ""
            j = 1
            Exit For
        End If
    Next
    If j = 0 Then
        lbl_warehouse.Caption = "": lbl_pesan.Caption = DisplayMsg(4018) '"Invalid warehouse code !"
        validasi = False
        cbo_warehouse.SetFocus
        Exit Function
    End If
    
    j = 0
    For i = 0 To cbo_location.ListCount - 1
        If UCase(Trim(cbo_location)) = UCase(Trim(cbo_location.List(i, 0))) Then
            cbo_location.Text = cbo_location.List(i, 0)
            lbl_location.Caption = cbo_location.List(i, 1)
            lbl_pesan.Caption = ""
            j = 1
            Exit For
        End If
    Next
    If j = 0 Then
        lbl_location.Caption = "": lbl_pesan.Caption = DisplayMsg(4014) '"Invalid location code !"
        validasi = False
        cbo_location.SetFocus
        Exit Function
    End If
    
    If Grid1.Rows < 2 Then
        lbl_location.Caption = "": lbl_pesan.Caption = DisplayMsg(4014)
        validasi = False
        cbo_location.SetFocus
        Exit Function
    End If

End Function

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"
    
    If db2.State <> adStateClosed Then db2.Close
    db2.Open Db.ConnectionString
    Call koneksi
    DTPicker3.Value = Format(Date, "dd MMM yyyy")
    lbl_warehouse.Caption = ""
    lbl_location.Caption = ""
    lbl_pesan.Caption = ""
    lbl_supply.Caption = ""
    cbo_location = ""
    cbo_supply = ""
    cbo_warehouse = ""
    bteHakPrice = hakPrice(Me.Name)
    Call setting
    Call Header
    Call set_item
    lbl_pesan.Caption = ""
    cbo_warehouse.DataChanged = True
    cbo_location.DataChanged = True
End Sub

Private Sub setting()
    cbo_warehouse.clear
    cbo_warehouse.columnCount = 3
    cbo_warehouse.TextColumn = 1
    
    i = 0
    If rs_warehouse.EOF = False Or rs_warehouse.BOF = False Then
        rs_warehouse.MoveFirst
        While rs_warehouse.EOF = False
            cbo_warehouse.AddItem ""
            cbo_warehouse.List(i, 0) = Trim(rs_warehouse!wh_code)
            cbo_warehouse.List(i, 1) = Trim(rs_warehouse!WH_Name)
            cbo_warehouse.List(i, 2) = Trim(rs_warehouse!stockcontrol_cls)
            rs_warehouse.MoveNext
            i = i + 1
        Wend
        cbo_warehouse.ColumnWidths = "50 pt; 175 pt; 0 pt"
        cbo_warehouse.ListWidth = 225
    End If
    
    'Setting Combo To Warehouse Code
    cbo_location.clear
    cbo_location.columnCount = 3
    cbo_location.TextColumn = 1
    
    i = 0
    If rs_warehouse.EOF = False Or rs_warehouse.BOF = False Then
        rs_warehouse.MoveFirst
        While rs_warehouse.EOF = False
            cbo_location.AddItem ""
            cbo_location.List(i, 0) = Trim(rs_warehouse!wh_code)
            cbo_location.List(i, 1) = Trim(rs_warehouse!WH_Name)
            cbo_location.List(i, 2) = Trim(rs_warehouse!stockcontrol_cls)
            rs_warehouse.MoveNext
            i = i + 1
        Wend
        cbo_location.ColumnWidths = "50 pt; 175 pt;0 pt"
        cbo_location.ListWidth = 225
    End If
    
    'Setting Combo Supply Cls
    cbo_supply.clear
    cbo_supply.columnCount = 2
    cbo_supply.TextColumn = 1
    cbo_supply.AddItem ""
    cbo_supply.List(0, 0) = "S1"
    cbo_supply.List(0, 1) = "Supply"
    cbo_supply.AddItem ""
    cbo_supply.List(1, 0) = "S"
    cbo_supply.List(1, 1) = "Consumption"
    cbo_supply.AddItem ""
    cbo_supply.List(2, 0) = "L"
    cbo_supply.List(2, 1) = "Loss"
    cbo_supply.AddItem ""
    cbo_supply.List(3, 0) = "RJ"
    cbo_supply.List(3, 1) = "Reject"
    cbo_supply.AddItem ""
    cbo_supply = "S1"
    

End Sub

Private Sub koneksi()
    Dim SqlW As String
    rs_part_supply.Open "select Top 1 * from part_supply", Db, adOpenKeyset, adLockOptimistic
    SqlW = " select * from (select wh_code,wh_name,stockControl_cls from warehouse_master " & _
        " union all " & _
        " select distinct(manufacture_line.manufacture_code)wh_code,trade_name wh_name,stockControl_Cls='01' from manufacture_line join trade_master on manufacture_line.manufacture_code=trade_master.trade_code)tbJ order by wh_code "
    rs_warehouse.Open SqlW, Db, adOpenKeyset, adLockOptimistic
    rs_trade_master.Open "select * from trade_master where trade_cls='1'", Db, adOpenKeyset, adLockOptimistic
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If rs_part_supply.State <> adStateClosed Then rs_part_supply.Close
    If rs_warehouse.State <> adStateClosed Then rs_warehouse.Close
    If rs_trade_master.State <> adStateClosed Then rs_trade_master.Close
End Sub

Private Sub up_ImportOffline()

    Dim adoCmd As New Command
    Dim rsCheck As New ADODB.Recordset
    
    Dim objExcel As New Excel.application
    Dim objWorkSheet As New Worksheet
    Dim objWorkBook As Workbook
    Dim i As Integer
    Dim iCol As Integer
    Dim colcount As Integer
    Dim RS As New ADODB.Recordset
    Dim strSQL As String
    Dim iGrdRow As Double
    Dim Year, Month, div, SubDiv, Block As String
    Dim HA As Double
    Dim l_unit_cls As String, l_curr_cls As String, L_price As Double
    

    If G_CekExcelApp = False Then lbl_pesan.Caption = "Excel Application is not found": Exit Sub
    
    'cdg.Filter = "Excel Files (*.xls)|*.xls"
    cdg.filter = "Excel Worksheets (*.xls)|*.xls|"
    
    On Error GoTo errCancel
    cdg.CancelError = True
    
    On Error GoTo err
    
    cdg.ShowOpen
    txtUpload.Text = cdg.filename
    If cdg.filename <> "" Then

        Me.MousePointer = vbHourglass
        Set objExcel = New Excel.application
        Set objWorkBook = objExcel.Workbooks.Open(cdg.filename)
        Set objWorkSheet = objWorkBook.Sheets("Sheet1")
        objExcel.Visible = False
        i = 3
        iGrdRow = 1
        colcount = 22
        With objWorkSheet
        
            Do While .Cells(i, 1).Value <> ""
                        Grid1.AddItem ""
                                      
                        Grid1.TextMatrix(iGrdRow, bteColProdCode) = Trim(.Cells(i, 1))
                        Grid1.TextMatrix(iGrdRow, bteColQty) = Trim(.Cells(i, 2))
                        Grid1.TextMatrix(iGrdRow, bteColSJNo) = Trim(.Cells(i, 3))
                        Grid1.TextMatrix(iGrdRow, bteColBctype) = Trim(.Cells(i, 4))
                        Grid1.TextMatrix(iGrdRow, bteColBcNo) = Trim(.Cells(i, 5))
                        Grid1.TextMatrix(iGrdRow, bteColBCDate) = Format(Trim(.Cells(i, 6)), Format("dd-MM-yyyy"))
                        Grid1.TextMatrix(iGrdRow, bteColRemark) = Trim(.Cells(i, 7))
                
                        If Trim(.Cells(i, 1)) <> "" Then
                            '**** cek Data cocok / tdk dgn Database
                            
                            l_item_code = Trim(.Cells(i, 1))
                            
                            rsDB.Open "select MakerItem_Code, Item_Name, Unit_Cls  from Item_Master where Item_Code='" & Trim(.Cells(i, 1)) & "'", Db, adOpenKeyset, adLockOptimistic
                            
                            If rsDB.EOF = True Then
                                Grid1.TextMatrix(iGrdRow, bteColNote) = "Invalid Part Number !!"
                                Grid1.Cell(flexcpBackColor, iGrdRow, bteColProdCode, iGrdRow, bteColNote) = &H8080FF
                                cmd_submit.Enabled = False
                                l_prod_code = CDbl(l_prod_code) + 1
                                lblProdCode = "Invalid Product Code (" & l_prod_code & ")"
                            Else
                                rs_Unit.Open "select Description from Unit_Cls where Unit_Cls='" & Trim(rsDB!Unit_cls) & "'", Db, adOpenKeyset, adLockOptimistic
                                If rs_Unit.EOF = False Then
                                    l_unit_cls = Trim(rs_Unit!Description)
                                Else
                                    l_unit_cls = ""
                                End If
                                
                                rs_PR.Open "select top 1 Price, Currency_Code from Part_Receipt where Item_Code='" & Trim(.Cells(i, 1)) & "' ORDER BY Receipt_Date DESC", Db, adOpenKeyset, adLockOptimistic
                                If rs_PR.EOF = False Then
                                    l_curr_cls = IIf(IsNull(rs_PR!currency_code), "", Trim(rs_PR!currency_code))
                                    L_price = IIf(IsNull(rs_PR!Price), 0, Trim(rs_PR!Price))
                                End If
                                
                                rs_Curr.Open "select Description  From Curr_Cls where Curr_Cls='" & l_curr_cls & "'", Db, adOpenKeyset, adLockOptimistic
                                If rs_Curr.EOF = False Then
                                    l_curr_cls = IIf(IsNull(rs_Curr!Description), "", Trim(rs_Curr!Description))
                                End If
                                
                                Grid1.TextMatrix(iGrdRow, bteColPartNo) = Trim(rsDB!MakerItem_Code)
                                
                                Grid1.TextMatrix(iGrdRow, bteColDesc) = Trim(rsDB!item_name)
                                
                                Grid1.TextMatrix(iGrdRow, bteColUnitCls) = Trim(rs_Unit!Description)
                                
                                Grid1.TextMatrix(iGrdRow, bteColCurr) = Format(Trim(GetCurrentStock), gs_formatQty)
                                
                                Grid1.TextMatrix(iGrdRow, bteColPrice) = Format(Trim(L_price), gs_formatAmount)
                                
                                Grid1.TextMatrix(iGrdRow, bteColCurrency) = Trim(l_curr_cls)
                                
                                Grid1.TextMatrix(iGrdRow, bteColAmount) = Format(Trim(L_price) * Trim(.Cells(i, 2)), gs_formatAmountIDR)
                                
                                rs_Curr.Close
                                rs_PR.Close
                                rs_Unit.Close
                                
                                End If
                            End If
                            rsDB.Close
                            
                        If IsNumeric(.Cells(i, 2)) = False Then
                            Grid1.TextMatrix(iGrdRow, bteColNote) = "Invalid Qty Format!!"
                            Grid1.Cell(flexcpBackColor, iGrdRow, bteColProdCode, iGrdRow, bteColNote) = &H8080FF
                            l_qty_format = CDbl(l_prod_code) + 1
                            l_qty_format = "Invalid Product Code (" & l_qty_format & ")"
                            cmd_submit.Enabled = False
                        ElseIf Len(.Cells(i, 3)) > 20 Then
                            Grid1.TextMatrix(iGrdRow, bteColNote) = "Invalid Surat Jalan No!!"
                            Grid1.Cell(flexcpBackColor, iGrdRow, bteColProdCode, iGrdRow, bteColNote) = &H8080FF
                            cmd_submit.Enabled = False
                        ElseIf Grid1.TextMatrix(iGrdRow, bteColCurr) = "99,999.00" Then
                            Grid1.TextMatrix(iGrdRow, bteColNote) = "Invalid Current Stock!!"
                            Grid1.TextMatrix(iGrdRow, bteColCurr) = "0.00"
                            Grid1.Cell(flexcpBackColor, iGrdRow, bteColProdCode, iGrdRow, bteColNote) = &H8080FF
                            l_qty_format = CDbl(l_qty_format) + 1
                            l_qty_format = "Invalid Qty Format (" & l_qty_format & ")"
                            cmd_submit.Enabled = False
                        ElseIf CDbl(Grid1.TextMatrix(iGrdRow, bteColCurr)) < CDbl(Grid1.TextMatrix(iGrdRow, bteColQty)) Then
                            Grid1.TextMatrix(iGrdRow, bteColNote) = "Qty Must Be Equal Or Lower Then Current Stock !!"
                            Grid1.Cell(flexcpBackColor, iGrdRow, bteColProdCode, iGrdRow, bteColNote) = &H8080FF
                            l_stock_qty = CDbl(l_stock_qty) + 1
                            lblStockQty = "Validasi Stock Qty (" & l_stock_qty & ")"
                            ' & Grid1.Rows - 1 & "
                            cmd_submit.Enabled = False
                        Else
                            If IsDate(.Cells(i, 6)) = False Then
                                Grid1.TextMatrix(iGrdRow, bteColNote) = "Invalid BC Date Format (dd-mm-yyyy) "
                                Grid1.Cell(flexcpBackColor, iGrdRow, bteColProdCode, iGrdRow, bteColNote) = &H8080FF
                                cmd_submit.Enabled = False
                            End If
                            
                        End If
                        
                        iGrdRow = iGrdRow + 1
                    
                    
                lbl_pesan = "Reading row : " & i - 1
                DoEvents
                i = i + 1
            Loop
            
        End With
        
        ' clean object excel
        objExcel.Workbooks.Close
        Set objWorkSheet = Nothing
        
        Set objExcel = Nothing

        'LblTotalRec = "Total : " & Grid1.Rows - 1 & " record (s)"
        lbl_pesan.Caption = "Reading Excel finish"
        
        Me.MousePointer = vbDefault
    End If
    Exit Sub
errCancel:
err:
    lbl_pesan.Caption = err.Description
    objExcel.Workbooks.Close
    Set objWorkBook = Nothing
    Set objWorkSheet = Nothing
    Set objWorkBook = Nothing
    Set objExcel = Nothing
    Me.MousePointer = vbDefault
    
End Sub

Private Function GetCurrentStock() As Double
    
    Dim adoRs As New ADODB.Recordset
    
    GetCurrentStock = 0
    
    sql = "select isnull(sum(lm_current), 0) lm_current, " & _
        "isnull(sum(tm_current), 0) tm_current, " & _
        "isnull(sum(nm_current), 0) nm_current " & _
        "from stock_master " & _
        "where item_code = '" & l_item_code & "' "
    If Not IsNull(cbo_warehouse.Text) Then sql = sql & "and warehouse_code = '" & cbo_warehouse.Text & "' "
    
    adoRs.Open sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not adoRs.EOF Then
        Select Case DateDiff("m", uf_GetLastClosing("fulldate"), DTPicker3.Value)
        Case 0: GetCurrentStock = adoRs.Fields("lm_current")
        Case 1: GetCurrentStock = adoRs.Fields("tm_current")
        Case 2: GetCurrentStock = adoRs.Fields("nm_current")
        Case Else: GetCurrentStock = "99999"
        End Select
    End If
    adoRs.Close
    
End Function

Public Function uf_GetLastClosing(Request As String) As String

    '###############################################################
    '#                                                             #
    '#  Notes : To Get Last Closing Month,Year, or full date       #
    '#                                                             #
    '###############################################################
    
    Dim sqlControl As String, RsInvControl As New ADODB.Recordset
    Dim InvYear As String
    Dim InvMonth As String
    Dim lotno As String
    
    sqlControl = "select * from inventory_control where fix_cls='1' order by inventory_year desc ,inventory_month desc"
    
    If Request = "fulldate" Then
        sqlControl = "   select    " & _
        "        cast (   " & _
        "        cast(year as varchar(4) ) +case when month <10 then '0' else'' end +cast (month as varchar(2) )+'01'    " & _
        "            as dateTime)ClosingDate     " & _
        "        from    " & _
        "        (   " & _
        "        select top 1 max(inventory_month)month,inventory_year year   " & _
        "         from inventory_control    " & _
        "        where fix_cls='1'   " & _
        "        group by inventory_year   " & _
        "        order by inventory_year desc   " & _
        "        )tbA  "
    End If
    
    If RsInvControl.State <> adStateClosed Then RsInvControl.Close
    RsInvControl.Open sqlControl, Db, adOpenForwardOnly, adLockReadOnly
    
    If RsInvControl.EOF = False Then '#Inventory CLosing Data exist
        If Request <> "fulldate" Then
            RsInvControl.MoveFirst
            InvYear = Trim(RsInvControl!Inventory_Year)
            InvMonth = Trim(RsInvControl!Inventory_Month)
        End If
    End If
     
    If Request = "month" Then '#Request for month
        uf_GetLastClosing = InvMonth
    ElseIf Request = "year" Then '#Request for year
        uf_GetLastClosing = InvYear
    ElseIf Request = "fulldate" Then '#Request for fulldate
        uf_GetLastClosing = IIf(IsNull(RsInvControl!closingdate), 0, Format(RsInvControl!closingdate, "yyyy-MM-dd"))
    End If
    
    RsInvControl.Close

End Function

Sub uf_UpdateSupplyStockMaster()
    Dim RS As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim prm As ADODB.Parameter
    Dim li_Row As Integer

    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_update_stock_EZR"
    
    cmd.Parameters.append cmd.CreateParameter("TransDate", adDBTime, adParamInput, , DTPicker3.Value)
    cmd.Parameters.append cmd.CreateParameter("WHCode", adVarChar, adParamInput, 15, cbo_warehouse.Text)
    cmd.Parameters.append cmd.CreateParameter("ItemCode", adVarChar, adParamInput, 25, ProductionCode)
    cmd.Parameters.append cmd.CreateParameter("Status", adVarChar, adParamInput, 10, "S")
    Set prm = cmd.CreateParameter("Qty", adNumeric, adParamInput, , Qty)
    prm.Precision = 18
    prm.NumericScale = 5
    cmd.Parameters.append prm
    
    
    Set RS = cmd.Execute
    
End Sub

Sub uf_UpdateReceiptStockMaster()
    Dim RS As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim prm As ADODB.Parameter
    Dim li_Row As Integer

    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_update_stock_EZR"
    
    cmd.Parameters.append cmd.CreateParameter("TransDate", adDBTime, adParamInput, , DTPicker3.Value)
    cmd.Parameters.append cmd.CreateParameter("WHCode", adVarChar, adParamInput, 15, cbo_location.Text)
    cmd.Parameters.append cmd.CreateParameter("ItemCode", adVarChar, adParamInput, 25, ProductionCode)
    cmd.Parameters.append cmd.CreateParameter("Status", adVarChar, adParamInput, 10, "R")
    Set prm = cmd.CreateParameter("Qty", adNumeric, adParamInput, , Qty)
    prm.Precision = 18
    prm.NumericScale = 5
    cmd.Parameters.append prm
    
    
    Set RS = cmd.Execute
    
End Sub

Sub uf_UpdateStockMaster()
    Dim RS As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim prm As ADODB.Parameter
    Dim li_Row As Integer

    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_update_stock_EZR"
    
    cmd.Parameters.append cmd.CreateParameter("TransDate", adDBTime, adParamInput, , DTPicker3.Value)
    cmd.Parameters.append cmd.CreateParameter("WHCode", adVarChar, adParamInput, 15, cbo_warehouse.Text)
    cmd.Parameters.append cmd.CreateParameter("ItemCode", adVarChar, adParamInput, 25, ProductionCode)
    cmd.Parameters.append cmd.CreateParameter("Status", adVarChar, adParamInput, 10, cbo_supply.Text)
    Set prm = cmd.CreateParameter("Qty", adNumeric, adParamInput, , Qty)
    prm.Precision = 18
    prm.NumericScale = 5
    cmd.Parameters.append prm
    
    Set RS = cmd.Execute
    
End Sub

