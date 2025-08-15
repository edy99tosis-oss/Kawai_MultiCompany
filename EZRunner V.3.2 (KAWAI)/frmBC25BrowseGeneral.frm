VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmBC25BrowseGeneral 
   BackColor       =   &H00FDDFE3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Browse Data"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9420
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBC25BrowseGeneral.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   9420
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8895
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H0080FFFF&
         Caption         =   "&Search"
         Height          =   375
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   210
         Width           =   975
      End
      Begin VB.TextBox txtCari 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2640
         TabIndex        =   3
         Tag             =   "TTFF*/"
         Top             =   240
         Width           =   3015
      End
      Begin MSForms.ComboBox cboKriteria 
         Height          =   315
         Left            =   720
         TabIndex        =   2
         Tag             =   "TTFF*/"
         Top             =   240
         Width           =   1815
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3201;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label50 
         BackStyle       =   0  'Transparent
         Caption         =   "Find"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Tag             =   "TTFF*/"
         Top             =   270
         Width           =   855
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid grid 
      Height          =   5295
      Left            =   240
      TabIndex        =   4
      TabStop         =   0   'False
      Tag             =   "TTTT*/"
      Top             =   1080
      Width           =   8895
      _cx             =   15690
      _cy             =   9340
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
      SelectionMode   =   1
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
End
Attribute VB_Name = "frmBC25BrowseGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public gs_TableName As String

Private Sub up_FillCombo()
    With cboKriteria
        .clear
        .AddItem "1 - Kode"
        .AddItem "2 - Uraian"
        
        .ListWidth = 110
        .ListRows = 15
        
        .ListIndex = 0
    End With
End Sub

Private Sub up_GridLoadBarang(pKriteria As String, pSearch As String)
Dim RS As New Recordset
Dim pFilter As String

If pKriteria = "1" Then
    pFilter = "Where Kode like '%" & pSearch & "%'"
Else
    pFilter = "Where Uraian like '%" & pSearch & "%'"
End If

sql = "Select * From (Select Kode = Kode_Barang, Uraian = Uraian_Barang, Merk, NOHS From Bea_Cukai_Kode_Barang) A " & pFilter

Set RS = Db.Execute(sql)
    
With Grid
    Set .DataSource = RS
    
    .TextMatrix(0, 0) = "Kode"
    .TextMatrix(0, 1) = "Uraian"
End With
End Sub

Private Sub up_GridLoadKPPBC(pKriteria As String, pSearch As String)
Dim RS As New Recordset
Dim pFilter As String

If pKriteria = "1" Then
    pFilter = "Where Kode like '%" & pSearch & "%'"
Else
    pFilter = "Where Uraian like '%" & pSearch & "%'"
End If

sql = "Select * From (Select Kode = KODE_KANTOR, Uraian = NAMA_KANTOR From Bea_Cukai_Kantor_Pabean) A " & pFilter

Set RS = Db.Execute(sql)
    
With Grid
    Set .DataSource = RS
    
    .TextMatrix(0, 0) = "Kode"
    .TextMatrix(0, 1) = "Uraian"
End With
End Sub

Private Sub up_GridLoadDokumenAsal(pKriteria As String, pSearch As String)
Dim RS As New Recordset
Dim pFilter As String

If pKriteria = "1" Then
    pFilter = "Where Kode like '%" & pSearch & "%'"
Else
    pFilter = "Where Uraian like '%" & pSearch & "%'"
End If

sql = "Select * From (Select Kode= Kode_Dokumen, Uraian = Uraian_Dokumen From Bea_Cukai_Dokumen Where Kode_Dokumen in (16,23,27,52) Union All Select 99, 'LAINNYA') A " & pFilter

Set RS = Db.Execute(sql)
    
With Grid
    Set .DataSource = RS
    
    .TextMatrix(0, 0) = "Kode"
    .TextMatrix(0, 1) = "Uraian"
End With
End Sub

Private Sub up_GridLoadKondisi(pKriteria As String, pSearch As String)
Dim RS As New Recordset
Dim pFilter As String

If pKriteria = "1" Then
    pFilter = "Where Kode like '%" & pSearch & "%'"
Else
    pFilter = "Where Uraian like '%" & pSearch & "%'"
End If

sql = "Select * From (Select Kode = KODE_KONDISI, Uraian = URAIAN_KONDISI From Bea_Cukai_Kondisi_Barang) A " & pFilter

Set RS = Db.Execute(sql)
    
With Grid
    Set .DataSource = RS
    
    .TextMatrix(0, 0) = "Kode"
    .TextMatrix(0, 1) = "Uraian"
End With
End Sub

Private Sub up_GridLoadKategori(pKriteria As String, pSearch As String)
Dim RS As New Recordset
Dim pFilter As String

If pKriteria = "1" Then
    pFilter = "Where Kode like '%" & pSearch & "%'"
Else
    pFilter = "Where Uraian like '%" & pSearch & "%'"
End If

sql = "Select * From (Select Kode = KODE_KATEGORI, Uraian = URAIAN_KATEGORI From Bea_Cukai_Kategori_BarangBC25) A " & pFilter

Set RS = Db.Execute(sql)
    
With Grid
    Set .DataSource = RS
    
    .TextMatrix(0, 0) = "Kode"
    .TextMatrix(0, 1) = "Uraian"
End With
End Sub

Private Sub up_GridLoadPenggunaan(pKriteria As String, pSearch As String)
Dim RS As New Recordset
Dim pFilter As String

If pKriteria = "1" Then
    pFilter = "Where Kode like '%" & pSearch & "%'"
Else
    pFilter = "Where Uraian like '%" & pSearch & "%'"
End If

sql = "Select * From (Select Kode = KODE_GUNA, Uraian = URAIAN_GUNA From Bea_Cukai_Kode_Guna) A " & pFilter

Set RS = Db.Execute(sql)
    
With Grid
    Set .DataSource = RS
    
    .TextMatrix(0, 0) = "Kode"
    .TextMatrix(0, 1) = "Uraian"
End With
End Sub

Private Sub up_GridLoadValuta(pKriteria As String, pSearch As String)
Dim RS As New Recordset
Dim pFilter As String

If pKriteria = "1" Then
    pFilter = "Where Kode like '%" & pSearch & "%'"
Else
    pFilter = "Where Uraian like '%" & pSearch & "%'"
End If

sql = "Select * From (Select Kode = KODE_VALUTA, Uraian = URAIAN_VALUTA From Bea_Cukai_Valuta) A " & pFilter

Set RS = Db.Execute(sql)
    
With Grid
    Set .DataSource = RS
    
    .TextMatrix(0, 0) = "Kode"
    .TextMatrix(0, 1) = "Uraian"
End With
End Sub

Private Sub up_GridLoadKemasan(pKriteria As String, pSearch As String)
Dim RS As New Recordset
Dim pFilter As String

If pKriteria = "1" Then
    pFilter = "Where Kode like '%" & pSearch & "%'"
Else
    pFilter = "Where Uraian like '%" & pSearch & "%'"
End If

sql = "Select * From (Select Kode = KODE_KEMASAN, Uraian = URAIAN_KEMASAN From Bea_Cukai_Kemasan) A " & pFilter

Set RS = Db.Execute(sql)
    
With Grid
    Set .DataSource = RS
    
    .TextMatrix(0, 0) = "Kode"
    .TextMatrix(0, 1) = "Uraian"
End With
End Sub

Private Sub up_GridLoadKantorPabean(pKriteria As String, pSearch As String)
Dim RS As New Recordset
Dim pFilter As String

If pKriteria = "1" Then
    pFilter = "Where Kode like '%" & pSearch & "%'"
Else
    pFilter = "Where Uraian like '%" & pSearch & "%'"
End If

sql = "Select * From (Select Kode = KODE_KANTOR, Uraian = NAMA_KANTOR From Bea_Cukai_Kantor_Pabean) A " & pFilter

Set RS = Db.Execute(sql)
    
With Grid
    Set .DataSource = RS
    
    .TextMatrix(0, 0) = "Kode"
    .TextMatrix(0, 1) = "Uraian"
End With
End Sub

Private Sub up_GridLoadSatuan(pKriteria As String, pSearch As String)
Dim RS As New Recordset
Dim pFilter As String

If pKriteria = "1" Then
    pFilter = "Where Kode like '%" & pSearch & "%'"
Else
    pFilter = "Where Uraian like '%" & pSearch & "%'"
End If

sql = "Select * From (Select Kode = KODE_SATUAN, Uraian = URAIAN_SATUAN From Bea_Cukai_Satuan) A " & pFilter

Set RS = Db.Execute(sql)
    
With Grid
    Set .DataSource = RS
    
    .TextMatrix(0, 0) = "Kode"
    .TextMatrix(0, 1) = "Uraian"
End With
End Sub

Private Sub up_GridLoad()
If gs_TableName = "Kantor Pabean" Then
    up_GridLoadKantorPabean Left(cboKriteria, 1), txtCari
ElseIf gs_TableName = "Kemasan" Then
    up_GridLoadKemasan Left(cboKriteria, 1), txtCari
ElseIf gs_TableName = "Valuta" Then
    up_GridLoadValuta Left(cboKriteria, 1), txtCari
ElseIf gs_TableName = "Penggunaan" Then
    up_GridLoadPenggunaan Left(cboKriteria, 1), txtCari
ElseIf gs_TableName = "Kategori" Then
    up_GridLoadKategori Left(cboKriteria, 1), txtCari
ElseIf gs_TableName = "Kondisi" Then
    up_GridLoadKondisi Left(cboKriteria, 1), txtCari
ElseIf gs_TableName = "Dokumen Asal" Then
    up_GridLoadDokumenAsal Left(cboKriteria, 1), txtCari
ElseIf gs_TableName = "KPPBC Impor" Then
    up_GridLoadKPPBC Left(cboKriteria, 1), txtCari
ElseIf gs_TableName = "KPPBC Lokal" Then
    up_GridLoadKPPBC Left(cboKriteria, 1), txtCari
ElseIf gs_TableName = "Barang Impor" Then
    up_GridLoadBarang Left(cboKriteria, 1), txtCari
ElseIf gs_TableName = "Barang" Then
    up_GridLoadBarang Left(cboKriteria, 1), txtCari
Else
    up_GridLoadSatuan Left(cboKriteria, 1), txtCari
End If
End Sub

Private Sub cmdSearch_Click()
    up_GridLoad
End Sub

Private Sub Form_Activate()
    up_GridLoad
End Sub

Private Sub Form_Load()
    up_FillCombo
End Sub

Private Sub Grid_DblClick()
If gs_TableName = "Kantor Pabean" Then
    frmBC25Detail.txtKPBBCBongkar = Grid.TextMatrix(Grid.RowSel, 0)
    frmBC25Detail.lblKPPBCBongkar = Grid.TextMatrix(Grid.RowSel, 1)
ElseIf gs_TableName = "Kemasan" Then
    frmBC25Detail.txtJenisKemasan = Grid.TextMatrix(Grid.RowSel, 0)
    frmBC25Detail.lblJenisKemasan = Grid.TextMatrix(Grid.RowSel, 1)
ElseIf gs_TableName = "Valuta" Then
    frmBC25Detail.txtValuta = Grid.TextMatrix(Grid.RowSel, 0)
    frmBC25Detail.lblValuta = Grid.TextMatrix(Grid.RowSel, 1)
ElseIf gs_TableName = "Penggunaan" Then
    frmBC25BrowseBarang.txtPenggunaan = Grid.TextMatrix(Grid.RowSel, 0)
    frmBC25BrowseBarang.lblPenggunaan = Grid.TextMatrix(Grid.RowSel, 1)
ElseIf gs_TableName = "Kategori" Then
    frmBC25BrowseBarang.txtKategoriBarang = Grid.TextMatrix(Grid.RowSel, 0)
    frmBC25BrowseBarang.lblKategori = Grid.TextMatrix(Grid.RowSel, 1)
ElseIf gs_TableName = "Kondisi" Then
    frmBC25BrowseBarang.txtKondisiBarang = Grid.TextMatrix(Grid.RowSel, 0)
    frmBC25BrowseBarang.lblKondisiBarang = Grid.TextMatrix(Grid.RowSel, 1)
ElseIf gs_TableName = "Dokumen Asal" Then
    frmBC25BrowseBarang.txtDokumenAsalImpor = Grid.TextMatrix(Grid.RowSel, 0)
    frmBC25BrowseBarang.lblDokAsalImpor = Grid.TextMatrix(Grid.RowSel, 1)
ElseIf gs_TableName = "KPPBC Impor" Then
    frmBC25BrowseBarang.txtKPPBCImpor = Grid.TextMatrix(Grid.RowSel, 0)
    frmBC25BrowseBarang.lblKPPBCImpor = Grid.TextMatrix(Grid.RowSel, 1)
ElseIf gs_TableName = "KPPBC Lokal" Then
    frmBC25BrowseBarang.txtKPPBCLokal = Grid.TextMatrix(Grid.RowSel, 0)
    frmBC25BrowseBarang.lblKPPBCLokal = Grid.TextMatrix(Grid.RowSel, 1)
ElseIf gs_TableName = "Barang Impor" Then
    frmBC25BrowseBarang.txtKodeBarangImpor = Grid.TextMatrix(Grid.RowSel, 0)
    frmBC25BrowseBarang.txtUraianBarangImpor = Grid.TextMatrix(Grid.RowSel, 1)
    frmBC25BrowseBarang.txtMerkImpor = Grid.TextMatrix(Grid.RowSel, 2)
    frmBC25BrowseBarang.txtNomorHSImpor = Grid.TextMatrix(Grid.RowSel, 3)
ElseIf gs_TableName = "Barang" Then
    frmBC25BrowseBarang.txtKodeBarang = Grid.TextMatrix(Grid.RowSel, 0)
    frmBC25BrowseBarang.txtUraianBarang = Grid.TextMatrix(Grid.RowSel, 1)
    frmBC25BrowseBarang.txtMerk = Grid.TextMatrix(Grid.RowSel, 2)
    frmBC25BrowseBarang.txtNomorHS = Grid.TextMatrix(Grid.RowSel, 3)
End If
    Unload Me
End Sub
