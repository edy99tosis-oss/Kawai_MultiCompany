VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FrmBC27LoadBarang 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Load Barang"
   ClientHeight    =   8655
   ClientLeft      =   5550
   ClientTop       =   1950
   ClientWidth     =   10020
   Icon            =   "FrmBC27LoadBarang.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8655
   ScaleWidth      =   10020
   StartUpPosition =   2  'CenterScreen
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
      Left            =   1770
      MaxLength       =   30
      TabIndex        =   1
      Top             =   840
      Width           =   6540
   End
   Begin VB.TextBox Text1 
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
      Left            =   1770
      MaxLength       =   30
      TabIndex        =   0
      Top             =   4590
      Width           =   6540
   End
   Begin VSFlex8Ctl.VSFlexGrid GridHeader 
      Height          =   2895
      Left            =   210
      TabIndex        =   2
      TabStop         =   0   'False
      Tag             =   "TTTT*/"
      Top             =   1320
      Width           =   9525
      _cx             =   16801
      _cy             =   5106
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
   Begin VSFlex8Ctl.VSFlexGrid GridBarang 
      Height          =   3375
      Left            =   210
      TabIndex        =   3
      TabStop         =   0   'False
      Tag             =   "TTTT*/"
      Top             =   5040
      Width           =   9525
      _cx             =   16801
      _cy             =   5953
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nomor Aju"
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
      Left            =   285
      TabIndex        =   6
      Top             =   810
      Width           =   915
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Load Barang TPB"
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
      Height          =   330
      Left            =   3990
      TabIndex        =   5
      Top             =   240
      Width           =   1950
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nomor HS"
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
      Left            =   285
      TabIndex        =   4
      Top             =   4605
      Width           =   870
   End
End
Attribute VB_Name = "FrmBC27LoadBarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bteColSeri As Byte, bteColKode As Byte, bteColUraian As Byte, bteColMerk As Byte, bteColTipe As Byte, bteColUkuran As Byte
Dim bteColSPF As Byte, bteColKodeSatuan As Byte, bteColJumlahSatuan As Byte, bteColHarga As Byte, bteColSelect As Byte, bteColCIF As Byte
Dim tampung As String, DokAsal As String, NoDok As String, KPPBC As String, NoAju As String, bteColCIFRupiah As String, bteColHSNo As String
Dim Tgl As Date


Const ColCheck As Integer = 0
Const ColNoAju As Integer = 1
Const ColNoDaftar As Integer = 2
Const ColTglDaftar As Integer = 3
Const ColStatus As Integer = 4
Const colIdGridHeader As Integer = 5
Const colKodeDokumen As Integer = 6
Const colKodeKantor As Integer = 7


Private Sub up_GridHeaderGrid()
    With GridHeader
        .clear
         .Rows = 1
         .ColS = 8
        
        .TextMatrix(0, ColCheck) = ""
        .TextMatrix(0, ColNoAju) = "Nomor Aju"
        .TextMatrix(0, ColNoDaftar) = "Nomor Daftar"
        .TextMatrix(0, ColTglDaftar) = "Tanggal Daftar"
        .TextMatrix(0, ColStatus) = "Status"
        .TextMatrix(0, colIdGridHeader) = "Id GridHeader"
        .TextMatrix(0, colKodeDokumen) = "Kode Dokumen"
        .TextMatrix(0, colKodeKantor) = "KPPBC"
               
        .ColWidth(ColCheck) = 300
        .ColWidth(ColNoAju) = 3500
        .ColWidth(ColNoDaftar) = 1800
        .ColWidth(ColTglDaftar) = 1800
        .ColWidth(ColStatus) = 2000
        .ColWidth(colIdGridHeader) = 500
        .ColWidth(colKodeDokumen) = 500
        .ColWidth(colKodeKantor) = 500
                
        .Cell(flexcpAlignment, 0, 0, 0, 4) = flexAlignCenterCenter
        
        .ColHidden(colIdGridHeader) = True
        .ColHidden(colKodeDokumen) = True
        .ColHidden(colKodeKantor) = True
        .ColHidden(ColCheck) = True
    End With
End Sub

Private Sub Form_Load()
    up_GridHeaderLoad
    up_GridBarang
End Sub

Private Sub up_GridHeaderLoad()
Dim strSQL As String
Dim RS As ADODB.Recordset
Dim ls_sql As String
'Dim rs As ADODB.Recordset
Dim cmd As ADODB.Command

Set RS = New ADODB.Recordset

    KoneksiMysql
        
            strSQL = " SELECT a.id, KODE_DOKUMEN, KODE_KANTOR, Nomor_aju, nomor_daftar, tanggal_daftar, " & vbCrLf & _
                     " a.kode_status, b.uraian_status FROM tpbdb.tpb_header as a Left Join tpbdb.referensi_status as b " & vbCrLf & _
                     " on a.kode_status=b.kode_status where b.kode_dokumen='23' and a.kode_status='80' and nomor_aju not like '050840%' "
                     
            
            'rsId.Open strSQL, ConnStr
            If RS.State <> adStateClosed Then RS.Close
            RS.Open strSQL, ConnStr, adOpenForwardOnly, adLockReadOnly, adCmdText
            
        
        'If rs.EOF = False Then

    
    'LblErrMsg.Caption = ""
    
    up_GridHeaderGrid
    
    Me.MousePointer = vbHourglass
    
    If RS.EOF = False Then
    
        i = 1
        With GridHeader
            While Not RS.EOF
                .Rows = .Rows + 1
                
                .TextMatrix(i, ColNoAju) = Trim(RS("Nomor_Aju"))
                .TextMatrix(i, ColNoDaftar) = Trim(RS("Nomor_Daftar"))
                .TextMatrix(i, ColTglDaftar) = Format(Trim(RS("tanggal_daftar")), "yyyy-MM-dd")
                .TextMatrix(i, ColStatus) = Trim(RS("Uraian_Status"))
                .TextMatrix(i, colIdGridHeader) = Trim(RS("id"))
                .TextMatrix(i, colKodeDokumen) = Trim(RS("Kode_Dokumen"))
                .TextMatrix(i, colKodeKantor) = Trim(RS("Kode_Kantor"))
                
                
                i = i + 1
            RS.MoveNext
            Wend
        End With
        
        'LblRecord = Format(i - 1, "#,##0") & " Record(s)"
        
        Me.MousePointer = vbDefault
    
    Else
    
        'LblErrMsg.Caption = DisplayMsg(13)
        
        Me.MousePointer = vbDefault
    
    End If
End Sub

Private Sub GridBarang_Click()
'GridBarang = 1
    If gridBarang <> "" Then
        frmBC27BrowseBarang.txtDokumenAsalImpor = DokAsal
        frmBC27BrowseBarang.txtNoImpor = NoDok
        frmBC27BrowseBarang.dtpTglImpor = Tgl
        frmBC27BrowseBarang.txtKodeBarangImpor = gridBarang.TextMatrix(gridBarang.Row, bteColKode)
        frmBC27BrowseBarang.txtUraianBarangImpor = gridBarang.TextMatrix(gridBarang.Row, bteColUraian)
        frmBC27BrowseBarang.txtMerkImpor = gridBarang.TextMatrix(gridBarang.Row, bteColMerk)
        frmBC27BrowseBarang.txtTipeImpor = gridBarang.TextMatrix(gridBarang.Row, bteColTipe)
        frmBC27BrowseBarang.txtUkuranImpor = gridBarang.TextMatrix(gridBarang.Row, bteColUkuran)
        frmBC27BrowseBarang.txtSpfLainImpor = gridBarang.TextMatrix(gridBarang.Row, bteColSPF)
        frmBC27BrowseBarang.txtSatuanImpor = gridBarang.TextMatrix(gridBarang.Row, bteColKodeSatuan)
        frmBC27BrowseBarang.txtJumlahSatuanImpor = gridBarang.TextMatrix(gridBarang.Row, bteColJumlahSatuan)
        frmBC27BrowseBarang.txtHargaPenyerahanImpor = Format(gridBarang.TextMatrix(gridBarang.Row, bteColHarga), "#,0.00")
        frmBC27BrowseBarang.txtUrutKeImpor = gridBarang.TextMatrix(gridBarang.Row, bteColSeri)
        frmBC27BrowseBarang.txtNoAjuImpor = NoAju
        frmBC27BrowseBarang.txtKPPBCImpor = KPPBC
        frmBC27BrowseBarang.dtpTglImpor = Tgl
        frmBC27BrowseBarang.txtNomorHSImpor = gridBarang.TextMatrix(gridBarang.Row, bteColHSNo)
        
        FrmBC27LoadBarang.Hide
        
    End If
End Sub

Private Sub GridBarang_KeyPress(KeyAscii As Integer)
If gridBarang.Col = ColCheck Then
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii <> Asc("S") And KeyAscii <> Asc("D") And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then
      KeyAscii = 0
    End If
    If KeyAscii = Asc(".") Then KeyAscii = 0
  End If
End Sub

Private Sub GridHeader_Click()
    If GridHeader > 0 Then
        tampung = GridHeader.TextMatrix(GridHeader.Row, colIdGridHeader)
        Tgl = GridHeader.TextMatrix(GridHeader.Row, ColTglDaftar)
        NoAju = GridHeader.TextMatrix(GridHeader.Row, ColNoAju)
        DokAsal = GridHeader.TextMatrix(GridHeader.Row, colKodeDokumen)
        NoDok = GridHeader.TextMatrix(GridHeader.Row, ColNoDaftar)
        KPPBC = GridHeader.TextMatrix(GridHeader.Row, colKodeKantor)
        Tgl = GridHeader.TextMatrix(GridHeader.Row, ColTglDaftar)
        up_GridDetail
    End If
End Sub

Private Sub GridHeader_KeyPress(KeyAscii As Integer)
  If GridHeader.Col = ColCheck Then
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii <> Asc("S") And KeyAscii <> Asc("D") And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then
      KeyAscii = 0
    End If
    If KeyAscii = Asc(".") Then KeyAscii = 0
  End If
End Sub

Private Sub ClearColGridHeader(Optional Kolom As String)
 Dim i As Integer
    With GridHeader
        .Col = ColCheck
        If Kolom <> "" Then
            For i = 1 To .Rows - 1
                If .Text = Kolom Then .Text = ""
                If .TextMatrix(i, ColCheck) <> "D" Then .TextMatrix(i, ColCheck) = ""
            Next i
            'clear
        Else
            For i = 1 To .Rows - 1
                If .TextMatrix(i, ColCheck) <> "" Then .TextMatrix(i, ColCheck) = ""
            Next i
        End If
    End With
End Sub

Private Sub ClearGridDetail(Optional Kolom As String)
 Dim i As Integer
    With gridBarang
        .Col = ColCheck
        If Kolom <> "" Then
            For i = 1 To .Rows - 1
                If .Text = Kolom Then .Text = ""
                If .TextMatrix(i, bteColSelect) <> "D" Then .TextMatrix(i, bteColSelect) = ""
            Next i
            'clear
        Else
            For i = 1 To .Rows - 1
                If .TextMatrix(i, bteColSelect) <> "" Then .TextMatrix(i, bteColSelect) = ""
            Next i
        End If
    End With
End Sub

Private Sub ClearGridBarang(Optional Kolom As String)
 Dim i As Integer
    With GridHeader
        .Col = ColCheck
        If Kolom <> "" Then
            For i = 1 To .Rows - 1
                If .Text = Kolom Then .Text = ""
                If .TextMatrix(i, ColCheck) <> "D" Then .TextMatrix(i, ColCheck) = ""
            Next i
            'clear
        Else
            For i = 1 To .Rows - 1
                If .TextMatrix(i, ColCheck) <> "" Then .TextMatrix(i, ColCheck) = ""
            Next i
        End If
    End With
End Sub

Private Sub up_GridBarang()

bteColSelect = 0
bteColSeri = 1
bteColKode = 2
bteColUraian = 3
bteColMerk = 4
bteColTipe = 5
bteColUkuran = 6
bteColSPF = 7
bteColKodeSatuan = 8
bteColJumlahSatuan = 9
bteColHarga = 10
bteColCIF = 11
bteColCIFRupiah = 12
bteColHSNo = 13


With gridBarang
        .clear
         .Rows = 1
         .ColS = 14
        
        .TextMatrix(0, bteColSelect) = ""
        .TextMatrix(0, bteColSeri) = "Seri"
        .TextMatrix(0, bteColKode) = "Kode Barang"
        .TextMatrix(0, bteColUraian) = "Uraian"
        .TextMatrix(0, bteColMerk) = "Merk"
        .TextMatrix(0, bteColTipe) = "Tipe"
        .TextMatrix(0, bteColUkuran) = "Ukuran"
        .TextMatrix(0, bteColSPF) = "Spesifikasi Lain"
        .TextMatrix(0, bteColKodeSatuan) = "Kode Satuan"
        .TextMatrix(0, bteColJumlahSatuan) = "Jumlah Satuan"
        .TextMatrix(0, bteColHarga) = "Harga Penyerahan"
        .TextMatrix(0, bteColCIF) = "CIF"
        .TextMatrix(0, bteColCIFRupiah) = "CIF Rupiah"
        .TextMatrix(0, bteColHSNo) = "HS No"
        
        .ColWidth(bteColSelect) = 300
        .ColWidth(bteColSeri) = 800
        .ColWidth(bteColKode) = 2000
        .ColWidth(bteColUraian) = 5000
        .ColWidth(bteColMerk) = 800
        .ColWidth(bteColTipe) = 800
        .ColWidth(bteColUkuran) = 800
        .ColWidth(bteColSPF) = 800
        .ColWidth(bteColKodeSatuan) = 800
        .ColWidth(bteColJumlahSatuan) = 800
        .ColWidth(bteColHarga) = 800
        
        .ColHidden(bteColMerk) = True
        .ColHidden(bteColTipe) = True
        .ColHidden(bteColUkuran) = True
        .ColHidden(bteColSPF) = True
        .ColHidden(bteColKodeSatuan) = True
        .ColHidden(bteColJumlahSatuan) = True
        .ColHidden(bteColHarga) = True
        .ColHidden(bteColSelect) = True
        .ColHidden(bteColCIF) = True
        .ColHidden(bteColCIFRupiah) = True
        .ColHidden(bteColHSNo) = True
        
        .Cell(flexcpAlignment, 0, 0, 0, 9) = flexAlignCenterCenter
End With
        
End Sub

Private Sub up_GridDetail()
Dim strSQL As String
Dim RS As ADODB.Recordset
Dim Row As Long
Dim cmd As ADODB.Command

Set RS = New ADODB.Recordset

up_GridBarang
    
    KoneksiMysql

    strSQL = " SELECT Seri_barang, Kode_barang, uraian, merk, tipe, ukuran, spesifikasi_lain, Jumlah_satuan, kode_satuan, " & vbCrLf & _
             " harga_penyerahan, CIF, CIF_Rupiah, Pos_Tarif FROM tpbdb.tpb_barang where id_header='" & tampung & "'"
    
    If RS.State <> adStateClosed Then RS.Close
    RS.Open strSQL, ConnStr, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If RS.EOF = False Then
    
        i = 1
        With gridBarang
            While Not RS.EOF
                 .Rows = .Rows + 1
                
'                .Cell(flexcpChecked, i, ColCheck) = flexUnchecked
'                .Cell(flexcpBackColor, i, ColCheck) = vbWhite
                .TextMatrix(i, bteColSeri) = Trim(RS("Seri_barang"))
                .TextMatrix(i, bteColKode) = Trim(RS("Kode_Barang"))
                .TextMatrix(i, bteColUraian) = Trim(RS("Uraian"))
                .TextMatrix(i, bteColMerk) = IIf(IsNull(Trim(RS("Merk"))) = True, "-", Trim(RS("Merk")))
                .TextMatrix(i, bteColTipe) = IIf(IsNull(Trim(RS("Tipe"))) = True, "-", Trim(RS("Tipe")))
                .TextMatrix(i, bteColUkuran) = IIf(IsNull(Trim(RS("Ukuran"))) = True, "-", Trim(RS("Ukuran")))
                .TextMatrix(i, bteColSPF) = IIf(IsNull(Trim(RS("Spesifikasi_Lain"))) = True, "-", Trim(RS("Ukuran")))
                .TextMatrix(i, bteColJumlahSatuan) = Trim(RS("Jumlah_satuan"))
                .TextMatrix(i, bteColKodeSatuan) = Trim(RS("Kode_Satuan"))
                .TextMatrix(i, bteColHarga) = IIf(IsNull(Trim(RS("Harga_Penyerahan"))) = True, 0, Trim(RS("Harga_Penyerahan")))
                .TextMatrix(i, bteColCIF) = IIf(IsNull(Trim(RS("CIF"))) = True, 0, Trim(RS("CIF")))
                .TextMatrix(i, bteColCIFRupiah) = IIf(IsNull(Trim(RS("CIF_Rupiah"))) = True, 0, Trim(RS("CIF_Rupiah")))
                .TextMatrix(i, bteColHSNo) = IIf(IsNull(Trim(RS("POS_Tarif"))) = True, 0, Trim(RS("POS_Tarif")))
                i = i + 1
            RS.MoveNext
            Wend
        End With
                
        Me.MousePointer = vbDefault
    
    Else
    
        'LblErrMsg.Caption = DisplayMsg(13)
        
        Me.MousePointer = vbDefault
    
    End If
End Sub

Private Sub KoneksiMysql()
Dim ConnString As String
Dim db_name As String
Dim db_server As String
Dim db_port As String
Dim db_user As String
Dim db_pass As String
Dim Conn As New ADODB.Connection
'//error traping
On Error GoTo buat_koneksi_Error
'/isi variable
db_name = "tpbdb"
db_server = "localhost"
db_port = "3306"
db_user = "beacukai"
db_pass = "beacukai"
'/buat connection string
ConnStr = "DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & db_server & ";DATABASE=" & db_name & ";UID=" & db_user & ";PWD=" & db_pass & ";PORT=" & db_port & ""
'/buka koneksi
With Conn
    .ConnectionString = ConnStr
    .Open
   'MsgBox "Koneksi Berhasil"
End With
'___________________________________________________________
On Error GoTo 0
Exit Sub

buat_koneksi_Error:
    MsgBox "Ada kesalahan dengan server, periksa apakah server sudah berjalan !", vbInformation, "Cek Server"
End Sub



