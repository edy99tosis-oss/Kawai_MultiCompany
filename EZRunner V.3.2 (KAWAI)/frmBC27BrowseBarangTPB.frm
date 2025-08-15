VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Begin VB.Form frmBC27BrowseBarangTPB 
   BackColor       =   &H00FDDFE3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Load Barang TPB"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9165
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   9165
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtHideNoAju 
      Height          =   375
      Left            =   6600
      TabIndex        =   6
      Top             =   6960
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   8895
      Begin VB.TextBox txtNomorAju 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         Tag             =   "TTFF*/"
         Top             =   240
         Width           =   3015
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H0080FFFF&
         Caption         =   "&Search"
         Height          =   375
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   210
         Width           =   975
      End
      Begin VB.Label Label50 
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor Aju"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Tag             =   "TTFF*/"
         Top             =   270
         Width           =   1335
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid gridAju 
      Height          =   2535
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Tag             =   "TTTT*/"
      Top             =   1080
      Width           =   8895
      _cx             =   15690
      _cy             =   4471
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
   Begin VSFlex8Ctl.VSFlexGrid gridHS 
      Height          =   2895
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Tag             =   "TTTT*/"
      Top             =   3960
      Width           =   8895
      _cx             =   15690
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
End
Attribute VB_Name = "frmBC27BrowseBarangTPB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



'-------------------------------------------
Const colNomorAju As Integer = 0
Const colSuratJalan As Integer = 1
Const colTanggal As Integer = 2
Const colKodeKantor As Integer = 3
Const colNamaKantor As Integer = 4
Const colCountAju As Integer = 5


'-------------------------------------------
Const colNomorSeri As Integer = 0
Const colNomorHS As Integer = 1
Const colUraian As Integer = 2
Const colKodeBarang As Integer = 3
Const colTipe As Integer = 4
Const colUkuran As Integer = 5
Const colSpfLain As Integer = 6
Const colMerk As Integer = 7
Const colCIF As Integer = 8
Const colJumlahSatuan As Integer = 9
Const colKodeSatuan As Integer = 10
Const colNetto As Integer = 11
Const colHargaPenyerahan As Integer = 12
Const colCountHS As Integer = 13

Private Sub up_GridHeaderAju()
    
    With gridAju
        .ColS = colCountAju
        .Rows = 1

        .TextMatrix(0, colNomorAju) = "Nomor Aju"
        .TextMatrix(0, colSuratJalan) = "Surat Jalan"
        .TextMatrix(0, colTanggal) = "Tanggal"

        .ColWidth(colNomorAju) = 3500
        .ColWidth(colSuratJalan) = 2000
        .ColWidth(colTanggal) = 1500
        .ColAlignment(colNomorAju) = flexAlignLeftCenter
        .ColAlignment(colSuratJalan) = flexAlignLeftCenter
        
        .ColFormat(colTanggal) = "dd MMM yyyy"
        .ColHidden(colKodeKantor) = True
        .ColHidden(colNamaKantor) = True
    End With
End Sub

Private Sub up_GridHeaderHS()
    
    With gridHS
        .ColS = colCountHS
        .Rows = 1

        .TextMatrix(0, colNomorSeri) = "Nomor Seri"
        .TextMatrix(0, colNomorHS) = "Nomor HS"
        .TextMatrix(0, colUraian) = "Uraian"

        .ColWidth(colNomorSeri) = 700
        .ColWidth(colNomorHS) = 2000
        .ColWidth(colUraian) = 4500
        .ColAlignment(colNomorHS) = flexAlignLeftCenter
        
        .ColHidden(colKodeBarang) = True
        .ColHidden(colTipe) = True
        .ColHidden(colUkuran) = True
        .ColHidden(colMerk) = True
        .ColHidden(colSpfLain) = True
        .ColHidden(colCIF) = True
        .ColHidden(colKodeSatuan) = True
        .ColHidden(colJumlahSatuan) = True
        .ColHidden(colHargaPenyerahan) = True
        .ColHidden(colNetto) = True
        

    End With
End Sub


Private Sub up_GridLoadAju()
    Dim li_Row As Integer
    up_GridHeaderAju
    
    Dim sql As String
    Dim RS As New Recordset

    sql = " Select H.NO_PENGAJUAN, SuratJalan_No, B.Pengajuan_Date, H.KODE_KANTOR, K.Nama_Kantor From Bea_Cukai_TPB_Header H " & vbCrLf & _
        " Left Join Bea_Cukai23_Doc B On H.NO_PENGAJUAN = Replace(B.No_Pengajuan,'-','') " & vbCrLf & _
        " Left Join Bea_Cukai_Kantor_Pabean K On K.Kode_Kantor = H.KODE_KANTOR " & vbCrLf & _
        " Where H.NO_PENGAJUAN like '%" & txtNomorAju & "%' " & vbCrLf & _
        "  " & vbCrLf & _
        "  "
    Set RS = Db.Execute(sql)
        
    With gridAju
        While Not RS.EOF
            .Rows = .Rows + 1
            li_Row = .Rows - 1
            i = i + 1
            
            .TextMatrix(li_Row, colNomorAju) = Trim(RS!No_Pengajuan)
            .TextMatrix(li_Row, colSuratJalan) = Trim(RS!SuratJalan_No)
            .TextMatrix(li_Row, colTanggal) = RS!Pengajuan_Date
            .TextMatrix(li_Row, colKodeKantor) = RS!KODE_KANTOR
            .TextMatrix(li_Row, colNamaKantor) = RS!Nama_Kantor
            
            RS.MoveNext
        Wend
        RS.Close
        Set RS = Nothing
    End With

End Sub

Private Sub up_GridLoadHS(pNomorAju As String)
    Dim li_Row As Integer
    up_GridHeaderHS
    
    Dim sql As String
    Dim RS As New Recordset

    sql = "Select SERI_BARANG, POS_TARIF, URAIAN, KODE_BARANG, CIF, HARGA_PENYERAHAN, " & vbCrLf & _
          "JUMLAH_SATUAN, KODE_SATUAN, NETTO, TIPE, UKURAN, MERK, SPESIFIKASI_LAIN " & vbCrLf & _
          "From Bea_Cukai_TPB_Barang Where NO_PENGAJUAN = '" & pNomorAju & "' Order By SERI_BARANG"
          
    Set RS = Db.Execute(sql)
        
    With gridHS
        While Not RS.EOF
            .Rows = .Rows + 1
            li_Row = .Rows - 1
            i = i + 1
            
            .TextMatrix(li_Row, colNomorSeri) = Trim(RS!SERI_BARANG)
            .TextMatrix(li_Row, colNomorHS) = Trim(RS!POS_TARIF)
            .TextMatrix(li_Row, colUraian) = RS!URAIAN
            .TextMatrix(li_Row, colKodeBarang) = RS!Kode_Barang
            .TextMatrix(li_Row, colTipe) = RS!Tipe
            .TextMatrix(li_Row, colUkuran) = RS!Ukuran
            .TextMatrix(li_Row, colMerk) = RS!MERK
            .TextMatrix(li_Row, colSpfLain) = RS!SPESIFIKASI_LAIN
            .TextMatrix(li_Row, colCIF) = RS!CIF
            .TextMatrix(li_Row, colJumlahSatuan) = RS!JUMLAH_SATUAN
            .TextMatrix(li_Row, colKodeSatuan) = RS!KODE_SATUAN
            .TextMatrix(li_Row, colNetto) = IIf(IsNull(RS.Fields("NETTO")), 0, RS.Fields("NETTO"))
            .TextMatrix(li_Row, colHargaPenyerahan) = IIf(IsNull(RS.Fields("HARGA_PENYERAHAN")), 0, RS.Fields("HARGA_PENYERAHAN"))
            
            RS.MoveNext
        Wend
        RS.Close
        Set RS = Nothing
    End With
End Sub

Private Sub cmdSearch_Click()
up_GridLoadAju
End Sub

Private Sub Form_Load()
up_GridLoadAju
up_GridLoadHS "-"
End Sub

Private Sub gridAju_Click()
If gridAju.RowSel > 0 Then
    up_GridLoadHS gridAju.TextMatrix(gridAju.RowSel, colNomorAju)
    txtHideNoAju = gridAju.TextMatrix(gridAju.RowSel, colNomorAju)
End If
End Sub

Private Sub gridHS_DblClick()
If gridHS.RowSel > 0 Then
    If frmBC27BrowseBarang.SSTab1.Tab = 1 Then
        frmBC27BrowseBarang.txtNoAjuImpor = txtHideNoAju
        frmBC27BrowseBarang.txtDokumenAsalImpor = "23"
        frmBC27BrowseBarang.lblDokAsalImpor = "BC 2.3"
        frmBC27BrowseBarang.txtNomorHSImpor = gridHS.TextMatrix(gridHS.RowSel, colNomorHS)
        frmBC27BrowseBarang.txtUraianBarangImpor = gridHS.TextMatrix(gridHS.RowSel, colUraian)
        frmBC27BrowseBarang.txtUrutKeImpor = gridHS.TextMatrix(gridHS.RowSel, colNomorSeri)
        frmBC27BrowseBarang.txtKPPBCImpor = gridAju.TextMatrix(gridAju.RowSel, colKodeKantor)
        frmBC27BrowseBarang.lblKPPBCImpor = gridAju.TextMatrix(gridAju.RowSel, colNamaKantor)
        
        frmBC27BrowseBarang.txtKodeBarangImpor = gridHS.TextMatrix(gridHS.RowSel, colKodeBarang)
        frmBC27BrowseBarang.txtTipeImpor = gridHS.TextMatrix(gridHS.RowSel, colTipe)
        'frmBC27BrowseBarang.txtHargaCIFUSDImpor = Format(gridHS.TextMatrix(gridHS.RowSel, colCIF), "#,0.00")
        
        'frmBC27BrowseBarang.txtNettoImpor = Format(gridHS.TextMatrix(gridHS.RowSel, colNetto), "#,0.00")
    '    frmBC25BrowseBarang.gd_HargaPenyerahanImpor = gridHS.TextMatrix(gridHS.RowSel, colHargaPenyerahan)
    Else
        frmBC27BrowseBarang.txtNoAjuLokal = txtHideNoAju
        frmBC27BrowseBarang.txtDokumenAsalLokal = "23"
        frmBC27BrowseBarang.lblDokAsalLokal = "BC 2.3"
        frmBC27BrowseBarang.txtNomorHSLokal = gridHS.TextMatrix(gridHS.RowSel, colNomorHS)
        frmBC27BrowseBarang.txtUraianBarangLokal = gridHS.TextMatrix(gridHS.RowSel, colUraian)
        frmBC27BrowseBarang.txtUrutKeLokal = gridHS.TextMatrix(gridHS.RowSel, colNomorSeri)
        frmBC27BrowseBarang.txtKPPBCLokal = gridAju.TextMatrix(gridAju.RowSel, colKodeKantor)
        frmBC27BrowseBarang.lblKPPBCLokal = gridAju.TextMatrix(gridAju.RowSel, colNamaKantor)
        
        frmBC27BrowseBarang.txtKodeBarangLokal = gridHS.TextMatrix(gridHS.RowSel, colKodeBarang)
        frmBC27BrowseBarang.txtTipeLokal = gridHS.TextMatrix(gridHS.RowSel, colTipe)
        'frmBC27BrowseBarang.txtHargaCIFUSDLokal = Format(gridHS.TextMatrix(gridHS.RowSel, colCIF), "#,0.00")
        'frmBC27BrowseBarang.txtNDPBMLokal = Format(gridHS.TextMatrix(gridHS.RowSel, colHargaPenyerahan), "#,0.00")
        
'        frmBC25BrowseBarang.txtNettoLokal = Format(gridHS.TextMatrix(gridHS.RowSel, colNetto), "#,0.00")
    '    frmBC25BrowseBarang.gd_HargaPenyerahanImpor = gridHS.TextMatrix(gridHS.RowSel, colHargaPenyerahan)
    End If
    
    Me.Hide
End If
End Sub
