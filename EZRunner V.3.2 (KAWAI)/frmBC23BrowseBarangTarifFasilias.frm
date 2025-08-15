VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Begin VB.Form frmBC23BrowseBarangTarifFasilitas 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Barang Tarif Fasilitas"
   ClientHeight    =   6180
   ClientLeft      =   5355
   ClientTop       =   2580
   ClientWidth     =   9060
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBC23BrowseBarangTarifFasilias.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6180
   ScaleWidth      =   9060
   Begin VB.TextBox txtNoPengajuan 
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   5640
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox txNoSeri 
      Height          =   375
      Left            =   6000
      TabIndex        =   2
      Top             =   5640
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H0080FFFF&
      Caption         =   "Close"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5640
      Width           =   975
   End
   Begin VSFlex8Ctl.VSFlexGrid grid 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Tag             =   "TTTT*/"
      Top             =   120
      Width           =   8805
      _cx             =   15531
      _cy             =   9551
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
Attribute VB_Name = "frmBC23BrowseBarangTarifFasilitas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-------------------------------------------
Const colNo As Integer = 0
Const colJenisPungutan As Integer = 1
Const colDitangguhkan As Integer = 2
Const colDibebaskan As Integer = 3
Const colTidakDipungut As Integer = 4
Const colcount As Integer = 5

Private Sub up_GridHeader()
    With Grid
        .ColS = colcount
        .Rows = 1

        .TextMatrix(0, colNo) = "No"
        .TextMatrix(0, colJenisPungutan) = "Jenis Pungutan"
        .TextMatrix(0, colDitangguhkan) = "Ditangguhkan"
        .TextMatrix(0, colDibebaskan) = "Dibebaskan"
        .TextMatrix(0, colTidakDipungut) = "Tidak Dipungut"
        
        .ColWidth(colNo) = 500
        .ColWidth(colJenisPungutan) = 2500
        .ColWidth(colDitangguhkan) = 1700
        .ColWidth(colDibebaskan) = 1700
        .ColWidth(colTidakDipungut) = 1700
       
        .ColFormat(colJenisPungutan) = "#,0.00"
        .ColFormat(colDitangguhkan) = "#,0.00"
        .ColFormat(colDibebaskan) = "#,0.00"
        .ColFormat(colTidakDipungut) = "#,0.00"
        
        
        .MergeCells = flexMergeRestrictRows
        .WordWrap = True
        
        .AllowUserResizing = flexResizeColumns
        
    End With
End Sub

Private Sub up_GridLoad()
    Dim RS As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim li_Row As Integer
    Dim i As Integer
    
    Dim NilaiDitangguhkan As Double
    Dim NilaiDibebaskan As Double
    Dim NilaiTidakDipungut As Double
    
    up_GridHeader
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC23TPBTarifFasilitasPerBarang_Sel"

    cmd.Parameters.append cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan.Text)
    cmd.Parameters.append cmd.CreateParameter("NoSeri", adInteger, adParamInput, 5, txNoSeri.Text)
    
     Set RS = cmd.Execute

    With Grid
        While Not RS.EOF
            .Rows = .Rows + 1
            li_Row = .Rows - 1
            
            i = i + 1
            
            .TextMatrix(li_Row, colNo) = i
            .TextMatrix(li_Row, colJenisPungutan) = IIf(IsNull(RS.Fields("Kode_Pungutan")), "", RS.Fields("Kode_Pungutan"))
            .TextMatrix(li_Row, colDitangguhkan) = IIf(IsNull(RS.Fields("NILAIDITANGGUHKAN")), 0, RS.Fields("NILAIDITANGGUHKAN"))
            .TextMatrix(li_Row, colDibebaskan) = IIf(IsNull(RS.Fields("NILAIDIBEBASKAN")), 0, RS.Fields("NILAIDIBEBASKAN"))
            .TextMatrix(li_Row, colTidakDipungut) = IIf(IsNull(RS.Fields("NILAITIDAKDIPUNGUT")), 0, RS.Fields("NILAITIDAKDIPUNGUT"))
            
            NilaiDitangguhkan = NilaiDitangguhkan + CDbl(.TextMatrix(li_Row, colDitangguhkan))
            NilaiDibebaskan = NilaiDibebaskan + CDbl(.TextMatrix(li_Row, colDibebaskan))
            NilaiTidakDipungut = NilaiTidakDipungut + CDbl(.TextMatrix(li_Row, colTidakDipungut))
            
            RS.MoveNext
        Wend
        
        .Rows = .Rows + 1
        li_Row = .Rows - 1

        
        .TextMatrix(li_Row, colDitangguhkan) = NilaiDitangguhkan
        .TextMatrix(li_Row, colDibebaskan) = NilaiDibebaskan
        .TextMatrix(li_Row, colTidakDipungut) = NilaiTidakDipungut
        
        .Cell(flexcpText, li_Row, colNo, li_Row, colJenisPungutan) = "TOTAL"
        .Cell(flexcpFontBold, li_Row, colNo, li_Row, colTidakDipungut) = True
        
        Const ClrTotal1 = &HFFC0C0
        .Cell(flexcpBackColor, li_Row, colNo, Grid.Rows - 1, colTidakDipungut) = ClrTotal1  '&HFFC0C0
        
        .MergeRow(li_Row) = True
        
        RS.Close
        Set RS = Nothing
    End With
            
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
up_GridLoad
End Sub

