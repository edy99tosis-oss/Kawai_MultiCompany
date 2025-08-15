VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Begin VB.Form frmBC23BrowseBarangDokumen 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Barang Dokumen"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13020
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBC23BrowseBarangDokumen.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   13020
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNoSeri 
      Height          =   375
      Left            =   9960
      TabIndex        =   11
      Top             =   4920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   240
      TabIndex        =   9
      Tag             =   "TFTT*/"
      Top             =   4320
      Width           =   12495
      Begin VB.Label LblerrMsg 
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
         Left            =   105
         TabIndex        =   10
         Tag             =   "TTFF*/"
         Top             =   195
         Width           =   12210
      End
   End
   Begin VB.TextBox txtKodeBarang 
      Height          =   375
      Left            =   6960
      TabIndex        =   8
      Top             =   4920
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox txtNoPengajuan 
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   4920
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H0080FFFF&
      Caption         =   "Close"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H0080FFFF&
      Caption         =   "Delete"
      Height          =   375
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H0080FFFF&
      Caption         =   "Add"
      Height          =   375
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VSFlex8Ctl.VSFlexGrid gridSumber 
      Height          =   3615
      Left            =   240
      TabIndex        =   0
      TabStop         =   0   'False
      Tag             =   "TTTT*/"
      Top             =   600
      Width           =   6045
      _cx             =   10663
      _cy             =   6376
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
   Begin VSFlex8Ctl.VSFlexGrid gridDokumen 
      Height          =   3615
      Left            =   6720
      TabIndex        =   1
      TabStop         =   0   'False
      Tag             =   "TTTT*/"
      Top             =   600
      Width           =   6045
      _cx             =   10663
      _cy             =   6376
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SUMBER DATA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   38
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATA DOKUMEN"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   40
      Left            =   6720
      TabIndex        =   2
      Top             =   120
      Width           =   1515
   End
End
Attribute VB_Name = "frmBC23BrowseBarangDokumen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------
Const colNoSumber As Integer = 0
Const colKodeDokumenSumber As Integer = 1
Const colJenisDokumenSumber As Integer = 2
Const colNomorDokumenSumber As Integer = 3
Const colTanggalSumber As Integer = 4
Const colCountSumber As Integer = 5

'-------------------------------------------
Const colNoData As Integer = 0
Const colKodeDokumenData As Integer = 1
Const colJenisDokumenData As Integer = 2
Const colNomorDokumenData As Integer = 3
Const colTanggalData As Integer = 4
Const colCountData As Integer = 5

Private Sub up_GridHeaderSumber()
    
    With gridSumber
        .ColS = colCountSumber
        .Rows = 1
        
        .TextMatrix(0, colNoSumber) = "No"
        .TextMatrix(0, colKodeDokumenSumber) = "Kode"
        .TextMatrix(0, colJenisDokumenSumber) = "Jenis"
        .TextMatrix(0, colNomorDokumenSumber) = "Nomor"
        .TextMatrix(0, colTanggalSumber) = "Tanggal"
        
        .ColWidth(colNoSumber) = 500
        .ColWidth(colKodeDokumenSumber) = 1000
        .ColWidth(colJenisDokumenSumber) = 1600
        .ColWidth(colNomorDokumenSumber) = 1600
        .ColWidth(colTanggalSumber) = 1200
        .ColAlignment(colNomorDokumenSumber) = flexAlignLeftCenter
        
        .ColFormat(colTanggalSumber) = "dd MMM yyyy"
        
        .FrozenCols = 2
    End With
End Sub

Private Sub up_GridHeaderDokumen()
    
    With gridDokumen
        .ColS = colCountData
        .Rows = 1
        
        .TextMatrix(0, colNoData) = "No"
        .TextMatrix(0, colKodeDokumenData) = "Kode"
        .TextMatrix(0, colJenisDokumenData) = "Jenis"
        .TextMatrix(0, colNomorDokumenData) = "Nomor"
        .TextMatrix(0, colTanggalData) = "Tanggal"
        
        .ColWidth(colNoData) = 500
        .ColWidth(colKodeDokumenData) = 1000
        .ColWidth(colJenisDokumenData) = 1600
        .ColWidth(colNomorDokumenData) = 1600
        .ColWidth(colTanggalData) = 1200
        .ColAlignment(colNomorDokumenData) = flexAlignLeftCenter
        
        .ColFormat(colTanggalData) = "dd MMM yyyy"
        
        .FrozenCols = 2
    End With
End Sub

Private Sub up_GridLoadSumber()
Dim RS As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim li_Row As Integer
    Dim i As Integer
    
    up_GridHeaderSumber
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC23TPBDokumen_Sel"
    
    cmd.Parameters.append cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan.Text)
    
    
     Set RS = cmd.Execute
     
    With gridSumber
        While Not RS.EOF
            .Rows = .Rows + 1
            li_Row = .Rows - 1
            i = i + 1
            
            .TextMatrix(li_Row, colNoSumber) = Trim(RS!Seri_Dokumen)
            .TextMatrix(li_Row, colKodeDokumenSumber) = Trim(RS!Kode_Jenis_Dokumen)
            .TextMatrix(li_Row, colJenisDokumenSumber) = Trim(RS!Uraian_Dokumen)
            .TextMatrix(li_Row, colNomorDokumenSumber) = Trim(RS!Nomor_Dokumen)
            .TextMatrix(li_Row, colTanggalSumber) = Trim(RS!Tanggal_Dokumen)
            
            
            RS.MoveNext
        Wend
        RS.Close
        Set RS = Nothing
    End With
End Sub

Private Sub up_GridLoadDokumen()
Dim RS As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim li_Row As Integer
    Dim i As Integer
    
    up_GridHeaderDokumen
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC23DetailBarangDokumenPerBarang_Sel"
    
    cmd.Parameters.append cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan.Text)
    cmd.Parameters.append cmd.CreateParameter("NoSeri", adVarChar, adParamInput, 5, txtNoSeri.Text)
    
     Set RS = cmd.Execute
     
    With gridDokumen
        While Not RS.EOF
            .Rows = .Rows + 1
            li_Row = .Rows - 1
            i = i + 1
            
            .TextMatrix(li_Row, colNoData) = Trim(RS!Seri_Dokumen)
            .TextMatrix(li_Row, colKodeDokumenData) = Trim(RS!Kode_Jenis_Dokumen)
            .TextMatrix(li_Row, colJenisDokumenData) = Trim(RS!Uraian_Dokumen)
            .TextMatrix(li_Row, colNomorDokumenData) = Trim(RS!Nomor_Dokumen)
            .TextMatrix(li_Row, colTanggalData) = Trim(RS!Tanggal_Dokumen)
            
            RS.MoveNext
        Wend
        RS.Close
        Set RS = Nothing
    End With
End Sub

Private Sub up_AddDokumen()
    Dim RS As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim NoDokumen As String
    
    NoDokumen = Trim(gridSumber.TextMatrix(gridSumber.RowSel, colNomorDokumenData))
    
    For i = 1 To gridDokumen.Rows - 1
        If NoDokumen = Trim(gridDokumen.TextMatrix(i, colNomorDokumenData)) Then
            LblerrMsg.Caption = "Nomor Dokumen already exists!"
            Exit Sub
        End If
    Next
            
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC23DetailBarangDokumen_Ins"
    
    cmd.Parameters.append cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
    cmd.Parameters.append cmd.CreateParameter("KodeBarang", adVarChar, adParamInput, 50, txtKodeBarang)
    cmd.Parameters.append cmd.CreateParameter("NoSeriBarang", adInteger, adParamInput, 5, txtNoSeri)
    cmd.Parameters.append cmd.CreateParameter("NoSeriDokumen", adInteger, adParamInput, 5, gridSumber.TextMatrix(gridSumber.RowSel, colNoSumber))
    
    cmd.Execute
        
    up_GridLoadDokumen
End Sub

Private Sub cmdAdd_Click()
    If gridSumber.Rows = 1 Then Exit Sub
    up_AddDokumen
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    If gridDokumen.Rows = 1 Then Exit Sub
    
    Dim RS As ADODB.Recordset
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC23DetailBarangDokumen_Del"
    
    cmd.Parameters.append cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
    cmd.Parameters.append cmd.CreateParameter("NoSeriBarang", adInteger, adParamInput, 5, txtNoSeri)
    cmd.Parameters.append cmd.CreateParameter("NoSeriDokumen", adInteger, adParamInput, 5, gridDokumen.TextMatrix(gridDokumen.RowSel, colNoData))
    
    cmd.Execute
        
    up_GridLoadDokumen
End Sub

Private Sub Form_Activate()
   LblerrMsg.Caption = ""
   up_GridLoadSumber
   up_GridLoadDokumen
End Sub

