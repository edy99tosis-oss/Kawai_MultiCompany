VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBC27BrowseDokumen 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Browse Dokumen"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8835
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBC27BrowseDokumen.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   8835
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H0080FFFF&
      Caption         =   "Close"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton cmdSubmit 
      BackColor       =   &H0080FFFF&
      Caption         =   "Submit"
      Height          =   375
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H0080FFFF&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H0080FFFF&
      Caption         =   "Delete"
      Height          =   375
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox txtKodeDokumen 
      Height          =   350
      Left            =   360
      MaxLength       =   5
      TabIndex        =   7
      Top             =   4440
      Width           =   1305
   End
   Begin VB.TextBox txtNamaDokumen 
      Height          =   350
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   4440
      Width           =   2625
   End
   Begin VB.TextBox txtNomorDokumen 
      Height          =   350
      Left            =   4560
      MaxLength       =   100
      TabIndex        =   5
      Top             =   4440
      Width           =   2235
   End
   Begin VB.TextBox txtNoAju 
      Height          =   350
      Left            =   2880
      TabIndex        =   4
      Top             =   5880
      Visible         =   0   'False
      Width           =   2625
   End
   Begin VB.TextBox txtTipe 
      Height          =   350
      Left            =   1920
      MaxLength       =   2
      TabIndex        =   3
      Top             =   5880
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.TextBox txtID 
      Height          =   350
      Left            =   3240
      MaxLength       =   5
      TabIndex        =   2
      Top             =   5880
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   240
      TabIndex        =   0
      Tag             =   "TFTT*/"
      Top             =   5160
      Width           =   8355
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
         TabIndex        =   1
         Tag             =   "TTFF*/"
         Top             =   195
         Width           =   8130
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid grid 
      Height          =   3615
      Left            =   240
      TabIndex        =   12
      TabStop         =   0   'False
      Tag             =   "TTTT*/"
      Top             =   240
      Width           =   8325
      _cx             =   14684
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
   Begin MSComCtl2.DTPicker dtpTglDokumen 
      Height          =   315
      Left            =   6960
      TabIndex        =   13
      Tag             =   "TTFF*/"
      Top             =   4440
      Width           =   1500
      _ExtentX        =   2646
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
      CustomFormat    =   "dd MMM yyyy"
      Format          =   174653443
      CurrentDate     =   37798
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00A6D2FF&
      Height          =   555
      Index           =   2
      Left            =   240
      Top             =   4320
      Width           =   8325
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kode"
      Height          =   195
      Index           =   1
      Left            =   375
      TabIndex        =   17
      Top             =   4035
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nomor"
      Height          =   195
      Index           =   0
      Left            =   4560
      TabIndex        =   16
      Top             =   4035
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
      Height          =   195
      Index           =   2
      Left            =   6960
      TabIndex        =   15
      Top             =   4035
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dokumen"
      Height          =   195
      Index           =   4
      Left            =   1800
      TabIndex        =   14
      Top             =   4035
      Width           =   825
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00A6D2FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   2
      Left            =   240
      Top             =   3960
      Width           =   8325
   End
End
Attribute VB_Name = "frmBC27BrowseDokumen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------
Const colSeri As Integer = 0
Const colKodeDokumen As Integer = 1
Const colJenisDokumen As Integer = 2
Const colNomorDokumen As Integer = 3
Const colTanggal As Integer = 4
Const colHideID As Integer = 5
Const colHideTipe As Integer = 6
Const colCountDokumen As Integer = 7


Private Sub up_Clear()
    txtKodeDokumen.Enabled = True
    LblerrMsg.Caption = ""
    txtID = ""
    txtKodeDokumen = ""
    txtNamaDokumen = ""
    txtNomorDokumen = ""
    dtpTglDokumen.Value = Now
    txtTipe = ""
    
 
End Sub

Private Function uf_ValidateInput() As Boolean

    If txtKodeDokumen.Text = "" Or txtNamaDokumen = "" Then
        txtKodeDokumen.SetFocus
        LblerrMsg = "Please Input Kode Dokumen!"
        uf_ValidateInput = False
        Exit Function
    ElseIf txtNomorDokumen.Text = "" Then
        txtNomorDokumen.SetFocus
        LblerrMsg = "Please Input Nomor Dokumen!"
        uf_ValidateInput = False
        Exit Function
    End If
    
    uf_ValidateInput = True
End Function

Private Sub up_LoadDokumen(pKode As String)
Dim sql As String
Dim RS As New Recordset

sql = "Select Uraian_Dokumen, Tipe_Dokumen From Bea_Cukai_Dokumen Where Kode_Dokumen = '" & pKode & "'"
Set RS = Db.Execute(sql)
    
If Not RS.EOF Then
    txtNamaDokumen = RS.Fields("Uraian_Dokumen")
    txtTipe = RS.Fields("Tipe_Dokumen")
Else
    txtNamaDokumen = ""
    txtTipe = ""
End If
End Sub

Private Sub up_GridHeaderDokumen()
    
    With Grid
        .ColS = colCountDokumen
        .Rows = 1

        .TextMatrix(0, colSeri) = "Seri"
        .TextMatrix(0, colKodeDokumen) = "Kode Dokumen"
        .TextMatrix(0, colJenisDokumen) = "Jenis Dokumen"
        .TextMatrix(0, colNomorDokumen) = "Nomor"
        .TextMatrix(0, colTanggal) = "Tanggal"
        
        .ColWidth(colSeri) = 700
        .ColWidth(colKodeDokumen) = 1500
        .ColWidth(colJenisDokumen) = 2000
        .ColWidth(colNomorDokumen) = 2000
        .ColWidth(colTanggal) = 1200
        .ColAlignment(colNomorDokumen) = flexAlignLeftCenter
        
        .ColFormat(colTanggal) = "dd MMM yyyy"
        .ColHidden(colHideID) = True
        .ColHidden(colHideTipe) = True
    End With
End Sub

Public Sub up_GridLoad()
Dim RS As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim li_Row As Integer

    up_GridHeaderDokumen
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC25TPBDokumen_Sel"
    
    cmd.Parameters.append cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoAju.Text)
     Set RS = cmd.Execute
     
    With Grid
        While Not RS.EOF
            .Rows = .Rows + 1
            li_Row = .Rows - 1

            .TextMatrix(li_Row, colSeri) = Trim(RS!Seri_Dokumen)
            .TextMatrix(li_Row, colKodeDokumen) = Trim(RS!Kode_Jenis_Dokumen)
            .TextMatrix(li_Row, colJenisDokumen) = Trim(RS!Uraian_Dokumen)
            .TextMatrix(li_Row, colNomorDokumen) = Trim(RS!Nomor_Dokumen)
            .TextMatrix(li_Row, colTanggal) = Trim(RS!Tanggal_Dokumen)
            .TextMatrix(li_Row, colHideID) = Trim(RS!ID_Dokumen)
            .TextMatrix(li_Row, colHideTipe) = Trim(RS!Tipe_Dokumen)
            RS.MoveNext
        Wend
        RS.Close
        Set RS = Nothing
    End With
End Sub

Private Sub up_SaveData()
    Dim RS As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim Y As Integer
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC25TPBDokumen_Upd"
    
    If txtID = "" Then txtID = "0"
    
    cmd.Parameters.append cmd.CreateParameter("KodeDokumen", adVarChar, adParamInput, 5, txtKodeDokumen)
    cmd.Parameters.append cmd.CreateParameter("NomorDokumen", adVarChar, adParamInput, 100, txtNomorDokumen)
    cmd.Parameters.append cmd.CreateParameter("TanggalDokumen", adDate, adParamInput, , Format(dtpTglDokumen.Value, "yyyy-MM-dd"))
    cmd.Parameters.append cmd.CreateParameter("TipeDokumen", adVarChar, adParamInput, 2, txtTipe)
    cmd.Parameters.append cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoAju)
    cmd.Parameters.append cmd.CreateParameter("IDDokumen", adInteger, adParamInput, , txtID)
    
    cmd.Execute Y
    
    If Y = 0 Then
        Set cmd = New ADODB.Command
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 0
        cmd.ActiveConnection = Db
        cmd.CommandText = "sp_BC25TPBDokumen_Ins"
        
        cmd.Parameters.append cmd.CreateParameter("KodeDokumen", adVarChar, adParamInput, 5, txtKodeDokumen)
        cmd.Parameters.append cmd.CreateParameter("NomorDokumen", adVarChar, adParamInput, 100, txtNomorDokumen)
        cmd.Parameters.append cmd.CreateParameter("TanggalDokumen", adDate, adParamInput, , Format(dtpTglDokumen.Value, "yyyy-MM-dd"))
        cmd.Parameters.append cmd.CreateParameter("TipeDokumen", adVarChar, adParamInput, 2, txtTipe)
        cmd.Parameters.append cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoAju)
        
        cmd.Execute
    End If
        
    up_Clear
    up_GridLoad
    
    If Y = 0 Then
        LblerrMsg = DisplayMsg(1000)
    Else
        LblerrMsg = DisplayMsg(1101)
    End If
End Sub

Private Sub up_Delete()
    Dim RS As ADODB.Recordset
    Dim cmd As ADODB.Command
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC25TPBDokumen_Del"
    
    cmd.Parameters.append cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoAju)
    cmd.Parameters.append cmd.CreateParameter("IDDokumen", adInteger, adParamInput, , txtID)
    cmd.Execute
    
    up_Clear
    up_GridLoad
    
    LblerrMsg.Caption = DisplayMsg(1201)
End Sub

Private Sub btnClose_Click()
Unload Me
frmBC25Detail.Show
End Sub

Private Sub cmdCancel_Click()
    up_Clear
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
    If txtID = "" Then
        LblerrMsg = DisplayMsg("8011")
        Exit Sub
    End If
    If MsgBox("Are you sure want to delete?", vbYesNo + vbExclamation, "Delete") = vbYes Then
        up_Delete
    End If
End Sub

Private Sub CmdSubmit_Click()
    If uf_ValidateInput = False Then Exit Sub
    up_SaveData
End Sub

Private Sub Form_Load()
    up_Clear
    up_GridLoad
End Sub

Private Sub grid_Click()
If Grid.RowSel > 0 Then
    txtKodeDokumen.Enabled = False
    txtKodeDokumen = Grid.TextMatrix(Grid.RowSel, colKodeDokumen)
    txtNamaDokumen = Grid.TextMatrix(Grid.RowSel, colJenisDokumen)
    txtNomorDokumen = Grid.TextMatrix(Grid.RowSel, colNomorDokumen)
    dtpTglDokumen = CDate(Grid.TextMatrix(Grid.RowSel, colTanggal))
    txtID = Grid.TextMatrix(Grid.RowSel, colHideID)
    txtTipe = Grid.TextMatrix(Grid.RowSel, colHideTipe)
End If
'txtCost.Text = .TextMatrix(Row, bteColCostCls)
'Txttitle.Text = .TextMatrix(Row, bteColCostTittle)
'
'txtdesc.Text = .TextMatrix(Row, bteColDesc)
End Sub

Private Sub txtKodeDokumen_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtKodeDokumen_LostFocus()
up_LoadDokumen txtKodeDokumen
End Sub

Private Sub txtNomorDokumen_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

