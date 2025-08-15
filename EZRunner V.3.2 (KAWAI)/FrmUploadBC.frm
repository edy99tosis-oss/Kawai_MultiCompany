VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmUploadBC 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Upload BC"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13740
   Icon            =   "FrmUploadBC.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8565
   ScaleWidth      =   13740
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00FDDFE3&
      Height          =   480
      Left            =   120
      TabIndex        =   18
      Top             =   7560
      Width           =   13485
      Begin VB.Label lblErrMsg 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Left            =   60
         TabIndex        =   19
         Top             =   150
         Width           =   13020
      End
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Back"
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   8160
      Width           =   1125
   End
   Begin VB.CommandButton CmdSubmit 
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
      Left            =   12465
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8160
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDDFE3&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   13485
      Begin VB.TextBox Lblsupp 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDDFE3&
         BorderStyle     =   0  'None
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
         Left            =   3615
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Width           =   3510
      End
      Begin VB.TextBox txtErr 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9960
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "0"
         Top             =   1365
         Width           =   1695
      End
      Begin VB.CommandButton cmdLocation 
         BackColor       =   &H0080FFFF&
         Caption         =   "...."
         Height          =   375
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1320
         Width           =   660
      End
      Begin VB.CommandButton cmdtemplate 
         BackColor       =   &H0080FFFF&
         Caption         =   "Template"
         Height          =   375
         Left            =   6600
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1320
         Width           =   1140
      End
      Begin VB.TextBox Text1 
         Height          =   360
         Left            =   1980
         TabIndex        =   1
         Top             =   1320
         Width           =   3705
      End
      Begin MSComCtl2.DTPicker Tgl1 
         Height          =   345
         Left            =   1980
         TabIndex        =   11
         Top             =   780
         Width           =   1515
         _ExtentX        =   2672
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
         Format          =   174718979
         CurrentDate     =   37868
      End
      Begin MSComCtl2.DTPicker Tgl2 
         Height          =   345
         Left            =   3975
         TabIndex        =   12
         Top             =   780
         Width           =   1515
         _ExtentX        =   2672
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
         Format          =   174718979
         CurrentDate     =   37868
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "to"
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
         Left            =   3675
         TabIndex        =   13
         Top             =   855
         Width           =   165
      End
      Begin VB.Label LblPart 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Receipt Date"
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
         Left            =   240
         TabIndex        =   10
         Top             =   840
         Width           =   1095
      End
      Begin VB.Line Line4 
         X1              =   3615
         X2              =   7115
         Y1              =   555
         Y2              =   555
      End
      Begin MSForms.ComboBox CboPart 
         Height          =   315
         Index           =   0
         Left            =   1980
         TabIndex        =   9
         Top             =   240
         Width           =   1515
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2672;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label LblPart 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Code "
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
         Left            =   240
         TabIndex        =   8
         Top             =   300
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invalid Surat Jalan"
         Height          =   195
         Index           =   3
         Left            =   8160
         TabIndex        =   6
         Top             =   1410
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Upload Location File"
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
         Index           =   6
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   1695
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid grid 
      Height          =   5160
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   13485
      _cx             =   23786
      _cy             =   9102
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
      GridColor       =   12582912
      GridColorFixed  =   12582912
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
      RowHeightMax    =   0
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
      Begin VB.TextBox txtlocation 
         Appearance      =   0  'Flat
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
         Left            =   5280
         MaxLength       =   35
         TabIndex        =   15
         Top             =   7440
         Width           =   2775
      End
   End
   Begin MSComDlg.CommonDialog cdg 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label LblTotalRec 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total : 0 Record (s)"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   11520
      TabIndex        =   20
      Tag             =   "TTFF*/"
      Top             =   7200
      Width           =   2100
   End
End
Attribute VB_Name = "FrmUploadBC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  
Dim sql As String, sqlGrid As String

Dim ls_PathExcel As String



Dim bteColSupplierCode As Byte
Dim bteColSupplierName As Byte
Dim bteColPONo As Byte
Dim bteColSuratJalanNo As Byte
Dim bteColReceiptDate As Byte
Dim bteColBctype As Byte
Dim bteColBCNo As Byte
Dim bteColBCDate As Byte
Dim bteColRemark As Byte

Private Sub CboPart_Change(Index As Integer)
    CboPart(0) = CboPart(0)
    lblSupp = ""
    Header
    If CboPart(0).MatchFound = True Then lblSupp = CboPart(0).List(CboPart(0).ListIndex, 1)
End Sub

Private Sub cmdBack_Click()

    DoEvents
    Unload Me
End Sub

Private Sub cmdLocation_Click()
Call Header
Call up_ImportOffline
End Sub

Private Sub CmdSubmit_Click()

Dim X As Double
Dim SupplierCode As String, PONO As String, SuratJalan As String, ReceiptDate As Date, _
                      BCType As String, BCNo As String, BCDate As Date

Me.MousePointer = vbHourglass


For X = 1 To Grid.Rows - 1

SupplierCode = Grid.TextMatrix(X, bteColSupplierCode)
PONO = Grid.TextMatrix(X, bteColPONo)
SuratJalan = Grid.TextMatrix(X, bteColSuratJalanNo)
ReceiptDate = Grid.TextMatrix(X, bteColReceiptDate)
BCType = Grid.TextMatrix(X, bteColBctype)
BCNo = Grid.TextMatrix(X, bteColBCNo)
If IsDate(Grid.TextMatrix(X, bteColBCDate)) Then
BCDate = Grid.TextMatrix(X, bteColBCDate)
End If
        
        
If Grid.Cell(flexcpBackColor, X, bteColSupplierCode, X, bteColRemark) <> &H8080FF Then
    
    
    Grid.TextMatrix(X, bteColRemark) = uf_Save(SupplierCode, PONO, SuratJalan, ReceiptDate, BCType, BCNo, BCDate)
    If Grid.TextMatrix(X, bteColRemark) <> "" Then
        Grid.Cell(flexcpBackColor, X, bteColSupplierCode, X, bteColRemark) = 16637923
    End If
    
End If
        
Next X
LblerrMsg.Caption = "Update Records Succes..!!"
Me.MousePointer = vbDefault
End Sub

Private Sub cmdtemplate_Click()
Call Header
Call up_ExportOffline
End Sub
Sub Header()
        
    Dim i As Integer
    
  
    bteColSupplierCode = 0
    bteColSupplierName = 1
    bteColPONo = 2
    bteColSuratJalanNo = 3
    bteColReceiptDate = 4
    bteColBctype = 5
    bteColBCNo = 6
    bteColBCDate = 7
    bteColRemark = 8
    
    With Grid
        .clear
        .Rows = 1
        .ColS = 9
        
        
        .TextMatrix(0, bteColSupplierCode) = "Supplier Code"
        .TextMatrix(0, bteColSupplierName) = "Supplier Name"
        .TextMatrix(0, bteColPONo) = "PO NO"
        .TextMatrix(0, bteColSuratJalanNo) = "Surat Jalan No"
        .TextMatrix(0, bteColReceiptDate) = "Receipt Date"
        .TextMatrix(0, bteColBctype) = "BC Type"
        .TextMatrix(0, bteColBCNo) = "BC No"
        .TextMatrix(0, bteColBCDate) = "BC Date"
        .TextMatrix(0, bteColRemark) = "Remarks"
        
        .ColWidth(bteColSupplierCode) = 1500
        .ColWidth(bteColSupplierName) = 2500
        .ColWidth(bteColPONo) = 2500
        .ColWidth(bteColSuratJalanNo) = 1500
        .ColWidth(bteColReceiptDate) = 1250
        .ColWidth(bteColBctype) = 1000
        .ColWidth(bteColBCNo) = 1500
        .ColWidth(bteColBCDate) = 1250
        .ColWidth(bteColRemark) = 3000
   
        
        .EditMaxLength = 1
    End With
End Sub
Private Sub up_ExportOffline()
Dim objExcel As New Excel.application
Dim RS As New ADODB.Recordset
Dim strSQL As String
Dim i As Double
Dim adoCmd As New Command
Dim rsCheck As New ADODB.Recordset



On Error GoTo errHandler

    If G_CekExcelApp = False Then LblerrMsg.Caption = "Excel Application is not found !": Exit Sub

    LblerrMsg.Caption = ""
    cdg.filter = "Excel Worksheets 2003 (*.xls)|*.xls|"
    cdg.filename = "BC Upload "
    cdg.CancelError = True

On Error GoTo errCancel
    cdg.ShowSave

   
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

        .Range("A1").Value = "Supplier Code"
        .Range("B1").Value = "Supplier Name"
        .Range("C1").Value = "PO No"
        .Range("D1").Value = "Surat Jalan No"
        .Range("E1").Value = "Receipt Date"
        .Range("F1").Value = "BC Type"
        .Range("G1").Value = "BC No"
        .Range("H1").Value = "BC Date"


        'Get data BC masih Kosong
        
        Dim Trade As String
        Dim StartDate As Date
        Dim EndDate As Date
        
        Trade = Trim(CboPart(0).Text)
        StartDate = Format(Tgl1.Value, "yyyy-mm-dd")
        EndDate = Format(Tgl2.Value, "yyyy-mm-dd")
        
        adoCmd.ActiveConnection = Db.ConnectionString
        adoCmd.CommandTimeout = 120
        adoCmd.CommandType = adCmdStoredProc
        adoCmd.CommandText = "SP_GetBCUpload"
        adoCmd.Parameters(1) = Trade
        adoCmd.Parameters(2) = StartDate
        adoCmd.Parameters(3) = EndDate
        
        Set rsCheck = adoCmd.Execute

        If Not rsCheck.EOF Then
            Dim iRowExl As Integer
            iRowExl = 2
           While Not rsCheck.EOF
            
                .Range("A" & iRowExl) = rsCheck("Supplier_Code")
            
                .Range("B" & iRowExl) = rsCheck("Trade_Name")
            
                .Range("C" & iRowExl) = rsCheck("PO_No")
            
                .Range("D" & iRowExl) = rsCheck("SuratJalan_No")
                
                .Range("E" & iRowExl) = Format(rsCheck("Receipt_Date"), "yyyy-mm-dd")
            
                .Range("F" & iRowExl) = IIf(IsNull(rsCheck("BC_Type")), "", rsCheck("BC_Type"))
            
                .Range("G" & iRowExl) = IIf(IsNull(rsCheck("BC40_No")), "", rsCheck("BC40_No"))
            
                .Range("H" & iRowExl) = IIf(IsNull(rsCheck("BC40_Date")), "", rsCheck("BC40_Date"))
                
              
                
                 iRowExl = iRowExl + 1
                rsCheck.MoveNext
            Wend
            
                    
        Else
            LblerrMsg.Caption = "There is no Data to Upload..!!!"
        
        End If


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
        LblerrMsg.Caption = err.Description
        Grid.FixedRows = 1
    End If
    If RS.State = adStateOpen Then
        RS.Close
        Set RS = Nothing
    End If
errCancel:
Me.MousePointer = vbDefault

End Sub

Private Sub Form_Load()
Tgl1.Value = Now
Tgl2.Value = Now

Call Header

Dim rst As New Recordset
Dim SQLT As String
Dim ir As Integer

'##Tampilkan Combo Customer code dari trade_master
SQLT = "Select rtrim(trade_code) as TC,Trade_name as TN from trade_master where trade_cls in ('2', '3') order by trade_code"

Set rst = New Recordset
rst.Open SQLT, Db, adOpenKeyset, adLockOptimistic
CboPart(0).clear
CboPart(0).ColumnCount = 2
CboPart(0).TextColumn = 1
CboPart(0).AddItem ""
CboPart(0).List(0, 0) = "ALL"
CboPart(0).List(0, 1) = "ALL"

ir = 1
While Not rst.EOF
    CboPart(0).AddItem ""
    CboPart(0).List(ir, 0) = rst!TC
    CboPart(0).List(ir, 1) = Trim$(rst!TN)
    ir = ir + 1
    rst.MoveNext
Wend
CboPart(0).ColumnWidths = "60 pt; 300 pt"
CboPart(0).ListWidth = 360
CboPart(0).ListRows = 15
CboPart(0).ListIndex = 0

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
    

    If G_CekExcelApp = False Then LblerrMsg.Caption = "Excel Application is not found": Exit Sub
    
    LblerrMsg.Caption = ""
    'cdg.Filter = "Excel Files (*.xls)|*.xls"
    cdg.filter = "Excel Worksheets (*.xls)|*.xls|"
    
    
    cdg.filename = ""
    
    On Error GoTo errCancel
    cdg.CancelError = True
    
    On Error GoTo err
    
    cdg.ShowOpen
    txtlocation.Text = cdg.filename
    txtErr.Text = 0
    If cdg.filename <> "" Then

        txtErr.Text = 0
        Me.MousePointer = vbHourglass
        Set objExcel = New Excel.application
        Set objWorkBook = objExcel.Workbooks.Open(cdg.filename)
        Set objWorkSheet = objWorkBook.Sheets("Sheet1")
        objExcel.Visible = False
        i = 2
        iGrdRow = 1
        colcount = 22
        With objWorkSheet
        
            Do While .Cells(i, 1).Value <> ""
                        Grid.AddItem ""
                        
              
                        
                        Grid.TextMatrix(iGrdRow, bteColSupplierCode) = Trim(.Cells(i, 1))
                        Grid.TextMatrix(iGrdRow, bteColSupplierName) = Trim(.Cells(i, 2))
                        Grid.TextMatrix(iGrdRow, bteColPONo) = Trim(.Cells(i, 3))
                        Grid.TextMatrix(iGrdRow, bteColSuratJalanNo) = Trim(.Cells(i, 4))
                        Grid.TextMatrix(iGrdRow, bteColReceiptDate) = Trim(.Cells(i, 5))
                        Grid.TextMatrix(iGrdRow, bteColBctype) = Trim(.Cells(i, 6))
                        Grid.TextMatrix(iGrdRow, bteColBCNo) = Trim(.Cells(i, 7))
                        Grid.TextMatrix(iGrdRow, bteColBCDate) = Trim(.Cells(i, 8))
                        
                
                        If Trim(.Cells(i, 6)) = "" Then
                            Grid.TextMatrix(iGrdRow, bteColRemark) = "Invalid BC Type !!"
                            Grid.Cell(flexcpBackColor, iGrdRow, bteColSupplierCode, iGrdRow, bteColRemark) = &H8080FF
                        ElseIf Trim(.Cells(i, 7)) = "" Then
                            Grid.TextMatrix(iGrdRow, bteColRemark) = "Invalid BC NO !!"
                            Grid.Cell(flexcpBackColor, iGrdRow, bteColSupplierCode, iGrdRow, bteColRemark) = &H8080FF
                            
                        Else
                            If IsDate(.Cells(i, 8)) Then
                        
                            Else
                            
                                Grid.TextMatrix(iGrdRow, bteColRemark) = "Invalid BC Date Format (YYYY-MM-DD) "
                                Grid.Cell(flexcpBackColor, iGrdRow, bteColSupplierCode, iGrdRow, bteColRemark) = &H8080FF
                            End If
                            
                        End If
                                                                         
                        
                        
                        
                        iGrdRow = iGrdRow + 1
               
                    
                    
                LblerrMsg = "Reading row : " & i - 1
                DoEvents
                i = i + 1
            Loop
            
        End With
        
        ' clean object excel
        objExcel.Workbooks.Close
        Set objWorkSheet = Nothing
        
        Set objExcel = Nothing

        LblTotalRec = "Total : " & Grid.Rows - 1 & " record (s)"
        LblerrMsg.Caption = "Reading Excel finish"
        
        Me.MousePointer = vbDefault
    End If
    Exit Sub
errCancel:
err:
    LblerrMsg.Caption = err.Description
    objExcel.Workbooks.Close
    Set objWorkBook = Nothing
    Set objWorkSheet = Nothing
    Set objWorkBook = Nothing
    Set objExcel = Nothing
    Me.MousePointer = vbDefault
    
End Sub

Private Function uf_Save(SupplierCode As String, PONO As String, SuratJalan As String, ReceiptDate As Date, _
                      BCType As String, BCNo As String, BCDate As Date) As String
        
       On Error GoTo err
        
        Dim adoCmd As New Command


        adoCmd.ActiveConnection = Db.ConnectionString
        adoCmd.CommandTimeout = 120
        adoCmd.CommandType = adCmdStoredProc
        adoCmd.CommandText = "SP_UpdateBC_Upload"
        adoCmd.Parameters(1) = SupplierCode
        adoCmd.Parameters(2) = PONO
        adoCmd.Parameters(3) = SuratJalan
        adoCmd.Parameters(4) = ReceiptDate
        adoCmd.Parameters(5) = BCType
        adoCmd.Parameters(6) = BCNo
        adoCmd.Parameters(7) = BCDate
        
        adoCmd.Execute
        uf_Save = ""
        Exit Function
                        
err:
      uf_Save = err.Description
      'lblErrMsg.Caption = err.Description
      Set adoCmd = Nothing
      Me.MousePointer = vbDefault
        


End Function

