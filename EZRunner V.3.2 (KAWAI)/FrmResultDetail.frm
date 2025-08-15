VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FrmProdResultDetail 
   BackColor       =   &H00FDDFE3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Production Result Detail "
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9015
   Icon            =   "FrmResultDetail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   9015
   StartUpPosition =   2  'CenterScreen
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
      Left            =   7860
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7080
      Width           =   1035
   End
   Begin VB.CommandButton CmdScan 
      BackColor       =   &H0080FFFF&
      Caption         =   "Scan Bar&code (F2)"
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
      Left            =   5220
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7080
      Width           =   2565
   End
   Begin VB.CommandButton Cmd_Back 
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
      Index           =   8
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7080
      Width           =   1125
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   1170
      Width           =   8790
      Begin VB.TextBox LblName 
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
         Height          =   195
         Index           =   3
         Left            =   6105
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "LblLocationName"
         Top             =   630
         Width           =   2490
      End
      Begin VB.TextBox LblName 
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
         Height          =   195
         Index           =   2
         Left            =   6105
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "LblLocationName"
         Top             =   270
         Width           =   2490
      End
      Begin VB.TextBox LblName 
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
         Height          =   195
         Index           =   1
         Left            =   1395
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "LblLocationName"
         Top             =   600
         Width           =   3180
      End
      Begin VB.TextBox LblName 
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
         Height          =   195
         Index           =   0
         Left            =   1395
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "LblLocationName"
         Top             =   240
         Width           =   3180
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Line Code"
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
         Left            =   4800
         TabIndex        =   10
         Top             =   630
         Width           =   855
      End
      Begin VB.Line Line4 
         X1              =   6105
         X2              =   8580
         Y1              =   870
         Y2              =   870
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Result Date"
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
         Left            =   4800
         TabIndex        =   8
         Top             =   270
         Width           =   990
      End
      Begin VB.Line Line3 
         X1              =   6105
         X2              =   8610
         Y1              =   510
         Y2              =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ware House"
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
         Left            =   90
         TabIndex        =   6
         Top             =   600
         Width           =   1035
      End
      Begin VB.Line Line2 
         X1              =   1395
         X2              =   4590
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Factory Code"
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
         Left            =   90
         TabIndex        =   4
         Top             =   240
         Width           =   1140
      End
      Begin VB.Line Line1 
         X1              =   1395
         X2              =   4590
         Y1              =   480
         Y2              =   480
      End
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   465
      Left            =   6960
      TabIndex        =   0
      Top             =   120
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   820
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   4365
      Left            =   90
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2400
      Width           =   8805
      _cx             =   15531
      _cy             =   7699
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
      SelectionMode   =   0
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
      ExplorerBar     =   1
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
   Begin VB.TextBox TxtSeqNo 
      Height          =   225
      Left            =   180
      TabIndex        =   15
      Top             =   7170
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Production Result Details"
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
      Height          =   390
      Left            =   180
      TabIndex        =   1
      Top             =   540
      Width           =   8370
   End
End
Attribute VB_Name = "FrmProdResultDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrSerialFrom As String
Dim StrSerialTo As String
Dim strItemCode As String
Dim StrLot As String

Dim dblSeqNo As Double

Dim bteColSelect As Byte
Dim bteColItem As Byte
Dim bteColItemName As Byte
Dim bteColLot As Byte
Dim bteColSerial As Byte
Dim bteColStatus As Byte
Dim bteColStatDesc As Byte

Private Sub Cmd_Back_Click(Index As Integer)
DoEvents
frmProdResult.Show
Unload Me
DoEvents
End Sub

Private Sub CmdScan_Click()
    frmProdScanBarcodeDet.Show vbModal
End Sub

Private Sub CmdSubmit_Click()
Dim IntRowCount As Integer
Dim StrBegSerial As String
Dim StrLastSerial As String
Dim strRemark As String
Dim dblQty As Double

For IntRowCount = 1 To grid.Rows - 1
    If grid.Cell(flexcpChecked, IntRowCount, bteColStatus) = flexChecked Then
        If StrBegSerial = "" Then
            StrBegSerial = Trim(grid.TextMatrix(IntRowCount, bteColSerial))
            StrLastSerial = Trim(grid.TextMatrix(IntRowCount, bteColSerial))
        Else
            StrLastSerial = Trim(grid.TextMatrix(IntRowCount, bteColSerial))
        End If
        dblQty = dblQty + 1
    End If
Next

For IntRowCount = 1 To grid.Rows - 1
    If grid.TextMatrix(IntRowCount, bteColSerial) >= StrBegSerial And grid.TextMatrix(IntRowCount, bteColSerial) <= StrLastSerial Then
        If grid.Cell(flexcpChecked, IntRowCount, bteColStatus) = flexUnchecked Then
            strRemark = strRemark & Trim(grid.TextMatrix(IntRowCount, bteColSerial)) & ","
        End If
    End If
Next

If Len(strRemark) > 0 Then
    strRemark = "Serial " & Left(strRemark, Len(strRemark) - 1) & " Skip "
End If

frmProdResult.TxtSerialFrom = StrBegSerial
frmProdResult.TxtSerialTo = StrLastSerial
frmProdResult.txtQty = dblQty
frmProdResult.txtremarks = strRemark
frmProdResult.Show
Unload Me

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then CmdScan_Click: KeyCode = 0
End Sub

Private Sub Form_Load()

StrSerialFrom = Trim(frmProdResult.TxtSerialFrom)
StrSerialTo = Trim(frmProdResult.TxtSerialTo)
StrLot = Trim(frmProdResult.txtLot)
strItemCode = Trim(frmProdResult.cbo(3))
StrLot = frmProdResult.txtLot

dblSeqNo = frmProdResult.dailyseqno

Call Header
Call BrowseGrid

End Sub

Private Sub Header()

'bteColselect = 0
bteColItem = 0
bteColItemName = 1
bteColLot = 2
bteColSerial = 3
bteColStatus = 4
bteColStatDesc = 5

With grid
    .ColS = 6
    .Rows = 1
    '.TextMatrix(0, bteColselect) = ""
    .TextMatrix(0, bteColItem) = "Product Code"
    .TextMatrix(0, bteColItemName) = "Product Name"
    .TextMatrix(0, bteColLot) = "LOT Number"
    .TextMatrix(0, bteColSerial) = "Serial Number"
    .TextMatrix(0, bteColStatus) = "Result"
    .TextMatrix(0, bteColStatDesc) = "Serial Status"

     '.ColWidth(bteColselect) = 300
     .ColWidth(bteColItem) = 1500
     .ColWidth(bteColItemName) = 2500
     .ColWidth(bteColLot) = 1200
     .ColWidth(bteColSerial) = 1500
     .ColWidth(bteColStatus) = 600
     .ColWidth(bteColStatDesc) = 1200

     .ColAlignment(bteColItem) = flexAlignLeftCenter
     .ColAlignment(bteColItemName) = flexAlignLeftCenter
     .ColAlignment(bteColLot) = flexAlignCenterCenter
     .ColAlignment(bteColSerial) = flexAlignCenterCenter
     .ColAlignment(bteColStatus) = flexAlignCenterCenter
     .ColAlignment(bteColStatDesc) = flexAlignLeftCenter
     
    .Cell(flexcpAlignment, 0, bteColStatus, 0, ColFix) = flexAlignCenterCenter
     
          
End With
End Sub
Private Sub BrowseGrid()
Dim i As Long
Dim strSQL As String
Dim RsDet As New ADODB.Recordset

If StrSerialFrom <> "" And StrSerialTo <> "" Then
    
'Menampilkan semua item pada daily Schedule yang sama
    
    strSQL = " select * From " & vbCrLf & _
                      " (Select a.item_Code,Item_Name,Serial_No,Serial_Status From Serial_Detail a  " & vbCrLf & _
                      "     inner join item_master b on a.item_Code=b.item_Code  " & vbCrLf & _
                      "     Where a.Item_code='" & strItemCode & "' And Serial_No>='" & StrSerialFrom & "' and  " & vbCrLf & _
                      "     Serial_No<='" & StrSerialTo & "' " & _
                      "     and Product_No=" & dblSeqNo
        
        If Val(TxtSeqNo) > 0 Then
                      strSQL = strSQL + "     and Result_No=" & Val(TxtSeqNo)
        End If
        
        strSQL = strSQL + " union  "
        
End If
                      
strSQL = strSQL + " " & _
                  " Select a.item_Code,Item_Name,Serial_No,Serial_Status From Serial_Detail a  " & vbCrLf & _
                  "     inner join item_master b on a.item_Code=b.item_Code " & vbCrLf & _
                  "     inner join daily_production c on  a.Product_No=c.seq_No  " & vbCrLf & _
                  "     where result_No is null and seq_no=" & dblSeqNo

If StrSerialFrom <> "" And StrSerialTo <> "" Then strSQL = strSQL + " ) d "
                  
strSQL = strSQL + "  order by serial_No "

    Set RsDet = Db.Execute(strSQL)
    
    i = 1
    With grid
        Do While Not RsDet.EOF
            .Rows = .Rows + 1
            .TextMatrix(i, bteColItem) = RsDet("Item_Code")
            .TextMatrix(i, bteColItemName) = RsDet("Item_Name")
            .TextMatrix(i, bteColLot) = StrLot
            .TextMatrix(i, bteColSerial) = RsDet("Serial_No")
            If RsDet.Fields("Serial_Status") = "3" Then
               .TextMatrix(i, bteColStatDesc) = "Result"
               .Cell(flexcpChecked, i, bteColStatus) = flexChecked
            Else
               .TextMatrix(i, bteColStatDesc) = "Process"
               .Cell(flexcpChecked, i, bteColStatus) = flexUnchecked
            End If
            RsDet.MoveNext
            i = i + 1
        Loop
    End With
'End If
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If grid.Col <> bteColStatus Then Cancel = True: Exit Sub
End Sub

Private Sub Grid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
  If grid.Col = bteColSelect Then
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii <> Asc("S") And KeyAscii <> Asc("D") And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then
      KeyAscii = 0
    End If
    If KeyAscii = Asc(".") Then KeyAscii = 0
  End If
End Sub

Sub CheckSerial(Item As String, Serial As String)
Dim InRowCount As Integer
Dim IntFindStat As Integer
IntFindStat = 0

frmProdScanBarcodeDet.LblErrMsg.Caption = ""

For InRowCount = 1 To grid.Rows - 1
    If Trim(Item) = Trim(grid.TextMatrix(InRowCount, bteColItem)) And Serial = Trim(grid.TextMatrix(InRowCount, bteColSerial)) Then
        grid.TextMatrix(InRowCount, bteColStatDesc) = "Result"
        grid.Cell(flexcpChecked, InRowCount, bteColStatus) = flexChecked
        IntFindStat = 1
        Exit For
    End If
Next

If IntFindStat = 0 Then
    frmProdScanBarcodeDet.LblErrMsg.Caption = " Invalid Item of Production Range ! "
End If
frmProdScanBarcodeDet.txtBarcode = ""

End Sub
