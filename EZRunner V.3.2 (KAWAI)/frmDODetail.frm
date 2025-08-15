VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmDODetail 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Delivery Order Detail"
   ClientHeight    =   9585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10185
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDODetail.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9585
   ScaleWidth      =   10185
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdScan 
      BackColor       =   &H0080FFFF&
      Caption         =   "Scan Bar&code (F2)"
      Height          =   375
      Left            =   4470
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   8730
      Width           =   2565
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Submit"
      Height          =   375
      Index           =   0
      Left            =   8467
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8730
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   577
      TabIndex        =   8
      Top             =   7980
      Width           =   9030
      Begin VB.Label LblErrMsg 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   90
         TabIndex        =   9
         Top             =   210
         Width           =   8835
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   7140
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8730
      Width           =   1140
   End
   Begin VB.CommandButton cmdSubMenu 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Back"
      Height          =   375
      Left            =   577
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8730
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   1575
      Left            =   577
      TabIndex        =   3
      Top             =   870
      Width           =   9030
      Begin VB.Label LblPoSeqNo 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   6000
         TabIndex        =   21
         Top             =   1260
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   195
         Index           =   4
         Left            =   2220
         TabIndex        =   20
         Top             =   1125
         Width           =   75
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   195
         Index           =   1
         Left            =   2220
         TabIndex        =   19
         Top             =   705
         Width           =   75
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   195
         Index           =   0
         Left            =   2220
         TabIndex        =   18
         Top             =   300
         Width           =   75
      End
      Begin VB.Label LblDoSeqNo 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   6000
         TabIndex        =   17
         Top             =   1020
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label LblSerialTo 
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXXXXXXXX"
         Height          =   255
         Left            =   6000
         TabIndex        =   16
         Top             =   660
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label LblSerialFrom 
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXXXXXXXX"
         Height          =   255
         Left            =   6030
         TabIndex        =   15
         Top             =   330
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label LblProdCode 
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXXXXXXXX"
         Height          =   255
         Left            =   2520
         TabIndex        =   13
         Top             =   1095
         Width           =   2535
      End
      Begin VB.Label LblPO 
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXXXXXXXX"
         Height          =   255
         Left            =   2520
         TabIndex        =   12
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SI/ PO Number"
         Height          =   195
         Index           =   3
         Left            =   660
         TabIndex        =   10
         Top             =   705
         Width           =   1305
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Code"
         Height          =   195
         Left            =   840
         TabIndex        =   7
         Top             =   1125
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DN number"
         Height          =   195
         Index           =   2
         Left            =   960
         TabIndex        =   5
         Top             =   300
         Width           =   975
      End
      Begin VB.Label LblDNNo 
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXXXXXXXX"
         Height          =   255
         Left            =   2520
         TabIndex        =   4
         Top             =   300
         Width           =   2535
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid grid 
      Height          =   5040
      Left            =   600
      TabIndex        =   11
      Top             =   2640
      Width           =   9030
      _cx             =   15928
      _cy             =   8890
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
      HighLight       =   2
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
      RowHeightMax    =   0
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
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   420
      Left            =   7680
      TabIndex        =   14
      Top             =   120
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   741
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery Order Detail"
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
      Left            =   3870
      TabIndex        =   6
      Top             =   270
      Width           =   2430
   End
End
Attribute VB_Name = "frmDODetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ClsProc As New ClsProc
Dim i As Long, StCek As Byte
Dim TempFrom As String, TempTo As String, tempSeq As String, TempPOSeq As String

Const col_Serial As Integer = 0
Const col_Item As Integer = 1
Const col_Doc As Integer = 2
Const col_Temp1 As Integer = 3
Const col_Temp2 As Integer = 4
Const col_Itemcode As Integer = 5
Const col_Check As Integer = 6

Const col_Count As Integer = 7

Private Sub CmdScan_Click()
    frmDOScanBarcode.Show vbModal
End Sub

Private Sub Form_Load()
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"
    IsiGrid
End Sub

Private Sub headerGrid()
With grid
    .clear
    .ColS = col_Count
    .Rows = 1
    
    .ColWidth(0) = 2000
    .ColWidth(1) = 3500
    .ColWidth(2) = 1000
'    .ColHidden(3) = True
'    .ColHidden(4) = True
'    .ColHidden(5) = True
    .ColWidth(6) = 1000

    .TextMatrix(0, 0) = "Serial No"
    .TextMatrix(0, 1) = "Product Name"
    .Cell(flexcpChecked, 0, 2) = flexUnchecked
    .TextMatrix(0, 2) = "Status"
    .TextMatrix(0, 6) = "Check"
    
    .Cell(flexcpAlignment, 0, col_Serial, 0, col_Temp2) = flexAlignCenterCenter
End With
End Sub

Sub IsiGrid()
Dim rsGrid As New ADODB.Recordset
Dim sqlResult As String, ls_PONo As String, li_Row As String

TempFrom = Trim(frmDOCreate.l1)
TempTo = Trim(frmDOCreate.l2)
tempSeq = Trim(frmDOCreate.l3)
TempPOSeq = Trim(frmDOCreate.l4)
LblSerialFrom = Trim(frmDOCreate.l5)
LblSerialTo = Trim(frmDOCreate.l6)

ls_PONo = frmDOCreate.gridBawah.TextMatrix(frmDOCreate.gridBawah.RowSel, 5)
LblDNNo = frmDOCreate.txtDoNO
LblPO = ls_PONo
lblProdCode = frmDOCreate.gridBawah.TextMatrix(frmDOCreate.gridBawah.RowSel, 1)

With grid
    Call headerGrid

    sql = "SELECT Serial_Detail.*, Item_Master.Item_Name FROM Serial_Detail " & _
            vbLf & " Inner Join Item_Master on Serial_Detail.Item_Code=Item_Master.Item_Code" & _
            vbLf & " WHERE --PO_No='" & ls_PONo & "' " & _
            vbLf & " Serial_No >='" & Trim(TempFrom) & "' And Serial_No<='" & Trim(TempTo) & "' Order By Serial_No"

'            vbLf & " AND Serial_Detail.Item_Code='" & LblProdCode & "' AND Serial_Status in ('3','4') " & _
'            vbLf & " And (DO_No='" & Trim(LblDNNo) & "' or DO_No is Null)" & _
'            vbLf & " And Serial_No >='" & Trim(TempFrom) & "' And Serial_No<='" & Trim(TempTo) & "' Order By Serial_No"
        
       ' If Trim(LblDoSeqNo) <> "" Then Sql = Sql & " And Do_SeqNo=" & Trim(TempSeq) & " Order By Serial_No"
            
    Set rsGrid = Db.Execute(sql)
    
    i = 1
    If Not (rsGrid.EOF) Then
        Do While Not rsGrid.EOF
            .Rows = .Rows + 1
            
            .TextMatrix(i, col_Serial) = Trim(rsGrid("Serial_No"))
            .TextMatrix(i, col_Item) = Trim(rsGrid("Item_Name"))
            .TextMatrix(i, col_Itemcode) = Trim(rsGrid("Item_Code"))
            i = i + 1
            rsGrid.MoveNext
        Loop
        
        ' Check Serial List
        Dim lnPRow As Integer, lnDRow As Integer
        lnDRow = 1
        
        Do
            If lnDRow > .Rows - 1 Then Exit Do
            
            lnPRow = 1
            Do
                If lnPRow > frmDOCreate.gridBawah.Rows - 1 Then Exit Do
                    LblSerialFrom = Trim(frmDOCreate.gridBawah.TextMatrix(lnPRow, 9) & "")
                    LblSerialTo = Trim(frmDOCreate.gridBawah.TextMatrix(lnPRow, 10) & "")
                    
                    If Trim(frmDOCreate.gridBawah.TextMatrix(lnPRow, 1) & "") = Trim(.TextMatrix(lnDRow, col_Itemcode)) Then
                        If Trim(.TextMatrix(lnDRow, col_Serial) & "") >= LblSerialFrom And Trim(.TextMatrix(lnDRow, col_Serial) & "") <= LblSerialTo Then
                            .Cell(flexcpChecked, lnDRow, col_Doc) = flexChecked
                            .TextMatrix(lnDRow, col_Temp1) = 4
                            .TextMatrix(lnDRow, col_Temp2) = 4
                            Exit Do
                        Else
                            .Cell(flexcpChecked, lnDRow, col_Doc) = flexUnchecked
                            .TextMatrix(lnDRow, col_Temp1) = 3
                            .TextMatrix(lnDRow, col_Temp2) = 3

                        End If
                    End If
                lnPRow = lnPRow + 1
            Loop
            lnDRow = lnDRow + 1
        Loop
        
        
    Else
        LblErrMsg = DisplayMsg(4006)
    End If
    Set rsGrid = Nothing
End With
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'Dim rsCek As New Recordset
'
'    LblErrMsg = ""
'    If Col <= 3 Then Cancel = 1
'    If Col = 4 And Row > 0 Then
'        StCek = 0
'        Sql = "select * from PurchaseOrder_Detail where PORequest_No = '" & Trim(grid.TextMatrix(Row, 0)) & "' "
'        If rsCek.State <> adStateClosed Then rsCek.Close
'        rsCek.Open Sql, Db, adOpenKeyset, adLockOptimistic
'        If Not (rsCek.BOF And rsCek.EOF) Then
'            If grid.Cell(flexcpChecked, Row, 4) = flexChecked Then
'                StCek = 1: Cancel = True
'            End If
'        End If
'        Set rsCek = Nothing
'    End If
End Sub

Private Sub grid_Click()
Dim BeforeCheck As String

With grid
    'If .Row > 0 Then
        If .Col = col_Doc Then
            If .CellChecked = flexChecked Then
                .CellChecked = flexChecked
                .TextMatrix(.Row, col_Temp2) = 4
            Else
                .CellChecked = flexUnchecked
                .TextMatrix(.Row, col_Temp2) = 3
            End If
        End If
    'End If
End With
End Sub

Private Sub Command1_Click(Index As Integer)
Dim tanya

Me.MousePointer = vbHourglass
Select Case Index
    Case 0: 'Submit
'        tanya = MsgBox("Do you really want to Process Delivery Note ?", vbQuestion & vbYesNo, "Confirmation")
'        If tanya = vbYes Then
        Call simpan
        
    Case 1: 'Cancel
        Call IsiGrid
End Select
Me.MousePointer = vbDefault
End Sub

Sub simpan()
Dim serialNo As String, li_Temp1 As Byte, li_Temp2 As Byte
Dim ItemCode As String, DN_No As String
Dim li_Row As Double, li_count As Double
Dim ls_From As String, ls_To As String
Dim SqlS As String
Dim TempAwal As String
Dim TempAkhir As String
Dim TempCount As Integer
Dim RCek As Integer, TempRCek  As Integer, X As Integer, n As Integer
Dim TempFound As Boolean, TempDoSeqNo As Integer

With grid

    With frmDOCreate.gridBawah
        For RCek = 1 To .Rows - 1
            .TextMatrix(RCek, 31) = ""
        Next
    End With
    
    ItemCode = lblProdCode
    TempDoSeqNo = 1
    For i = 1 To .Rows - 1
        
        serialNo = Trim(.TextMatrix(i, col_Serial))
        li_Temp1 = Trim(.TextMatrix(i, col_Temp1)) ' Serial Status Awal
        li_Temp2 = Trim(.TextMatrix(i, col_Temp2)) ' Serial Status Baru
        
        If li_Temp2 = 3 Then
            SqlS = "Update Serial_Detail Set DO_No=Null," & _
                    " Do_SeqNo=Null," & _
                    " Serial_Status=3 " & _
                    " Where Item_Code='" & ItemCode & "' " & _
                    " And Serial_No='" & serialNo & "' " & _
                    " And Po_No='" & LblPO & "' "
        Else
            SqlS = "Update Serial_Detail Set DO_No='" & LblDNNo & "'," & _
                    " Do_SeqNo='" & TempDoSeqNo & "'," & _
                    " Serial_Status=4 " & _
                    " Where Item_Code='" & ItemCode & "' " & _
                    " And Serial_No='" & serialNo & "' " & _
                    " And Po_No='" & LblPO & "' "
        End If
        Db.Execute (SqlS)
        
        If li_Temp2 = 4 And i <> .Rows - 1 Then
            If TempAwal = "" Then
                TempAwal = serialNo
            End If
            TempCount = TempCount + 1
            TempAkhir = Trim(.TextMatrix(i - 1, col_Serial))
        Else
            If TempCount > 0 Then
                TempAkhir = Trim(.TextMatrix(i - 1, col_Serial))
                MsgBox TempAwal & " - " & TempAkhir & "(" & TempCount & ")"
                With frmDOCreate.gridBawah
                    RCek = 1
                    TempFound = False
                    Do
                        If RCek > .Rows - 1 Then
                            If Not TempFound Then
                                X = 1
                                Do
                                    If X > frmDOCreate.gridAtas.Rows - 1 Then Exit Do
                                        If frmDOCreate.gridAtas.TextMatrix(X, 5) = LblPO And frmDOCreate.gridAtas.TextMatrix(X, 24) = TempPOSeq Then
                                            .Rows = .Rows + 1
                                            For n = 1 To 30
                                                .TextMatrix(.Rows - 1, n) = frmDOCreate.gridAtas.TextMatrix(X, n)
                                            Next
                                            .Cell(flexcpBackColor, .Rows - 1, 0) = vbWhite
                                            .Cell(flexcpBackColor, .Rows - 1, 4) = vbWhite
                                            .Cell(flexcpBackColor, .Rows - 1, 6) = vbWhite
                                            .Cell(flexcpBackColor, .Rows - 1, 15) = vbWhite
                                            .Cell(flexcpBackColor, .Rows - 1, 16) = vbWhite
                                            .TextMatrix(.Rows - 1, 31) = "Checked"
                                            .TextMatrix(.Rows - 1, 6) = TempCount
                                            .TextMatrix(.Rows - 1, 9) = TempAwal
                                            .TextMatrix(.Rows - 1, 10) = TempAkhir
                                            .TextMatrix(.Rows - 1, 25) = TempDoSeqNo
                                            TempDoSeqNo = TempDoSeqNo + 1
                                            Call frmDOCreate.gridBawah_AfterEdit(.Rows - 1, 6)
                                            Exit Do
                                        End If
                                    X = X + 1
                                Loop
                            End If
                            Exit Do
                        End If
                        'cek Record
                         If ItemCode = .TextMatrix(RCek, 1) And .TextMatrix(RCek, 5) = LblPO _
                            And .TextMatrix(RCek, 24) = TempPOSeq Then
                                If .TextMatrix(RCek, 31) = "" Then
                                    .TextMatrix(RCek, 31) = "Checked"
                                    .TextMatrix(RCek, 6) = TempCount
                                    .TextMatrix(RCek, 9) = TempAwal
                                    .TextMatrix(RCek, 10) = TempAkhir
                                    .TextMatrix(.Rows - 1, 25) = TempDoSeqNo
                                    TempDoSeqNo = TempDoSeqNo + 1
                                    Call frmDOCreate.gridBawah_AfterEdit(RCek, 6)
                                    TempFound = True
                                    Exit Do
                                Else
                                    TempRCek = RCek
                                End If
                         End If
                         RCek = RCek + 1
                    Loop
                    .Col = 24
                    .Sort = flexSortStringAscending
                    .Col = 6
                    .Sort = flexSortStringAscending
                    .Col = 9
                    .Sort = flexSortStringAscending
                    .Col = 0
                End With
                TempAwal = ""
                TempAkhir = ""
                TempCount = 0
            End If
        End If
    Next
           
End With

With frmDOCreate.gridBawah
    RCek = 1
    Do
        If RCek > .Rows - 1 Then Exit Do
        If ItemCode = .TextMatrix(RCek, 1) And .TextMatrix(RCek, 5) = LblPO _
              And .TextMatrix(RCek, 24) = TempPOSeq And .TextMatrix(RCek, 31) = "" Then
            .TextMatrix(RCek, 6) = 0
            .TextMatrix(RCek, 9) = ""
            .TextMatrix(RCek, 10) = ""
            .TextMatrix(RCek, 23) = "u"
            .RowHidden(RCek) = True
        End If
         RCek = RCek + 1
    Loop
End With

Call CmdSubMenu_Click
        
End Sub

Private Sub CmdSubMenu_Click()
    DoEvents
    frmDOCreate.Show
    frmDOCreate.Enabled = True
    DoEvents
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
    If ErrMsg = "" Then Unload Me Else LblErrMsg.Caption = ErrMsg
End Sub


Sub CheckSerial(Item As String, Serial As String)
Dim InRowCount As Integer
Dim IntFindStat As Integer
IntFindStat = 0

frmDOScanBarcode.LblErrMsg.Caption = ""

For InRowCount = 1 To grid.Rows - 1
    If Trim(Item) = Trim(grid.TextMatrix(InRowCount, col_Itemcode)) And Serial = Trim(grid.TextMatrix(InRowCount, bteColSerial)) Then
        grid.TextMatrix(InRowCount, col_Temp2) = "4"
        grid.Cell(flexcpChecked, InRowCount, col_Check) = flexChecked
        IntFindStat = 1
        Exit For
    End If
Next

If IntFindStat = 0 Then
    frmDOScanBarcode.LblErrMsg.Caption = " Invalid Item of Production Range ! "
End If
frmDOScanBarcode.txtBarcode = ""

End Sub


