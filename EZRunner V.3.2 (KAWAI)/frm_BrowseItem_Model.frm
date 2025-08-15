VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm_BrowseItem_Model 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Browse Model"
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6180
   Icon            =   "frm_BrowseItem_Model.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7260
   ScaleWidth      =   6180
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTmpModel 
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
      Left            =   0
      MaxLength       =   30
      TabIndex        =   10
      Top             =   6840
      Visible         =   0   'False
      Width           =   2940
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   795
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   5895
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
         Left            =   2880
         MaxLength       =   30
         TabIndex        =   6
         Top             =   240
         Width           =   2940
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search By"
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
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   885
      End
      Begin MSForms.ComboBox cbxChoose 
         Height          =   315
         Left            =   1200
         TabIndex        =   7
         Top             =   240
         Width           =   1620
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "2857;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   120
      TabIndex        =   3
      Top             =   6000
      Width           =   5895
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
         Left            =   120
         TabIndex        =   4
         Top             =   195
         Width           =   5655
      End
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Cancel"
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
      Left            =   4860
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6720
      Width           =   1155
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H0080FFFF&
      Caption         =   "&OK"
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
      Left            =   3630
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6720
      Width           =   1140
   End
   Begin VSFlex8Ctl.VSFlexGrid gridSearch 
      Height          =   4575
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   5925
      _cx             =   10451
      _cy             =   8070
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
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   275
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frm_BrowseItem_Model.frx":0E42
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
      Editable        =   2
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
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Browse Model Item"
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
      Left            =   1890
      TabIndex        =   0
      Top             =   120
      Width           =   2235
   End
End
Attribute VB_Name = "frm_BrowseItem_Model"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Seq As String
Public QtySerial As Double
Public ResultSeqNo1 As String
Public strItemCode As String
Public status As String
Public WIPseqNo As Double
Dim TempCekA As Integer

Public Sub IsiGrid()
    Dim RS As New ADODB.Recordset
    Dim query As String
    
    'normalisasi Delete serial no yang data hasil produksinya tidak ada
'    query = "delete from part_receipt where Seq_No=0"
'    Db.Execute query

    'normalisasi Delete serial no yang data hasil produksinya tidak ada
'    query = "Delete from prodresult_serialno where not exists (select seq_no from part_receipt where receipt_cls='P1'  and seq_no = prodresult_serialno.resultseq_no )"
'    Db.Execute query
    
'    query = "delete from dailyprod_serialno " & vbCrLf & _
'            "    where dpseq_no='" & Seq & "' and serial_no " & vbCrLf & _
'            " not in (select serial_no from prodresult_serialno where dpseq_no='" & Seq & "') and cancel_cls='0' and item_code <> '" & strItemCode & "'"
'    Db.Execute query
    
'    If status = "edit" Then
'        query = "select * from dailyprod_serialno where dpseq_no='" & Seq & "' and serial_no  in (select serial_no from prodresult_serialno where resultseq_no='" & ResultSeqNo1 & "') and cancel_cls='0' " & vbCrLf & _
'                "union " & vbCrLf & _
'                "select * from dailyprod_serialno where dpseq_no='" & Seq & "'  " & vbCrLf & _
'                "    and serial_no   not in (select serial_no from prodresult_serialno) and cancel_cls='0' "
'    Else
        query = "SELECT Model_Cls, Description FROM Model_Cls"
'    End If
    
    If RS.State = adStateOpen Then RS.Close
    RS.Open query, Db, adOpenForwardOnly, adLockReadOnly
    
    gridSearch.clear
    Header

    i = 0
    While Not RS.EOF
        With gridSearch
        i = i + 1
        .AddItem ""
        
        .Cell(flexcpBackColor, i, 0) = vbWhite
        
'        If Trim(Get_Record("select isnull(serial_no,'') serial_no from prodresult_serialno where DPSeq_no='" & Trim(rs!dpseq_no) & "' and ResultSeq_no='" & ResultSeqNo1 & "' and Serial_no='" & Trim(rs!serial_no) & "' ")) <> "" Then
'        .Cell(flexcpChecked, i, 0) = flexChecked
'        Else
        .Cell(flexcpChecked, i, 0) = flexUnchecked
'        End If
        
        .TextMatrix(i, 1) = Trim(RS!Model_Cls)
        .TextMatrix(i, 2) = Trim(RS!Description)
        RS.MoveNext
        
        End With
    Wend
    If gridSearch.Rows > 1 Then
        gridSearch.Row = 1
        
    End If
End Sub

Public Sub clear()
    txtBrowse.Text = ""
    txtBrowse.TabIndex = 0
    gridSearch.TabIndex = 1
    cmdOK.TabIndex = 2
    cmdCancel.TabIndex = 3
    LblErrMsg.Alignment = 2
    txtTmpModel.Text = ""
    TempCekA = 0
End Sub

Public Sub isiCbx()
    cbxChoose.AddItem
    cbxChoose.List(0, 0) = "Model Cls"
    cbxChoose.AddItem
    cbxChoose.List(1, 0) = "Description"
    cbxChoose.ColumnWidths = "30 pt "
    cbxChoose.ListIndex = 0
End Sub

Public Sub Header()
    With gridSearch
        .clear
        .Rows = 1
        .ColS = 3
'        .Cell(flexcpChecked, 0) = flexUnchecked
        .TextMatrix(0, 1) = "Model Cls"
        .TextMatrix(0, 2) = "Description"
        .ColWidth(0) = 300
        .ColWidth(1) = 950
        .ColWidth(2) = 2000
        
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
    End With
End Sub

'Private Sub CmdOK_Click()
'Dim strSQL As String
'Dim RsCheck As New ADODB.Recordset
'Dim jml As Double
'Dim sterr As Boolean
'jml = 0
'sterr = False
'
'For i = 1 To gridSearch.Rows - 1
'    If gridSearch.Cell(flexcpChecked, i, 0) = flexChecked Then
'        jml = jml + 1
'    End If
'
'    If WIPseqNo > 0 Then 'Punya Child
'        'Check apakah proses sebelumnya sudah diinput?
'        If gridSearch.Cell(flexcpChecked, i, 0) = flexChecked Then
'            strSQL = "select B.*from Daily_Production A left join Manufacture_Line B on A.Line_Code = B.Line_Code " & vbCrLf & _
'                        " where seq_no = " & Seq & " and B.NeedCheck_Cls ='1' "
'            Set RsCheck = New Recordset
'            If RsCheck.State <> adStateClosed Then RsCheck.Close
'            RsCheck.Open strSQL, Db, adOpenStatic, adLockReadOnly
'            If RsCheck.EOF = False Then
'                strSQL = "select * from ProdResult_SerialNo " & vbCrLf & _
'                        " where ResultSeq_No  in (select Seq_No from Part_Receipt " & vbCrLf & _
'                        "                           where DailySeq_No = (select seq_no from Daily_Production " & vbCrLf & _
'                        "                                    where WIPProdSeq_No = " & WIPseqNo & ") " & vbCrLf & _
'                        "                        ) " & vbCrLf & _
'                        " and serial_no ='" & gridSearch.TextMatrix(i, 1) & "'"
'                Set RsCheck = New Recordset
'                If RsCheck.State <> adStateClosed Then RsCheck.Close
'                RsCheck.Open strSQL, Db, adOpenStatic, adLockReadOnly
'                If RsCheck.EOF = True Then
'                    MsgBox "Please input Production Result for Previous Line First", vbCritical + vbOKOnly, "Error"
'                    gridSearch.Cell(flexcpBackColor, i, 0, i, 1) = vbRed
'                    sterr = True
'                    'Exit For
'                End If
'                RsCheck.Close
'                Set RsCheck = Nothing
'            End If
'        End If
'    End If
'Next i
'
'If sterr Then
'    frmProdResult.txtQty.Text = 0
'Else
'    frmProdResult.txtQty.Text = Format(jml, gs_formatQty)
'    Me.Hide
'End If
'
'End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
txtTmpModel.Text = ""

With gridSearch
    For i = 1 To gridSearch.Rows - 1
        If gridSearch.Cell(flexcpChecked, i, 0) = flexChecked Then
            If txtTmpModel.Text <> "" Then
                txtTmpModel.Text = txtTmpModel.Text + "," + .TextMatrix(i, 1)
            Else
                txtTmpModel.Text = .TextMatrix(i, 1)
            End If
        End If
    Next i
End With

Me.Hide

End Sub

Private Sub Form_Load()
    Me.MousePointer = vbDefault
    IsiGrid
    isiCbx
    clear
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub gridSearch_AfterEdit(ByVal Row As Long, ByVal Col As Long)

Dim cek As Integer
LblErrMsg.Caption = ""

With gridSearch
    If Row <> 0 Then
        If Col = 0 Then Call cekSelect(Col)
    Else
        If Col = 0 Then
            If .Cell(flexcpChecked, Row, Col) = 1 Then
                cek = 1 'flexChecked
            Else
                cek = 2 'flexUnchecked
            End If
                
            For i = 1 To .Rows - 1
                .Cell(flexcpChecked, i, Col) = cek
            Next i
        End If
    End If
    
    
'    If TempModel <> "" Then
'        TempModel = TempModel + "," + .TextMatrix(Row, 1)
'    Else
'        TempModel = .TextMatrix(Row, 1)
'    End If
    
    If TempCekA > 5 Then
        .Cell(flexcpChecked, Row, Col) = flexUnchecked
        LblErrMsg.Caption = "Cannot select more data"
    End If
    
'    If TempCekB > TempCekA Then
'    TempModelCls = .TextMatrix(Row, 1) + ","
'        TempModel = Replace(TempModel, .TextMatrix(Row, 1), "")
'        TempModel = Replace(TempModel, ",,", ",")
'    End If
''
''
'    TempCekB = TempCekA
End With

End Sub

Private Sub gridSearch_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim strSQL As String
Dim rsCheck As Recordset
If Col <> 0 Then Cancel = True
'LblErrMsg.Caption = ""
End Sub

Sub cekSelect(kol As Long)
Dim cek, noCek As Integer

With gridSearch
    '******** agar cekBox nya jika semuanya udah ke-select/not
    cek = 0
    For i = 1 To .Rows - 1
        If .Cell(flexcpChecked, i, kol) = flexChecked Then
            cek = cek + 1
        Else
            noCek = noCek + 1
        End If
    Next i
    
    TempCekA = cek
    
'    If cek = .Rows - 1 Then
'        .Cell(flexcpChecked, 0, kol) = flexChecked
'    ElseIf noCek >= 1 Then
'        '.Cell(flexcpChecked, 0, kol) = flexUnchecked
'    End If
    
    '****************
End With
End Sub

Private Sub txtbrowse_KeyPress(KeyAscii As Integer)
Dim temu As Integer

KeyAscii = Asc(UCase(Chr(KeyAscii)))
 
If KeyAscii = 13 Then
    
    LblErrMsg.Caption = ""
    
    If gridSearch.Rows > 1 Then
    
        gridSearch.Row = 1
        gridSearch.TopRow = 1
    
        If cbxChoose.Text = "Model Cls" Then
            temu = gridSearch.FindRow(Trim(txtBrowse.Text), , 1)
        ElseIf cbxChoose.Text = "Description" Then
            temu = gridSearch.FindRow(Trim(txtBrowse.Text), , 2)
        End If
        
        
        If temu > 0 Then
            gridSearch.Row = temu
            gridSearch.TopRow = temu
            gridSearch.Cell(flexcpChecked, temu, 0) = flexChecked
            
            txtBrowse.SelStart = 0
            txtBrowse.SelLength = Len(txtBrowse)
            txtBrowse.SetFocus
            LblErrMsg.Caption = ""
        Else
            LblErrMsg.Caption = cbxChoose.Text + " Not Found !"
            txtBrowse.SelStart = 0
            txtBrowse.SelLength = Len(txtBrowse)
            txtBrowse.SetFocus
        End If
        
    End If

End If
End Sub


