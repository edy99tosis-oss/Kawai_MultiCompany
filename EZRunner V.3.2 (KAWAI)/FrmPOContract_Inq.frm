VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmPOContract_Inq 
   BackColor       =   &H00FDDFE3&
   Caption         =   "PO Contract Inwuiry"
   ClientHeight    =   10620
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15465
   Icon            =   "FrmPOContract_Inq.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10620
   ScaleWidth      =   15465
   WindowState     =   2  'Maximized
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
      Left            =   14160
      Style           =   1  'Graphical
      TabIndex        =   17
      Tag             =   "FFTT*/"
      Top             =   10080
      Width           =   1140
   End
   Begin VB.CommandButton CmdSubMenu 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Sub &Menu"
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
      TabIndex        =   16
      Tag             =   "TFFT*/"
      Top             =   10080
      Width           =   1140
   End
   Begin VB.CommandButton CmdClear 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Clear"
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
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   15
      Tag             =   "FFTT*/"
      Top             =   10080
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   120
      TabIndex        =   13
      Tag             =   "TFTT*/"
      Top             =   9360
      Width           =   15165
      Begin VB.Label lblErrMsg 
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
         TabIndex        =   14
         Tag             =   "TTTF*/"
         Top             =   195
         Width           =   14805
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FDDFE3&
      Height          =   1635
      Left            =   120
      TabIndex        =   2
      Tag             =   "TTTF*/"
      Top             =   1200
      Width           =   15165
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H0080FFFF&
         Caption         =   "&Search"
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
         Left            =   8280
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "TTFF*/"
         Top             =   1080
         Width           =   1140
      End
      Begin VB.TextBox lblCust 
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
         Height          =   240
         Left            =   3360
         TabIndex        =   3
         TabStop         =   0   'False
         Tag             =   "TTFF*/"
         Top             =   720
         Width           =   4815
      End
      Begin MSComCtl2.DTPicker DTFrom 
         Height          =   315
         Left            =   1680
         TabIndex        =   5
         Tag             =   "TTFF*/"
         Top             =   240
         Width           =   1560
         _ExtentX        =   2752
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
         Format          =   129826819
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker DTTo 
         Height          =   315
         Left            =   3720
         TabIndex        =   6
         Tag             =   "TTFF*/"
         Top             =   240
         Width           =   1560
         _ExtentX        =   2752
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
         Format          =   129826819
         CurrentDate     =   37798
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
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
         Index           =   7
         Left            =   240
         TabIndex        =   12
         Tag             =   "TTFF*/"
         Top             =   1155
         Width           =   540
      End
      Begin MSForms.ComboBox cboStatus 
         Height          =   315
         Left            =   1680
         TabIndex        =   11
         Tag             =   "TTFF*/"
         Top             =   1080
         Width           =   1530
         VariousPropertyBits=   612386843
         MaxLength       =   50
         DisplayStyle    =   3
         Size            =   "2699;556"
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Code"
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
         TabIndex        =   10
         Tag             =   "TTFF*/"
         Top             =   720
         Width           =   1350
      End
      Begin MSForms.ComboBox cboCust 
         Height          =   315
         Left            =   1680
         TabIndex        =   9
         Tag             =   "TTFF*/"
         Top             =   645
         Width           =   1530
         VariousPropertyBits=   612386843
         MaxLength       =   10
         DisplayStyle    =   3
         Size            =   "2699;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Line Line2 
         X1              =   3360
         X2              =   8160
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Period"
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
         TabIndex        =   8
         Tag             =   "TTFF*/"
         Top             =   285
         Width           =   540
      End
      Begin VB.Label lblCaption 
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
         Index           =   2
         Left            =   3360
         TabIndex        =   7
         Tag             =   "TTFF*/"
         Top             =   285
         Width           =   165
      End
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   420
      Left            =   13440
      TabIndex        =   0
      Tag             =   "FTTF*/"
      Top             =   120
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   741
   End
   Begin EZRunnerv3.Anchor Anchor1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   820
   End
   Begin VSFlex8Ctl.VSFlexGrid grid 
      Height          =   6390
      Left            =   120
      TabIndex        =   18
      TabStop         =   0   'False
      Tag             =   "TTTT*/"
      Top             =   2880
      Width           =   15195
      _cx             =   26802
      _cy             =   11271
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
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   3
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   275
      RowHeightMax    =   275
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
      ExplorerBar     =   1
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
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Order Contract Inquiry"
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
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Tag             =   "TTTF*/"
      Top             =   480
      Width           =   15165
   End
End
Attribute VB_Name = "FrmPOContract_Inq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim bteColContractNo As Byte
Dim bteColContractDate As Byte
Dim bteColItemCode As Byte
Dim bteColItemName As Byte
Dim bteColPONo As Byte
Dim bteColSupplierCode As Byte
Dim bteColSupplierName As Byte
Dim bteColQtyPO As Byte
Dim bteColQtyReceipt As Byte
Dim bteColStatus As Byte
Dim bteColTot As Byte
Dim f_out As Boolean

Private Sub CboCust_Change()
    If cboCust.ListIndex >= 0 Then
        lblCust.Text = cboCust.List(cboCust.ListIndex, 1)
        lblErrMsg.Caption = ""
    Else
        lblCust.Text = ""
    End If
End Sub

Private Sub cmdClear_Click()
    clear
End Sub

Private Sub cmdSearch_Click()
    lblErrMsg.Caption = ""
    
    Me.MousePointer = vbHourglass
    gridLoad
    
    Me.MousePointer = vbDefault
    
End Sub

Private Sub CmdSubmit_Click()
On Error GoTo ErrHandler
    
    Dim i As Long
    Dim sqlUpdate As String
    
    ' Cek hak update
    If hakUpdate(Me.Name) = 0 Then
        lblErrMsg.Caption = DisplayMsg(3008) ' No permission to update
        Exit Sub
    End If
    
    ' Cek kalau tidak ada data
    If f_out = True Then
        lblErrMsg.Caption = DisplayMsg(5012) ' There is no data to submit
        Exit Sub
    End If
    
    ' Loop grid
    With grid
        For i = 1 To .Rows - 1
            ' Tentukan status berdasarkan checkbox
            If .Cell(flexcpChecked, i, bteColStatus) = flexChecked Then
                sqlUpdate = "UPDATE PO_Contract_Master " & _
                            "SET Contract_Status='01', " & _
                            "Last_Update=GETDATE(), " & _
                            "Last_User='" & userLogin & "' " & _
                            "WHERE Contract_No='" & .TextMatrix(i, bteColContractNo) & "'"
            Else
                sqlUpdate = "UPDATE PO_Contract_Master " & _
                            "SET Contract_Status='02', " & _
                            "Last_Update=GETDATE(), " & _
                            "Last_User='" & userLogin & "' " & _
                            "WHERE Contract_No='" & .TextMatrix(i, bteColContractNo) & "'"
            End If
            
            ' Eksekusi SQL
            Db.Execute sqlUpdate
        Next i
    End With
    
    lblErrMsg.Caption = DisplayMsg(1101) ' Update Data Success !
    Exit Sub
    
' =========================
ErrHandler:
    lblErrMsg.Caption = "Error: " & err.Description
End Sub

Private Sub Form_Load()
 If gb_Simulation = True Then Call up_InitSimulation(Me)
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
    
    clear

With Anchor1
  .RegString = "AnchorCtrl,Positions," & Me.Name & "0|0"
  .DoInit
End With

End Sub

Sub clear()
    AddToComboSupplier
    AddToComboStatus
    
    DTFrom.Value = Format(Now, "dd MMM yyyy")
    DTTo.Value = Format(Now, "dd MMM yyyy")
    
    headerGrid
End Sub

Private Sub headerGrid()
    bteColContractNo = 0
    bteColContractDate = 1
    bteColItemCode = 2
    bteColItemName = 3
    bteColPONo = 4
    bteColSupplierCode = 5
    bteColSupplierName = 6
    bteColQtyPO = 7
    bteColQtyReceipt = 8
    bteColStatus = 9
    bteColTot = 10

    With grid
        .ColS = bteColTot
        .clear
        
        .ColDataType(bteColStatus) = flexDTBoolean
    
        .Rows = 1
       
        .TextMatrix(0, bteColContractNo) = "Contract No."
        .TextMatrix(0, bteColContractDate) = "Contract Date"
        .TextMatrix(0, bteColItemCode) = "Item Code"
        .TextMatrix(0, bteColItemName) = "Item Name"
        .TextMatrix(0, bteColPONo) = "PO No."
        .TextMatrix(0, bteColSupplierCode) = "Supplier Code"
        .TextMatrix(0, bteColSupplierName) = "Supplier Name"
        .TextMatrix(0, bteColQtyPO) = "Qty PO"
        .TextMatrix(0, bteColQtyReceipt) = "Qty Receipt"
         .TextMatrix(0, bteColStatus) = "Status"
       
        .ColWidth(bteColContractNo) = 1750
        .ColWidth(bteColContractDate) = 1500
        .ColWidth(bteColItemCode) = 1500
        .ColWidth(bteColItemName) = 2500
        .ColWidth(bteColPONo) = 2500
        .ColWidth(bteColSupplierCode) = 1750
        .ColWidth(bteColSupplierName) = 3200
        .ColWidth(bteColQtyPO) = 1000
        .ColWidth(bteColQtyReceipt) = 1200
        .ColWidth(bteColStatus) = 1000
        
        
        For i = 0 To .ColS - 1
            .ColAlignment(i) = flexAlignLeftTop ' default: kiri
        Next i

        .ColAlignment(bteColContractNo) = flexAlignLeftTop
        .ColAlignment(bteColContractDate) = flexAlignCenterCenter
        .ColAlignment(bteColItemCode) = flexAlignLeftTop
        .ColAlignment(bteColItemName) = flexAlignLeftTop
        .ColAlignment(bteColSupplierCode) = flexAlignLeftTop
        .ColAlignment(bteColPONo) = flexAlignLeftTop
        .ColAlignment(bteColSupplierCode) = flexAlignLeftTop
        .ColAlignment(bteColSupplierName) = flexAlignLeftTop
        .ColAlignment(bteColQtyPO) = flexAlignRightCenter
        .ColAlignment(bteColQtyReceipt) = flexAlignRightCenter
        
        For i = 0 To .ColS - 1
            .Row = 0
            .Col = i
            .CellAlignment = flexAlignCenterCenter
        Next i
        
        .Editable = flexEDKbdMouse
        
    End With
End Sub

Sub AddToComboSupplier()
    
    Dim sqlcust As String
    Dim RsCust As New Recordset

    sqlcust = "SELECT 'ALL'Trade_Code, 'ALL' Trade_Name UNION ALL  SELECT RTRIM(Trade_Code) Trade_Code, RTRIM(Trade_Name) Trade_Name FROM Trade_Master " & _
        "WHERE Trade_Cls = '2' OR Trade_Cls = '3'"
        
    Set RsCust = Db.Execute(sqlcust)
    
    With cboCust
        .clear
        .columnCount = 2
        .ColumnWidths = "50pt;275pt"
        .ListWidth = 325
        .ListRows = 15
       
        i = 0
        Do While Not RsCust.EOF
            .AddItem
            .List(i, 0) = Trim(RsCust("Trade_Code"))
            .List(i, 1) = IIf(IsNull(RsCust("Trade_Name")), " ", Trim(RsCust("Trade_Name")))
            
            RsCust.MoveNext
            i = i + 1
        Loop
         .ListIndex = 0
         
        RsCust.Close
    End With
    
End Sub

Sub AddToComboStatus()
    
    Dim sqlcust As String
    Dim RsStatus As New Recordset

    sqlcust = "SELECT 'ALL'Status UNION ALL  SELECT Status =  Status_Cls + ' - '+  RTRIM(Description) FROM dbo.Status_Cls "
        
    Set RsStatus = Db.Execute(sqlcust)
    
    With cboStatus
        .clear
        .columnCount = 1
        .ColumnWidths = "75pt"
        .ListWidth = 75
        .ListRows = 3
        
        i = 0
        Do While Not RsStatus.EOF
            .AddItem
            .List(i, 0) = Trim(RsStatus("Status"))
            
            RsStatus.MoveNext
            i = i + 1
        Loop
        
        .ListIndex = 0
        RsStatus.Close
    End With
    
End Sub

Private Sub CmdSubMenu_Click()
DoEvents
    frmMainMenu.Show
    DoEvents
    Unload Me
End Sub

Private Sub gridLoad()
    Dim RsIsiG As New ADODB.Recordset
    Dim sql As String
    Dim i As Long
    Dim lsStatus As String

        
    Call headerGrid
    
    sql = "EXEC dbo.sp_POContractInq_Sel " & vbCrLf & _
          " @SupplierCode = '" & cboCust.Text & "'," & vbCrLf & _
          " @DTFrom = '" & DTFrom.Value & "'," & vbCrLf & _
          " @DTTo = '" & DTTo.Value & "'," & vbCrLf & _
          " @CompleteStatus = '" & Left(cboStatus.Text, 3) & "'"

    If RsIsiG.State = 1 Then RsIsiG.Close
    RsIsiG.Open sql, Db, adOpenKeyset, adLockOptimistic

    If Not RsIsiG.EOF Then
        i = 0
        
        Do While Not RsIsiG.EOF
            i = i + 1
            grid.Rows = i + 1
            
            grid.TextMatrix(i, bteColContractNo) = Trim(RsIsiG("Contract_No") & "")
            grid.TextMatrix(i, bteColContractDate) = Trim(RsIsiG("Contract_Date") & "")
            grid.TextMatrix(i, bteColItemCode) = Trim(RsIsiG("Item_Code") & "")
            grid.TextMatrix(i, bteColItemName) = Trim(RsIsiG("Item_Name") & "")
            grid.TextMatrix(i, bteColPONo) = Trim(RsIsiG("PO_No") & "")
            grid.TextMatrix(i, bteColSupplierCode) = Trim(RsIsiG("Supplier_Code") & "")
            grid.TextMatrix(i, bteColSupplierName) = Trim(RsIsiG("Supplier_Name") & "")
            grid.TextMatrix(i, bteColQtyPO) = RsIsiG("Qty_PO")
            grid.TextMatrix(i, bteColQtyReceipt) = RsIsiG("Qty_Receipt")
                        
            If RsIsiG("Contract_Status") = "01" Then
                grid.Cell(flexcpChecked, i, bteColStatus) = flexChecked
            Else
                grid.Cell(flexcpChecked, i, bteColStatus) = flexUnchecked
            End If
            
            RsIsiG.MoveNext
        Loop
    Else
        lblErrMsg = DisplayMsg(8012)
    End If
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
   If Col = bteColStatus Then
        Cancel = False
    Else
        Cancel = True
    End If
End Sub
