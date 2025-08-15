VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmPOContractInquiry 
   BackColor       =   &H00FDDFE3&
   Caption         =   "PO Contract Inquiry"
   ClientHeight    =   10590
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14520
   Icon            =   "FrmPOContractInquiry.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   15615
   ScaleWidth      =   28560
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
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
      Index           =   0
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   14
      Tag             =   "TFFT*/"
      Top             =   9840
      Width           =   1140
   End
   Begin VB.CommandButton cmdexcel 
      BackColor       =   &H0080FFFF&
      Caption         =   "E&xcel"
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
      Left            =   13080
      Style           =   1  'Graphical
      TabIndex        =   13
      Tag             =   "FFTT*/"
      Top             =   9840
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   675
      Left            =   240
      TabIndex        =   11
      Tag             =   "TFTT*/"
      Top             =   9000
      Width           =   14055
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
         TabIndex        =   12
         Tag             =   "TFTF*/"
         Top             =   240
         Width           =   13725
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   1335
      Left            =   240
      TabIndex        =   2
      Tag             =   "TTTF*/"
      Top             =   960
      Width           =   14160
      Begin VB.CommandButton cmdsearch 
         BackColor       =   &H0080FFFF&
         Caption         =   "Sea&rch"
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
         Left            =   9360
         Style           =   1  'Graphical
         TabIndex        =   3
         Tag             =   "TTFF*/"
         Top             =   720
         Width           =   1140
      End
      Begin MSComCtl2.DTPicker DTPFrom 
         Height          =   315
         Left            =   1680
         TabIndex        =   4
         Tag             =   "TTFF*/"
         Top             =   720
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
         Format          =   128122883
         CurrentDate     =   37810
      End
      Begin MSComCtl2.DTPicker DTPTo 
         Height          =   315
         Left            =   3720
         TabIndex        =   5
         Tag             =   "TTFF*/"
         Top             =   720
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
         Format          =   128122883
         CurrentDate     =   37810
      End
      Begin VB.Line Line2 
         X1              =   3735
         X2              =   9255
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
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
         Left            =   3345
         TabIndex        =   10
         Tag             =   "TTFF*/"
         Top             =   780
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier CD"
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
         Left            =   360
         TabIndex        =   9
         Tag             =   "TTFF*/"
         Top             =   315
         Width           =   1035
      End
      Begin MSForms.ComboBox cbosupplier 
         Height          =   315
         Left            =   1680
         TabIndex        =   8
         Tag             =   "TTFF*/"
         Top             =   240
         Width           =   1950
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "3440;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblname 
         BackStyle       =   0  'Transparent
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
         Left            =   3735
         TabIndex        =   7
         Tag             =   "TTTF*/"
         Top             =   300
         Width           =   5535
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PO Date"
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
         Left            =   360
         TabIndex        =   6
         Tag             =   "TTFF*/"
         Top             =   780
         Width           =   705
      End
   End
   Begin EZRunnerv3.Anchor Anchor1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   820
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   420
      Left            =   12360
      TabIndex        =   1
      TabStop         =   0   'False
      Tag             =   "FTTF*/"
      Top             =   240
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   741
   End
   Begin MSComDlg.CommonDialog cdlReport 
      Left            =   -1290
      Top             =   2820
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   6420
      Left            =   240
      TabIndex        =   15
      TabStop         =   0   'False
      Tag             =   "TTTT*/"
      Top             =   2400
      Width           =   14145
      _cx             =   24950
      _cy             =   11324
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
      FocusRect       =   5
      HighLight       =   2
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   3
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   0
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
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   1
      ShowComboButton =   1
      WordWrap        =   -1  'True
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
   Begin VB.Label Label3 
      Alignment       =   2  'Center
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
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Tag             =   "TTTF*/"
      Top             =   360
      Width           =   14340
   End
End
Attribute VB_Name = "FrmPOContractInquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sql As String
Dim RS As New ADODB.Recordset

Dim bteID As Byte
Dim btePONo As Byte
Dim bteSuppCode As Byte
Dim bteSuppName As Byte
Dim btePODate As Byte
Dim bteDelDate As Byte
Dim bteItemCode As Byte
Dim bteItemName As Byte
Dim bteQty As Byte
Dim bteQtyRemainingPO As Byte
Dim bteQtyRemainingReceipt As Byte
Dim bteQtyReceipt As Byte
Dim btePrice As Byte
Dim bteAmount As Byte
Dim TotCol As Byte
Dim btePPN As Byte
Dim bteTotalAmount As Byte
Dim bteCount As Byte
Dim tempPOBefore As String

Sub Header()
   
    bteID = 0
    btePONo = 1
    bteSuppCode = 2
    bteSuppName = 3
    btePODate = 4
    bteDelDate = 5
    bteItemCode = 6
    bteItemName = 7
    bteQty = 8
    bteQtyRemainingPO = 9
    bteQtyReceipt = 10
    bteQtyRemainingReceipt = 11
    btePrice = 12
    bteAmount = 13
    TotCol = 14

    With Grid
        .clear
        .Rows = 2
        .ColS = TotCol

        .TextMatrix(0, bteID) = "ID"
        .TextMatrix(0, btePONo) = "PO No"
        .TextMatrix(0, bteSuppCode) = "Supplier Code"
        .TextMatrix(0, bteSuppName) = "Supplier Name"
        .TextMatrix(0, btePODate) = "PO Date"
        .TextMatrix(0, bteDelDate) = "Delivery Date"
        .TextMatrix(0, bteItemCode) = "Item Code"
        .TextMatrix(0, bteItemName) = "Item Name"
        .TextMatrix(0, bteQty) = "Qty"
        .TextMatrix(0, bteQtyRemainingPO) = "Qty Remaining PO"
        .TextMatrix(0, bteQtyReceipt) = "Qty Receipt"
        .TextMatrix(0, bteQtyRemainingReceipt) = "Qty Remaining Receipt"
        .TextMatrix(0, btePrice) = "Price"
        .TextMatrix(0, bteAmount) = "Amount"

        .TextMatrix(1, bteID) = "ID"
        .TextMatrix(1, btePONo) = "PO No"
        .TextMatrix(1, bteSuppCode) = "Supplier Code"
        .TextMatrix(1, bteSuppName) = "Supplier Name"
        .TextMatrix(1, btePODate) = "PO Date"
        .TextMatrix(1, bteDelDate) = "Delivery Date"
        .TextMatrix(1, bteItemCode) = "Item Code"
        .TextMatrix(1, bteItemName) = "Item Name"
        .TextMatrix(1, bteQty) = "Qty"
        .TextMatrix(1, bteQtyRemainingPO) = "Qty Remaining PO"
        .TextMatrix(1, bteQtyReceipt) = "Qty Receipt"
        .TextMatrix(1, bteQtyRemainingReceipt) = "Qty Remaining Receipt"
        .TextMatrix(1, btePrice) = "Price"
        .TextMatrix(1, bteAmount) = "Amount"

        .ColWidth(bteID) = 1500
        .ColWidth(btePONo) = 1800
        .ColWidth(bteSuppCode) = 1500
        .ColWidth(bteSuppName) = 2500
        .ColWidth(btePODate) = 1500
        .ColWidth(bteDelDate) = 1500
        .ColWidth(bteItemCode) = 1500
        .ColWidth(bteItemName) = 3000
        .ColWidth(bteQty) = 1000
        .ColWidth(bteQtyRemainingPO) = 1500
        .ColWidth(bteQtyReceipt) = 1000
        .ColWidth(bteQtyRemainingReceipt) = 1500
        .ColWidth(btePrice) = 1300
        .ColWidth(bteAmount) = 1300

        .ColDataType(btePODate) = flexDTDate
        .ColDataType(btePONo) = flexDTDate

        .MergeRow(bteID) = True

        For i = 0 To bteAmount
            .MergeCol(i) = True
        Next i

        .MergeCells = flexMergeFixedOnly

'        .FixedCols = 2
        .ColAlignment(bteID) = flexAlignLeftCenter
        .ColAlignment(btePONo) = flexAlignLeftCenter
        .ColAlignment(bteDelDate) = flexAlignCenterCenter
        .ColAlignment(btePODate) = flexAlignCenterCenter
        .ColAlignment(bteItemCode) = flexAlignLeftCenter

        '.Cell(flexcpAlignment, bteID, btePONo, bteSuppCode,) = flexAlignCenterCenter
        .Cell(flexcpAlignment, 0, 0, 0, .ColS - 1) = flexAlignCenterCenter
        .Cell(flexcpAlignment, 1, 1, 1, .ColS - 1) = flexAlignCenterCenter

        .ColHidden(bteID) = True

        .RowHeight(0) = 225
        .RowHeight(1) = 225

    End With

    lblErrMsg.Caption = ""
 
End Sub

Private Sub CboSupplier_Change()
    
    If cboSupplier.ListIndex <> -1 Then
        lblname.Caption = cboSupplier.Column(1)
    End If
    
End Sub

Private Sub CmdExcel_Click()
If Grid.Rows > 1 Then
    up_Excel
Else
    lblErrMsg.Caption = DisplayMsg("0013")
End If
    
End Sub

Private Sub cmdSearch_Click()

    If cboSupplier.Text = "" Then
        lblErrMsg = DisplayMsg(1054)
        cboSupplier.SetFocus
        Exit Sub
    ElseIf CDate(DTPTo) < CDate(DTPFrom) Then
        lblErrMsg.Caption = DisplayMsg(4077) & " " & Format(DTPTo, "dd MMM yyyy")   '"delivery Date must be higher than "
        DTPTo.SetFocus
        Exit Sub
    End If
    
    If cboSupplier.Text <> "" Then
      cboSupplier.MatchEntry = 1
      cboSupplier.Text = cboSupplier.Text
      
        If cboSupplier.MatchFound = False Then
          lblErrMsg = DisplayMsg(4050)
          cboSupplier.SetFocus
          cboSupplier.MatchEntry = 2
          Exit Sub
        End If
      
      cboSupplier.MatchEntry = 2
    End If
    
    Browse
End Sub

Private Sub Command1_Click(Index As Integer)
Dim top As Integer
lblErrMsg.Caption = ""

    Select Case Index
    Case 0: Unload Me
            frmMainMenu.Show
   
    End Select
End Sub

Private Sub Form_Load()

    Header
    
    up_FillCombo
        
    DTPFrom.Value = Format(Now, "yyyy-mm-01")
    DTPTo.Value = Format(Now, "dd MMM yyyy")
    
    With Anchor1
      .RegString = "AnchorCtrl,Positions," & Me.Name & "0|0"
      .DoInit
    End With
    
End Sub

Sub up_FillCombo()
    sql = "SELECT Trade_Code, Trade_Name FROM Trade_Master where trade_cls='2' or trade_cls='3'"
    Set RS = Db.Execute(sql)
    
    With cboSupplier
    .clear
    .columnCount = 2
    .ColumnWidths = "80 pt;300 pt"
    .ListWidth = 380
    .ListRows = 15
    .AddItem ""
    .List(0, 0) = strAll
    .List(0, 1) = strAll
    i = 1
    
    Do Until RS.EOF
        .AddItem ""
        .List(i, 0) = Trim(RS!Trade_Code)
        .List(i, 1) = Trim(RS!trade_name)
        i = i + 1
        RS.MoveNext
    Loop
    .ListIndex = 0
    End With
End Sub

Private Sub Browse()
    Dim RS As ADODB.Recordset
    Dim cmd As ADODB.Command
     
Me.MousePointer = vbHourglass

    Header
      
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_POContractInquiry_Sel"

    cmd.Parameters.append cmd.CreateParameter("SuppCode", adVarChar, adParamInput, 6, cboSupplier.Text)
    cmd.Parameters.append cmd.CreateParameter("PODateFrom", adDBTime, adParamInput, , DTPFrom.Value)
    cmd.Parameters.append cmd.CreateParameter("PODateTo", adDBTime, adParamInput, , DTPTo.Value)

    Set RS = cmd.Execute
    
    
    If (RS.BOF And RS.EOF) Then
        lblErrMsg.Caption = DisplayMsg(4006)
        Me.MousePointer = vbDefault
        Exit Sub
    Else
        
        i = 2
        
        With Grid
            Do While Not RS.EOF
                .Rows = .Rows + 1
                                
                .TextMatrix(i, btePONo) = Trim(RS("PO_No"))
                .TextMatrix(i, bteSuppCode) = Trim(RS("Supplier_Code"))
                .TextMatrix(i, bteSuppName) = Trim(RS("Trade_Name") & "")
                .TextMatrix(i, btePODate) = Format(RS("PO_Date"), "dd mmm yyyy")
                .TextMatrix(i, bteDelDate) = Format(RS("delivery_date"), "dd mmm yyyy")
                .TextMatrix(i, bteItemCode) = Trim(RS("Item_Code") & "")
                .TextMatrix(i, bteItemName) = Trim(RS("Item_Name") & "")
                .TextMatrix(i, bteQty) = Format(RS("Qty"), gs_formatQty)
                .TextMatrix(i, bteQtyRemainingPO) = Format(RS("QtyRemaining_PO"), gs_formatQty)
                .TextMatrix(i, bteQtyReceipt) = Format(RS("Qty_Receipt"), gs_formatQty)
                .TextMatrix(i, bteQtyRemainingReceipt) = Format(RS("QtyRemaining_Receipt"), gs_formatQty)
                .TextMatrix(i, btePrice) = Format(RS("Price"), gs_formatPriceIDR)
                .TextMatrix(i, bteAmount) = Format(RS("Amount"), gs_formatAmountIDR)
                               
                i = i + 1
                
                RS.MoveNext
                                
            Loop
        End With
    
    End If
           
    Me.MousePointer = vbDefault
    
End Sub

Private Sub up_Excel()
    Dim RS As ADODB.Recordset
    Dim cmd As ADODB.Command
    
    
    Dim xlapp As New Excel.application
    Dim Idx As Long
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    
    Dim strFileName As String
    
    Me.MousePointer = vbHourglass
    On Error GoTo errHandler
    
    lblErrMsg.Caption = ""
                           
        Me.MousePointer = vbHourglass
                
        Set cmd = New ADODB.Command
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 0
        cmd.ActiveConnection = Db
        cmd.CommandText = "sp_POContractInquiry_Sel"
        
        cmd.Parameters.append cmd.CreateParameter("SuppCode", adVarChar, adParamInput, 6, cboSupplier.Text)
        cmd.Parameters.append cmd.CreateParameter("PODateFrom", adDBTime, adParamInput, , DTPFrom.Value)
        cmd.Parameters.append cmd.CreateParameter("PODateTo", adDBTime, adParamInput, , DTPTo.Value)
    
        Set RS = cmd.Execute
        
        If (RS.BOF And RS.EOF) Then
            lblErrMsg.Caption = DisplayMsg(4006)
            Me.MousePointer = vbDefault
        Exit Sub
    Else
    
        With xlapp
    
            .Workbooks.Add
                
            .Range("A1:A6").EntireRow.Insert
            .Range("A2:M2").Merge
            .Range("A2:M4").horizontalAlignment = xlHAlignCenter
            .Range("A6:M6").Font.Bold = True
            .Range("A6:M6").horizontalAlignment = xlHAlignCenter

            .Range("A2").Font.Bold = True
            .Range("A2").Font.Size = 10
            .Range("A2") = "Purchase Order Contract Inquiry"
            
            .Range("A3:B3").Font.Bold = True
            .Range("A3:B3").Font.Size = 8
            .Range("A3") = "Supplier Code"
            .Range("B3") = ": " & cboSupplier.Text
            .Range("A3:B3").horizontalAlignment = xlLeft
            
            .Range("A4:B4").Font.Bold = True
            .Range("A4:B4").Font.Size = 8
            .Range("A4") = "Delivery Date"
            .Range("B4") = ": " & Format(DTPFrom.Value, "[$-409]d-mmm-yyyy;@") & " To " & Format(DTPTo.Value, "[$-409]d-mmm-yyyy;@")
            .Range("A4:B4").horizontalAlignment = xlLeft
                       
            '********************Garis Header****************************
            .Range("A6:M6").Borders(xlEdgeTop).LineStyle = xlContinuous
            .Range("A6:M6").Borders(xlEdgeBottom).LineStyle = xlDouble
            '***********************************************************
            
            '********************Garis Footer************************************************************
            .Range("A" & 6 + Grid.Rows - 2 & ":M" & 6 + Grid.Rows - 2).Borders(xlEdgeBottom).LineStyle = xlDouble
            '********************************************************************************************
            
            .Range("A6:M6").horizontalAlignment = xlHAlignCenter
                        
            .Range("A6") = "PO No"
            .Range("B6") = "Supplier Code"
            .Range("C6") = "Supplier Name"
            .Range("D6") = "PO Date"
            .Range("E6") = "Delivery Date"
            .Range("F6") = "Item Code"
            .Range("G6") = "Item Name"
            .Range("H6") = "Qty"
            .Range("I6") = "Qty Remaining PO"
            .Range("J6") = "Qty Receipt"
            .Range("K6") = "Qty Remaining Receipt"
            .Range("L6") = "Price"
            .Range("M6") = "Amount"
                        
            '******Complete*********
            Dim Row As Integer, i As Integer
            Row = 1
            i = 6
            Do While i <= 4 + Grid.Rows
                .Range("A" & i) = Grid.TextMatrix(Row, btePONo)
                .Range("B" & i) = Grid.TextMatrix(Row, bteSuppCode)
                .Range("C" & i) = Grid.TextMatrix(Row, bteSuppName)
                .Range("D" & i) = Grid.TextMatrix(Row, btePODate)
                .Range("E" & i) = Grid.TextMatrix(Row, bteDelDate)
                .Range("F" & i) = Grid.TextMatrix(Row, bteItemCode)
                .Range("G" & i) = Grid.TextMatrix(Row, bteItemName)
                .Range("H" & i) = Grid.TextMatrix(Row, bteQty)
                .Range("I" & i) = Grid.TextMatrix(Row, bteQtyRemainingPO)
                .Range("J" & i) = Grid.TextMatrix(Row, bteQtyReceipt)
                .Range("K" & i) = Grid.TextMatrix(Row, bteQtyRemainingReceipt)
                .Range("L" & i) = Grid.TextMatrix(Row, btePrice)
                .Range("M" & i) = Grid.TextMatrix(Row, bteAmount)
                              
                i = i + 1
                Row = Row + 1
            Loop
            '***********************
            
            .Range("A6", "M" & 6 + Grid.Rows - 1).Columns.Font.Name = "Arial"
            .Range("A6", "M" & 6 + Grid.Rows - 1).Columns.Font.Size = 8

            .Range("D7:D" & 7 + Grid.Rows - 1).NumberFormat = "[$-409]d-mmm-yyyy;@"
            .Range("E7:E" & 7 + Grid.Rows - 1).NumberFormat = "[$-409]d-mmm-yyyy;@"
            .Range("A7:g" & 7 + Grid.Rows - 1).horizontalAlignment = xlLeft
            .Range("H7:M" & 7 + Grid.Rows - 1).NumberFormat = gs_formatQty
            .Range("A:M").Columns.AutoFit
    
            .Visible = True
            .ActiveSheet.PageSetup.PaperSize = xlPaperA4
            .ActiveSheet.PageSetup.Orientation = 2
            .ActiveSheet.PageSetup.LeftMargin = application.InchesToPoints(0.4)
            .ActiveSheet.PageSetup.RightMargin = application.InchesToPoints(0.4)
            .WindowState = xlMaximized
        End With
            
        Screen.MousePointer = vbDefault
        MousePointer = vbDefault

    End If

ErrExit:
    Set RS = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
errHandler:
    lblErrMsg.Caption = "[" & err.number & "] " & err.Description
    err.clear
    Resume ErrExit
    
End Sub


