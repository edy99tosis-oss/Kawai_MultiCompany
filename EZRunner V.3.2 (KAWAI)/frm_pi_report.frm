VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_pi_report 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Inventory Report"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15120
   Icon            =   "frm_pi_report.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Cmd_Find 
      BackColor       =   &H0080FFFF&
      Caption         =   "Find [F3]"
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
      Left            =   13770
      Style           =   1  'Graphical
      TabIndex        =   20
      Tag             =   "FFTF*/"
      Top             =   1620
      Width           =   1155
   End
   Begin VB.TextBox txtItemCode 
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
      Left            =   11490
      TabIndex        =   19
      Tag             =   "FFTF*/"
      Top             =   1620
      Width           =   2175
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   13050
      TabIndex        =   18
      Tag             =   "FTTF*/"
      Top             =   360
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   714
   End
   Begin VB.CommandButton Cmd_Save 
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
      Index           =   1
      Left            =   12810
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "FFTT*/"
      Top             =   9840
      Width           =   1035
   End
   Begin VB.CommandButton Cmd_Save 
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
      Index           =   9
      Left            =   3750
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "TTFF*/"
      Top             =   1560
      Width           =   1035
   End
   Begin VB.CommandButton Cmd_Save 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Sub &Menu"
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
      Left            =   270
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "TFFT*/"
      Top             =   9840
      Width           =   1125
   End
   Begin VB.CommandButton Cmd_Save 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Last Page"
      Enabled         =   0   'False
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
      Index           =   7
      Left            =   6450
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      Tag             =   "TFFT*/"
      Top             =   9840
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.CommandButton Cmd_Save 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Next Page"
      Enabled         =   0   'False
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
      Index           =   6
      Left            =   5130
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      Tag             =   "TFFT*/"
      Top             =   9840
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.CommandButton Cmd_Save 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Prev Page"
      Enabled         =   0   'False
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
      Index           =   5
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Tag             =   "TFFT*/"
      Top             =   9840
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.CommandButton Cmd_Save 
      BackColor       =   &H00C0E0FF&
      Caption         =   "First Page"
      Enabled         =   0   'False
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
      Index           =   4
      Left            =   2535
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Tag             =   "TFFT*/"
      Top             =   9840
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   255
      TabIndex        =   6
      Tag             =   "TFTT*/"
      Top             =   9060
      Width           =   14685
      Begin VB.Label LblErrMsg 
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
         Left            =   75
         TabIndex        =   7
         Tag             =   "TFTF*/"
         Top             =   225
         Width           =   14430
      End
   End
   Begin VB.CommandButton Cmd_Save 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Print"
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
      Left            =   13920
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "FFTT*/"
      Top             =   9840
      Width           =   1035
   End
   Begin MSComCtl2.DTPicker DMonth 
      Height          =   315
      Left            =   2310
      TabIndex        =   1
      Tag             =   "TTFF*/"
      Top             =   1590
      Width           =   1290
      _ExtentX        =   2275
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
      CustomFormat    =   "MMM yyyy"
      Format          =   141230083
      UpDown          =   -1  'True
      CurrentDate     =   37798
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   6705
      Left            =   255
      TabIndex        =   12
      Tag             =   "TTTT*/"
      Top             =   2205
      Width           =   14685
      _cx             =   25903
      _cy             =   11827
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
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
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
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Find Item Code"
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
      Left            =   9930
      TabIndex        =   21
      Tag             =   "FFTF*/"
      Top             =   1680
      Width           =   1485
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Inventory Report"
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
      Left            =   270
      TabIndex        =   17
      Tag             =   "TFTF*/"
      Top             =   315
      Width           =   14610
   End
   Begin VB.Line Line1 
      X1              =   6675
      X2              =   9690
      Y1              =   1395
      Y2              =   1395
   End
   Begin VB.Label LblLocationName 
      BackStyle       =   0  'Transparent
      Caption         =   "LblLocationName"
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
      Left            =   6675
      TabIndex        =   16
      Tag             =   "TTFF*/"
      Top             =   1170
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Warehouse Name"
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
      Left            =   4995
      TabIndex        =   15
      Tag             =   "TTFF*/"
      Top             =   1170
      Width           =   1605
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date (Month)"
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
      Left            =   585
      TabIndex        =   14
      Tag             =   "TTFF*/"
      Top             =   1620
      Width           =   1335
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "WareHouse CD"
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
      Left            =   585
      TabIndex        =   13
      Tag             =   "TTFF*/"
      Top             =   1170
      Width           =   1335
   End
   Begin MSForms.ComboBox CboLocationCD 
      Height          =   315
      Left            =   2310
      TabIndex        =   0
      Tag             =   "TTFF*/"
      Top             =   1140
      Width           =   2370
      VariousPropertyBits=   612386843
      MaxLength       =   15
      DisplayStyle    =   3
      Size            =   "4180;556"
      ListRows        =   15
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "frm_pi_report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim dateUp As Date

Dim bteColProdCod As Byte
Dim bteColPartNo As Byte
Dim bteColDesc As Byte
Dim bteColAddress As Byte
Dim bteColPreMonth As Byte
Dim bteColReceipt As Byte
Dim bteColSupply As Byte
Dim bteColLossReject As Byte
Dim bteColEnd As Byte
Dim bteColInventory As Byte
Dim bteColGroupDesc As Byte
Dim bteColUnit As Byte
Dim bteColReason As Byte
Dim bteColStockStatus As Byte
Dim bteColPO As Byte
Dim bteColBeginTotal As Byte
Dim bteColIncoming As Byte
Dim bteColReq As Byte
Dim bteColGrandTotal As Byte
Dim bteColGrandTotal2 As Byte

Dim bytSort As Byte

Private Sub Header()

    bteColProdCod = 0
    bteColPartNo = 1
    bteColDesc = 2
    bteColGroupDesc = 3
    bteColUnit = 4
    bteColAddress = 5
    bteColPreMonth = 6
    bteColReceipt = 7
    bteColSupply = 8
    bteColLossReject = 9
    bteColEnd = 10
    bteColInventory = 11
    bteColReason = 12
    bteColStockStatus = 13
    bteColBeginTotal = 14
    bteColPO = 15
    bteColIncoming = 16
    bteColReq = 17
    bteColGrandTotal = 18
    bteColGrandTotal2 = 19
    
    grid.Rows = 1
    grid.ColS = 20
    
    grid.TextMatrix(0, bteColProdCod) = "Product Code"
    grid.TextMatrix(0, bteColPartNo) = "Part Number"
    grid.TextMatrix(0, bteColDesc) = "Description"
    grid.TextMatrix(0, bteColAddress) = "Address"
    grid.TextMatrix(0, bteColPreMonth) = "Pre Month Stock"
    grid.TextMatrix(0, bteColReceipt) = "Receipt Total"
    grid.TextMatrix(0, bteColSupply) = "Supply Total"
    grid.TextMatrix(0, bteColLossReject) = "Loss/Reject"
    grid.TextMatrix(0, bteColEnd) = "End of Month Stock"
    grid.TextMatrix(0, bteColInventory) = "Inventory"
    grid.TextMatrix(0, bteColGroupDesc) = "Group Desc"
    grid.TextMatrix(0, bteColUnit) = "Unit"
    grid.TextMatrix(0, bteColReason) = "Reason"
    grid.TextMatrix(0, bteColStockStatus) = "Stock Status"
    grid.TextMatrix(0, bteColPO) = "Purchase Order"
    grid.TextMatrix(0, bteColBeginTotal) = "Premonth Total"
    grid.TextMatrix(0, bteColIncoming) = "Incoming Total"
    grid.TextMatrix(0, bteColReq) = "Requirement"
    grid.TextMatrix(0, bteColGrandTotal) = "Grand Total(Incoming)"
    grid.TextMatrix(0, bteColGrandTotal2) = "Grand Total(PO)"
    
    
    grid.ColWidth(bteColProdCod) = 1400
    grid.ColWidth(bteColPartNo) = 2400
    grid.ColWidth(bteColDesc) = 3500
    grid.ColWidth(bteColAddress) = 800
    grid.ColWidth(bteColPreMonth) = 1500
    grid.ColWidth(bteColReceipt) = 1250
    grid.ColWidth(bteColSupply) = 1250
    grid.ColWidth(bteColLossReject) = 1200
    grid.ColWidth(bteColEnd) = 1800
    grid.ColWidth(bteColInventory) = 1500
    grid.ColWidth(bteColGroupDesc) = 1200
    grid.ColWidth(bteColUnit) = 600
    grid.ColWidth(bteColReason) = 3500
    grid.ColWidth(bteColStockStatus) = 1500
    grid.ColWidth(bteColReq) = 1200
    grid.ColWidth(bteColPO) = 1200
    grid.ColWidth(bteColGrandTotal) = 1200
    grid.ColWidth(bteColGrandTotal2) = 1200
    
    grid.ColAlignment(bteColProdCod) = flexAlignLeftCenter
    grid.ColAlignment(bteColPartNo) = flexAlignLeftCenter
    grid.ColAlignment(bteColDesc) = flexAlignLeftCenter
    grid.ColAlignment(bteColAddress) = flexAlignLeftCenter
    grid.ColAlignment(bteColPreMonth) = flexAlignRightCenter
    grid.ColAlignment(bteColReceipt) = flexAlignRightCenter
    grid.ColAlignment(bteColSupply) = flexAlignRightCenter
    grid.ColAlignment(bteColLossReject) = flexAlignRightCenter
    grid.ColAlignment(bteColEnd) = flexAlignRightCenter
    grid.ColAlignment(bteColInventory) = flexAlignRightCenter
    grid.ColAlignment(bteColGroupDesc) = flexAlignLeftCenter
    grid.ColAlignment(bteColUnit) = flexAlignLeftCenter
    grid.ColAlignment(bteColReason) = flexAlignLeftCenter
    grid.ColAlignment(bteColStockStatus) = flexAlignLeftCenter
    
    grid.ColHidden(bteColPartNo) = True
    grid.ColHidden(bteColGroupDesc) = True
    
    grid.ColHidden(bteColUnit) = True
    grid.ColHidden(bteColReason) = True
    grid.ColHidden(bteColStockStatus) = True
    grid.ColHidden(bteColPO) = True
    grid.ColHidden(bteColBeginTotal) = True
    grid.ColHidden(bteColIncoming) = True
    grid.ColHidden(bteColReq) = True
    grid.ColHidden(bteColGrandTotal) = True
    grid.ColHidden(bteColGrandTotal2) = True
    
    
    grid.FrozenCols = bteColPreMonth
    
    grid.Cell(flexcpAlignment, 0, 0, 0, bteColStockStatus) = flexAlignCenterCenter
End Sub

Private Sub CboLocationCD_Change()
Call clearGrid
If CboLocationCD.MatchFound Then
   LblLocationName = CboLocationCD.List(CboLocationCD.ListIndex, 1)
   LblErrMsg = ""
Else
   LblLocationName = ""
   LblErrMsg = DisplayMsg(4014) '"Location CD is not found !"
End If
End Sub
Sub clearGrid()
grid.clear
grid.Rows = 1
Call Header
End Sub

Private Sub CboLocationCD_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
Dim j As Integer
If KeyCode = 13 Then
Call clearGrid
  j = 0
For i = 0 To CboLocationCD.ListCount - 1
    If UCase(Trim(CboLocationCD)) = UCase(Trim(CboLocationCD.List(i, 0))) Then
        CboLocationCD = Trim(CboLocationCD.List(i, 0))
        LblLocationName = Trim(CboLocationCD.List(i, 1))
        j = 1: Exit For
    End If
Next

If j = 0 Then
    LblErrMsg = DisplayMsg(4018) '"Invalid warehouse code !"
    Exit Sub
Else
    LblErrMsg = ""
End If
End If
End Sub

Private Sub CboLocationCD_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Cmd_Find_Click()
LblErrMsg = ""
If txtItemCode.Text = "" Then
    LblErrMsg = DisplayMsg(1009)
    Exit Sub
ElseIf grid.Rows <= 1 Then
    LblErrMsg = "Please Search of Data!"
    Exit Sub
Else
    Call searchdigrid
End If
End Sub
Private Sub searchdigrid()
Dim lngRow As Long
    Dim strText As String
    Dim booFound As Boolean
    
    With grid
            
        For lngRow = .Row + 1 To .Rows - 1
            
            strText = Trim(.TextMatrix(lngRow, bteColProdCod))
            If InStr(1, UCase(strText), UCase(txtItemCode.Text)) <> 0 Then
                .Row = lngRow
                .TopRow = lngRow
                '.SelectionMode = flexSelectionByRow
                booFound = True
                LblErrMsg = ""
                Exit For
            End If
                        
        Next
        .SetFocus
        If Not booFound Then
            .Row = 1
            .TopRow = 1
            '.SelectionMode = flexSelectionFree
            LblErrMsg = DisplayMsg("8012")
        End If
        
    End With
End Sub
Private Sub Cmd_Save_Click(Index As Integer)
Dim j As Integer, dates As String
Dim selisih As Double

Me.MousePointer = vbHourglass

Select Case Index
        Case 1
            toExcel
       Case 8:
                
                frmMainMenu.Show
                
                Unload Me
       Case 9:
                Dim strSQL As String
                Dim i As Long
                If CboLocationCD.Text = "" Then
                   LblErrMsg = DisplayMsg(1042) '"Please choose warehouse !"
                   Me.MousePointer = vbDefault
                Else
                Me.MousePointer = vbHourglass
                    strSQL = "exec [sp_normalize_receipt_supply_BY_Warehouse] '" & Trim(CboLocationCD.Text) & "'"
                    Db.Execute strSQL
                
                    LblErrMsg = ""
                    
                       Call Header
                       grid.Rows = 1
                       Call Browse
                       For i = 4 To 7
                           Cmd_save(i).Enabled = True
                       Next i
                    Me.MousePointer = vbDefault
                    
                End If
        Case 0:
           '     exit sub
              Dim application As New CRAXDDRT.application
              Dim report As New CRAXDDRT.report
              Dim rsRpt As New ADODB.Recordset
              Dim Rpt As New FrmRpt3
              Dim sqlControl As String, RsInvControl As New ADODB.Recordset
              
                sqlControl = "select * from inventory_control where fix_cls='1' order by inventory_year,inventory_month"
                
                If RsInvControl.State <> adStateClosed Then RsInvControl.Close
                RsInvControl.Open sqlControl, Db, adOpenKeyset, adLockOptimistic
                
                If RsInvControl.EOF = True And RsInvControl.BOF = True Then LblErrMsg = "Inventory Stock hasn't been closed !": Exit Sub
                
                RsInvControl.MoveLast
                 
                 j = 0
                For i = 0 To CboLocationCD.ListCount - 1
                    If UCase(Trim(CboLocationCD)) = UCase(Trim(CboLocationCD.List(i, 0))) Then
                        CboLocationCD = Trim(CboLocationCD.List(i, 0))
                        LblLocationName = Trim(CboLocationCD.List(i, 1))
                        j = 1: Exit For
                    End If
                Next
                
                If j = 0 Then LblErrMsg = DisplayMsg(4018) '"Invalid warehouse code !": Exit Sub
        


            'LblErrMsg = up_ValidateDateRange(DMonth.Value, False)
            'If Trim(LblErrMsg) <> "" Then Exit Sub
            selisih = up_GetDateRange(DMonth.Value)
            
                            
            LblErrMsg = ""
            Me.MousePointer = vbHourglass
              
            If selisih = 0 Or selisih = 1 Or selisih = 2 Then
            
                  sql = "select gc.description groupDesc, uc.description Unit_Desc,descriptions=case isnull(im.sheetcoil_cls,0) when 0 then  im.item_name else rtrim(im.item_name) + ' (' + rtrim(sh.description) + ', T' + cast(im.thickness as varchar(15)) + ' x W' + cast(im.width as varchar(15)) + ' x L' + cast(im.length as varchar(15)) + ')'  end " & _
                            ",sm.*, rtrim(address) address, rtrim(makeritem_code) makeritem_code, isnull(wh_name,(select trade_name from trade_master where trade_code=sm.warehouse_code)) wh_name from stock_master sm left join warehouse_master wm on " & _
                            " sm.warehouse_code = wm.wh_code " & _
                            " left join item_master im on sm.item_code=im.item_code " & _
                            " left join sheetcoil_cls sh on im.sheetcoil_cls = sh.sheetcoil_cls " & _
                            " left join group_cls gc on im.group_cls=gc.group_cls " & _
                            " left join unit_cls uc on im.unit_cls=uc.unit_cls " & _
                            " where warehouse_code='" & Trim(CboLocationCD) & "'" '& _

                   sql = sql & " order by  warehouse_code,sm.item_Code"
                
                
                  If rsRpt.State <> adStateClosed Then rsRpt.Close
                  rsRpt.Open sql, Db, adOpenDynamic, adLockOptimistic
                
                  If rsRpt.EOF Then LblErrMsg.Caption = DisplayMsg(4006): Me.MousePointer = vbDefault: Exit Sub
                                        
                  Set report = application.OpenReport(App.path & "\Reports\rpt_pi_report.rpt")
                  report.Database.Tables(1).SetDataSource rsRpt
                ''#####################################################################
                ''# Qty Digit and decimal
                report.FormulaFields(5).Text = "" & gi_decimalDigitQty & ""
                report.FormulaFields(6).Text = "" & gi_decimalDigitQty & ""
                ''#####################################################################
                  
                   Select Case up_GetDateRange(DMonth.Value)
        
                    Case 0:
                            report.Sections(4).Suppress = False
                            report.Sections(5).Suppress = True
                            report.Sections(6).Suppress = True
        
                     Case 1:
        
                            report.Sections(4).Suppress = True
                            report.Sections(5).Suppress = False
                            report.Sections(6).Suppress = True
                                                                           
                     Case 2:
        
                            report.Sections(4).Suppress = True
                            report.Sections(5).Suppress = True
                            report.Sections(6).Suppress = False
                                                                                                              
                  End Select
                        
            Else
                 
                sql = " select gc.description groupDesc, uc.description Unit_Desc, sm.*,item_name,address, " & vbCrLf & _
                            "  description=(case isnull(im.sheetcoil_cls,'') when '' then rtrim(im.item_name)  else rtrim(cast(im.item_name as varchar(15))) + rtrim(cast(im.item_name as varchar(50)))+ rtrim(cast(',' as varchar(1))) + rtrim(cast('T' as varchar(1))) + rtrim(cast(IM.Thickness as varchar(18)))  + rtrim(cast('x' as varchar(1))) + rtrim(cast('W' as varchar(1))) + rtrim(cast(IM.width as varchar(18))) +  rtrim(cast(im.length as varchar(18))) end)   " & vbCrLf & _
                            " ,wh_name, makeritem_code,  " & vbCrLf & _
                            " PO = isnull((select sum(qty) from purchaseorder_detail PD inner join PurchaseOrder_master PM On PM.PO_no=PD.PO_NO where PM.Fix_cls='1' and month(PM.Delivery_date)='" & Format(DMonth.Month, "00") & "' and year(PM.Delivery_Date)='" & DMonth.Year & " ' and item_code=Sm.Item_Code),0), " & vbCrLf & _
                            " StokBegin = isnull((select sum(Premonth) from stock_history where item_code=sm.item_code and stock_month='" & DMonth.Month & "' and stock_year ='" & DMonth.Year & "' ),0), " & vbCrLf & _
                            " Incoming = isnull((select sum(Qty) from part_receipt where receipt_cls ='R' and Item_code=sm.item_code and month(receipt_date)='" & Format(DMonth.Month, "00") & "'),0), " & vbCrLf & _
                            " Req = isnull((select case when sum(childRequirement_qty)-sum(childRequirementResult_qty)-sum(offchildrequirement_qty)<0 then 0 else sum(childRequirement_qty) end as SisaReqQty  " & vbCrLf & _
                            " from requirement  " & vbCrLf & _
                            " where year(childrequirement_date) ='" & Format(DMonth, "yyyy") & "'  " & vbCrLf & _
                            "       and month(childrequirement_date) = '" & Format(DMonth.Month, "00") & "' " & vbCrLf & _
                            "       and (complete_cls is null or complete_cls<>'1')  " & vbCrLf & _
                            "       and childItem_code=sm.item_code),0) " & vbCrLf & _
                            " from stock_history sm  left join warehouse_master wm  " & vbCrLf & _
                            " on sm.warehouse_code=wm.wh_code  left join item_master im on im.item_code=sm.item_code   " & vbCrLf & _
                            " left join group_cls gc on im.group_cls=gc.group_cls   " & vbCrLf & _
                            " left join unit_cls uc on im.unit_cls=uc.unit_cls  " & vbCrLf & _
                            " left join  sheetcoil_cls sh on im.sheetcoil_cls=sh.sheetcoil_cls  " & vbCrLf & _
                            " where warehouse_code='" & Trim(CboLocationCD) & "' and stock_year='" & Format(DMonth, "yyyy") & "' and stock_month='" & DMonth.Month & "'"
                sql = sql & " order by  warehouse_code,sm.item_Code"
                
                
                  If rsRpt.State <> adStateClosed Then rsRpt.Close
                  rsRpt.Open sql, Db, adOpenDynamic, adLockOptimistic

                  If rsRpt.EOF Then LblErrMsg.Caption = DisplayMsg(4006): Me.MousePointer = vbDefault: Exit Sub

                  Set report = application.OpenReport(App.path & "\Reports\rpt_pi_report2.rpt")
                  report.Database.Tables(1).SetDataSource rsRpt
                ''#####################################################################
                ''# Qty Digit and decimal
                report.FormulaFields(5).Text = "" & gi_decimalDigitQty & ""
                report.FormulaFields(6).Text = "" & gi_decimalDigitQty & ""
                ''#####################################################################
                
            End If
            
            reportcode = "pireport"
            printorient = "2"
            sqlprint = sql
            datePiList = Format(DMonth.Value, "MMM yyyy")
            dtMPList = DMonth.Value
             dates = Format(DMonth.Value, "MMM yyyy")
             report.FormulaFields(1).Text = "'" & dates & "'"
             report.ReportTitle = "Inventory Report"
            
              Rpt.CRViewer1.ReportSource = report
              Rpt.CRViewer1.ViewReport
              Rpt.CRViewer1.Zoom 1
            
              Rpt.WindowState = 2
              Rpt.Show 1
            
              Me.MousePointer = vbDefault
                
End Select
        
        

End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    LblErrMsg.Caption = ErrMsg
End If
End Sub

Private Sub DMonth_Change()
Call clearGrid
If Format(DMonth.Value, "MM") < Format(dateUp, "MM") And Val(Format(DMonth.Value, "MM")) = 1 And Val(Format(dateUp, "MM")) = 12 Then _
            DMonth.Year = DMonth.Year + 1: GoTo pass
    If Format(DMonth.Value, "MM") > Format(dateUp, "MM") And Val(Format(DMonth.Value, "MM")) = 12 And Val(Format(dateUp, "MM")) = 1 Then _
            DMonth.Year = DMonth.Year - 1
pass:
    dateUp = Format(DMonth.Value, "dd MMM yyyy")
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
LblLocationName = ""
LblErrMsg = ""

DMonth = Format(Date, "MMM yyyy")
dateUp = DMonth.Value

CtrlMenu1.FormName = Me.Name
Me.Caption = "Inventory Report"
Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"

Call StockLocation
DMonth = Format(Now, "mmmm yyyy")
Call Header

End Sub

Private Sub StockLocation()
Dim sql As String, ls_sql As String, RsStock As New ADODB.Recordset
Dim i As Integer

If RsStock.State <> adStateClosed Then RsStock.Close

ls_sql = " select * from (select wh_code, wh_name  from warehouse_master where stockcontrol_cls='01' union  " & _
      " select trade_code wh_code, trade_name wh_name from trade_master where trade_code in(select manufacture_code from manufacture_line))tbWarehouse order by wh_code "

RsStock.Open ls_sql, Db, adOpenKeyset, adLockOptimistic
CboLocationCD.columnCount = 2
CboLocationCD.clear

i = 0
Do While Not RsStock.EOF
   CboLocationCD.AddItem ""
   CboLocationCD.List(i, 0) = Trim(RsStock("wh_code"))
   CboLocationCD.List(i, 1) = Trim(RsStock("wh_name"))
   i = i + 1
   RsStock.MoveNext
Loop

CboLocationCD.ColumnWidths = "50 pt; 150 pt"
CboLocationCD.ListWidth = 200
CboLocationCD.ListRows = 15
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With grid
.TextMatrix(Row, Col) = Format(.TextMatrix(Row, Col), gs_formatQty)
End With
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'If Grid.Col <> bteColInventory Then
   Cancel = True
'End If
End Sub

Private Sub Browse()

Dim RsStock As New ADODB.Recordset
Dim sqlControl As String
Dim RsInvControl As New ADODB.Recordset
Dim StokAwal As String
Dim closingdate As Date
Dim SelisihClosing As Double


If RsStock.State <> adStateClosed Then RsStock.Close

SelisihClosing = up_GetDateRange(DMonth)

If SelisihClosing = 0 Then
    StokAwal = "lm_premonth"
    closingdate = DMonth
ElseIf SelisihClosing = 1 Then
    StokAwal = "tm_premonth"
    closingdate = DateSerial(Year(DMonth), Month(DMonth) - 1, 1)
ElseIf SelisihClosing = 2 Then
    StokAwal = "nm_premonth"
    closingdate = DateSerial(Year(DMonth), Month(DMonth) - 2, 1)
End If

If SelisihClosing = 0 Or SelisihClosing = 1 Or SelisihClosing = 2 Then
    sql = " select gc.description groupDesc, uc.description Unit_Desc, sm.*,item_name,address, " & vbCrLf & _
                "  description=(case isnull(im.sheetcoil_cls,'') when '' then rtrim(im.item_name)  else rtrim(cast(im.item_name as varchar(15))) + rtrim(cast(im.item_name as varchar(50)))+ rtrim(cast(',' as varchar(1))) + rtrim(cast('T' as varchar(1))) + rtrim(cast(IM.Thickness as varchar(18)))  + rtrim(cast('x' as varchar(1))) + rtrim(cast('W' as varchar(1))) + rtrim(cast(IM.width as varchar(18))) +  rtrim(cast(im.length as varchar(18))) end)   " & vbCrLf & _
                " ,wh_name, makeritem_code,  " & vbCrLf & _
                " PO = 0, --isnull((select sum(qty) from purchaseorder_detail PD inner join PurchaseOrder_master PM On PM.PO_no=PD.PO_NO where PM.Fix_cls='1' and month(PM.Delivery_date)='" & Format(DMonth.Month, "00") & "' and year(PM.Delivery_Date)='" & DMonth.Year & " ' and item_code=Sm.Item_Code),0), " & vbCrLf & _
                " StokBegin = 0, --isnull((select sum(" & StokAwal & ") from stock_master where item_code=sm.item_code),0), " & vbCrLf & _
                " Incoming = 0, --isnull((select sum(Qty) from part_receipt where receipt_cls ='R' and Item_code=sm.item_code and month(receipt_date)='" & Format(DMonth.Month, "00") & "' and year(receipt_date)='" & DMonth.Year & "'),0), " & vbCrLf & _
                " Req = 0 --isnull((select case when sum(childRequirement_qty)-sum(childRequirementResult_qty)-sum(offchildrequirement_qty)<0 then 0 else sum(childRequirement_qty) end as SisaReqQty  " & vbCrLf & _
                " --from requirement  " & vbCrLf & _
                " --where year(childrequirement_date) ='" & Format(DMonth, "yyyy") & "'  " & vbCrLf & _
                "       --and month(childrequirement_date) = '" & Format(DMonth.Month, "00") & "' " & vbCrLf & _
                "       --and (complete_cls is null or complete_cls<>'1')  " & vbCrLf & _
                "       --and childItem_code=sm.item_code),0) " & vbCrLf & _
                " from stock_master sm  left join warehouse_master wm  " & vbCrLf & _
                " on sm.warehouse_code=wm.wh_code  left join item_master im on im.item_code=sm.item_code   " & vbCrLf & _
                " left join group_cls gc on im.group_cls=gc.group_cls   " & vbCrLf & _
                " left join unit_cls uc on im.unit_cls=uc.unit_cls  " & vbCrLf & _
                " left join  sheetcoil_cls sh on im.sheetcoil_cls=sh.sheetcoil_cls  " & vbCrLf & _
                " where warehouse_code='" & Trim(CboLocationCD) & "'"
    sql = sql & " order by  warehouse_code,sm.item_Code"
Else
    sql = " select gc.description groupDesc, uc.description Unit_Desc, sm.*,item_name,address, " & vbCrLf & _
                "  description=(case isnull(im.sheetcoil_cls,'') when '' then rtrim(im.item_name)  else rtrim(cast(im.item_name as varchar(15))) + rtrim(cast(im.item_name as varchar(50)))+ rtrim(cast(',' as varchar(1))) + rtrim(cast('T' as varchar(1))) + rtrim(cast(IM.Thickness as varchar(18)))  + rtrim(cast('x' as varchar(1))) + rtrim(cast('W' as varchar(1))) + rtrim(cast(IM.width as varchar(18))) +  rtrim(cast(im.length as varchar(18))) end)   " & vbCrLf & _
                " ,wh_name, makeritem_code,  " & vbCrLf & _
                " PO =  0, --isnull((select sum(qty) from purchaseorder_detail PD inner join PurchaseOrder_master PM On PM.PO_no=PD.PO_NO where PM.Fix_cls='1' and month(PM.Delivery_date)='" & Format(DMonth.Month, "00") & "' and year(PM.Delivery_Date)='" & DMonth.Year & " ' and item_code=Sm.Item_Code),0), " & vbCrLf & _
                " StokBegin = 0, --isnull((select sum(Premonth) from stock_history where item_code=sm.item_code and stock_month='" & DMonth.Month & "' and stock_year ='" & DMonth.Year & "' ),0), " & vbCrLf & _
                " Incoming = 0, --isnull((select sum(Qty) from part_receipt where receipt_cls ='R' and Item_code=sm.item_code and month(receipt_date)='" & Format(DMonth.Month, "00") & "' and year(receipt_date)='" & DMonth.Year & "'),0), " & vbCrLf & _
                " Req = 0 --isnull((select case when sum(childRequirement_qty)<0 then 0 else sum(childRequirement_qty) end as SisaReqQty  " & vbCrLf & _
                " --from requirement  " & vbCrLf & _
                " --where year(childrequirement_date) ='" & Format(DMonth, "yyyy") & "'  " & vbCrLf & _
                "       --and month(childrequirement_date) = '" & Format(DMonth.Month, "00") & "' " & vbCrLf & _
                "       --and (complete_cls is null or complete_cls<>'1')  " & vbCrLf & _
                "       --and childItem_code=sm.item_code),0) " & vbCrLf & _
                " from stock_history sm  left join warehouse_master wm  " & vbCrLf & _
                " on sm.warehouse_code=wm.wh_code  left join item_master im on im.item_code=sm.item_code   " & vbCrLf & _
                " left join group_cls gc on im.group_cls=gc.group_cls   " & vbCrLf & _
                " left join unit_cls uc on im.unit_cls=uc.unit_cls  " & vbCrLf & _
                " left join  sheetcoil_cls sh on im.sheetcoil_cls=sh.sheetcoil_cls  " & vbCrLf & _
                " where warehouse_code='" & Trim(CboLocationCD) & "' and stock_year='" & Format(DMonth, "yyyy") & "' and stock_month='" & DMonth.Month & "'"
    sql = sql & " order by  warehouse_code,sm.item_Code"
End If

RsStock.Open sql, Db, adOpenDynamic, adLockOptimistic

sqlControl = "select * from inventory_control where fix_cls='1' order by inventory_year,inventory_month"

If RsInvControl.State <> adStateClosed Then RsInvControl.Close
RsInvControl.Open sqlControl, Db, adOpenKeyset, adLockOptimistic

If RsInvControl.EOF = True And RsInvControl.BOF = True Then
    LblErrMsg = DisplayMsg(4022) '"Inventory Stock hasn't been closed !"
    Exit Sub
End If

RsInvControl.MoveLast


'LblErrMsg = up_ValidateDateRange(DMonth.Value, False)
'If Trim(LblErrMsg) <> "" Then Exit Sub

Select Case up_GetDateRange(DMonth)
    Case 0:
            
            i = 0
            Do While Not RsStock.EOF
                i = i + 1
                grid.AddItem i
                grid.TextMatrix(i, bteColProdCod) = Trim(RsStock!Item_Code)
                grid.TextMatrix(i, bteColPartNo) = Trim(RsStock!MakerItem_Code) & ""
                'Grid.TextMatrix(i, bteColDesc) = uf_GetItemDescription(Trim(RsStock!item_code))
                grid.TextMatrix(i, bteColDesc) = Trim(RsStock!Description) & ""
                grid.TextMatrix(i, bteColGroupDesc) = Trim(RsStock!groupdesc) & ""
                grid.TextMatrix(i, bteColUnit) = Trim(RsStock!Unit_Desc) & ""
                grid.TextMatrix(i, bteColAddress) = Trim(RsStock!Address) & ""
                grid.TextMatrix(i, bteColPreMonth) = Format(RsStock!lm_premonth, gs_formatQty)
                grid.TextMatrix(i, bteColReceipt) = Format(RsStock!lm_receipt, gs_formatQty)
                grid.TextMatrix(i, bteColSupply) = Format(RsStock!lm_supply, gs_formatQty)
                grid.TextMatrix(i, bteColLossReject) = Format(RsStock!lm_lossreject, gs_formatQty)
                grid.TextMatrix(i, bteColEnd) = Format(RsStock!lm_current, gs_formatQty)
                grid.TextMatrix(i, bteColInventory) = IIf(RsStock!lm_inventory = Null, "", Format(RsStock!lm_inventory, gs_formatQty))
                grid.TextMatrix(i, bteColReason) = IIf(IsNull(RsStock!lm_reason), "", Trim(RsStock!lm_reason))
                'Grid.TextMatrix(i, bteColStockStatus) = IIf(IsNull(RsStock!Stock_status), "", Format(Trim(RsStock!Stock_status), gs_formatQty))
                grid.TextMatrix(i, bteColPO) = IIf(IsNull(RsStock!po), "", Format(Trim(RsStock!po), gs_formatQty))
                grid.TextMatrix(i, bteColBeginTotal) = IIf(IsNull(RsStock!StokBegin), "", Format(Trim(RsStock!StokBegin), gs_formatQty))
                grid.TextMatrix(i, bteColIncoming) = IIf(IsNull(RsStock!Incoming), "", Format(Trim(RsStock!Incoming), gs_formatQty))
                grid.TextMatrix(i, bteColReq) = IIf(IsNull(RsStock!req), "", Format(Trim(RsStock!req), gs_formatQty))
                grid.TextMatrix(i, bteColGrandTotal) = Format(CDbl(RsStock!StokBegin) + CDbl(RsStock!Incoming) - CDbl(RsStock!req), gs_formatQty)
                grid.TextMatrix(i, bteColGrandTotal2) = Format(CDbl(RsStock!StokBegin) + CDbl(RsStock!po) - CDbl(RsStock!req), gs_formatQty)
                RsStock.MoveNext
                
            Loop
                
    Case 1:

            i = 0
            Do While Not RsStock.EOF
                i = i + 1
                grid.AddItem i
                grid.TextMatrix(i, bteColProdCod) = Trim(RsStock!Item_Code)
                grid.TextMatrix(i, bteColPartNo) = Trim(RsStock!MakerItem_Code) & ""
                'Grid.TextMatrix(i, bteColDesc) = uf_GetItemDescription(Trim(RsStock!item_code))
                grid.TextMatrix(i, bteColDesc) = Trim(RsStock!Description) & ""
                grid.TextMatrix(i, bteColAddress) = Trim(RsStock!Address) & ""
                grid.TextMatrix(i, bteColGroupDesc) = Trim(RsStock!groupdesc) & ""
                grid.TextMatrix(i, bteColUnit) = Trim(RsStock!Unit_Desc) & ""
                grid.TextMatrix(i, bteColPreMonth) = Format(RsStock!tm_premonth, gs_formatQty)
                grid.TextMatrix(i, bteColReceipt) = Format(RsStock!tm_receipt, gs_formatQty)
                grid.TextMatrix(i, bteColSupply) = Format(RsStock!tm_supply, gs_formatQty)
                grid.TextMatrix(i, bteColLossReject) = Format(RsStock!tm_lossreject, gs_formatQty)
                grid.TextMatrix(i, bteColEnd) = Format(RsStock!tm_current, gs_formatQty)
                grid.TextMatrix(i, bteColInventory) = Format(RsStock!tm_inventory, gs_formatQty)
                 grid.TextMatrix(i, bteColInventory) = IIf(RsStock!tm_inventory = Null, "", Format(RsStock!tm_inventory, gs_formatQty))
                 grid.TextMatrix(i, bteColReason) = IIf(IsNull(RsStock!tm_reason), "", Trim(RsStock!tm_reason))
                'Grid.TextMatrix(i, bteColStockStatus) = IIf(IsNull(RsStock!Stock_status), "", Trim(RsStock!Stock_status))
                grid.TextMatrix(i, bteColPO) = IIf(IsNull(RsStock!po), "", Format(Trim(RsStock!po), gs_formatQty))
                grid.TextMatrix(i, bteColBeginTotal) = IIf(IsNull(RsStock!StokBegin), "", Format(Trim(RsStock!StokBegin), gs_formatQty))
                grid.TextMatrix(i, bteColIncoming) = IIf(IsNull(RsStock!Incoming), "", Format(Trim(RsStock!Incoming), gs_formatQty))
                grid.TextMatrix(i, bteColReq) = IIf(IsNull(RsStock!req), "", Format(Trim(RsStock!req), gs_formatQty))
                grid.TextMatrix(i, bteColGrandTotal) = Format(CDbl(RsStock!StokBegin) + CDbl(RsStock!Incoming) - CDbl(RsStock!req), gs_formatQty)
                grid.TextMatrix(i, bteColGrandTotal2) = Format(CDbl(RsStock!StokBegin) + CDbl(RsStock!po) - CDbl(RsStock!req), gs_formatQty)
                RsStock.MoveNext
                
            Loop
                        
                        
    Case 2:

            i = 0
            Do While Not RsStock.EOF
                i = i + 1
                grid.AddItem i
                grid.TextMatrix(i, bteColProdCod) = Trim(RsStock!Item_Code)
                grid.TextMatrix(i, bteColPartNo) = Trim(RsStock!MakerItem_Code) & ""
                'Grid.TextMatrix(i, bteColDesc) = uf_GetItemDescription(Trim(RsStock!item_code))
                grid.TextMatrix(i, bteColDesc) = Trim(RsStock!Description) & ""
                grid.TextMatrix(i, bteColAddress) = Trim(RsStock!Address) & ""
                grid.TextMatrix(i, bteColGroupDesc) = Trim(RsStock!groupdesc) & ""
                grid.TextMatrix(i, bteColUnit) = Trim(RsStock!Unit_Desc) & ""
                grid.TextMatrix(i, bteColPreMonth) = Format(RsStock!nm_premonth, gs_formatQty)
                grid.TextMatrix(i, bteColReceipt) = Format(RsStock!nm_receipt, gs_formatQty)
                grid.TextMatrix(i, bteColSupply) = Format(RsStock!nm_supply, gs_formatQty)
                grid.TextMatrix(i, bteColLossReject) = Format(RsStock!nm_lossreject, gs_formatQty)
                grid.TextMatrix(i, bteColEnd) = Format(RsStock!nm_current, gs_formatQty)
                grid.TextMatrix(i, bteColInventory) = Format(RsStock!nm_inventory, gs_formatQty)
                grid.TextMatrix(i, bteColInventory) = IIf(RsStock!nm_inventory = Null, "", Format(RsStock!nm_inventory, gs_formatQty))
                grid.TextMatrix(i, bteColReason) = IIf(IsNull(RsStock!nm_reason), "", Trim(RsStock!nm_reason))
                'Grid.TextMatrix(i, bteColStockStatus) = IIf(IsNull(RsStock!Stock_status), "", Trim(RsStock!Stock_status))
                grid.TextMatrix(i, bteColPO) = IIf(IsNull(RsStock!po), "", Format(Trim(RsStock!po), gs_formatQty))
                grid.TextMatrix(i, bteColBeginTotal) = IIf(IsNull(RsStock!StokBegin), "", Format(Trim(RsStock!StokBegin), gs_formatQty))
                grid.TextMatrix(i, bteColIncoming) = IIf(IsNull(RsStock!Incoming), "", Format(Trim(RsStock!Incoming), gs_formatQty))
                grid.TextMatrix(i, bteColReq) = IIf(IsNull(RsStock!req), "", Format(Trim(RsStock!req), gs_formatQty))
                grid.TextMatrix(i, bteColGrandTotal) = Format(CDbl(RsStock!StokBegin) + CDbl(RsStock!Incoming) - CDbl(RsStock!req), gs_formatQty)
                grid.TextMatrix(i, bteColGrandTotal2) = Format(CDbl(RsStock!StokBegin) + CDbl(RsStock!po) - CDbl(RsStock!req), gs_formatQty)
                RsStock.MoveNext
                
            Loop
 Case Else
            i = 0
            Do While Not RsStock.EOF
                i = i + 1
                grid.AddItem i
                grid.TextMatrix(i, bteColProdCod) = Trim(RsStock!Item_Code)
                grid.TextMatrix(i, bteColPartNo) = Trim(RsStock!MakerItem_Code) & ""
                grid.TextMatrix(i, bteColDesc) = Trim(RsStock!Description) & ""
                grid.TextMatrix(i, bteColAddress) = Trim(RsStock!Address) & ""
                grid.TextMatrix(i, bteColGroupDesc) = Trim(RsStock!groupdesc) & ""
                grid.TextMatrix(i, bteColUnit) = Trim(RsStock!Unit_Desc) & ""
                grid.TextMatrix(i, bteColPreMonth) = Format(RsStock!Premonth, gs_formatQty)
                grid.TextMatrix(i, bteColReceipt) = Format(RsStock!Receipt, gs_formatQty)
                grid.TextMatrix(i, bteColSupply) = Format(RsStock!supply, gs_formatQty)
                grid.TextMatrix(i, bteColLossReject) = Format(RsStock!LossReject, gs_formatQty)
                grid.TextMatrix(i, bteColEnd) = Format(RsStock!current, gs_formatQty)
                 grid.TextMatrix(i, bteColInventory) = IIf(RsStock!inventory = Null, "", Format(RsStock!inventory, gs_formatQty))
                 grid.TextMatrix(i, bteColReason) = IIf(IsNull(RsStock!reason) Or Trim(RsStock!reason) = "Null", "", Trim(RsStock!reason))
                'Grid.TextMatrix(i, bteColStockStatus) = IIf(IsNull(RsStock!Stock_status), "", Trim(RsStock!Stock_status))
                grid.TextMatrix(i, bteColPO) = IIf(IsNull(RsStock!po), "", Format(Trim(RsStock!po), gs_formatQty))
                grid.TextMatrix(i, bteColBeginTotal) = IIf(IsNull(RsStock!StokBegin), "", Format(Trim(RsStock!StokBegin), gs_formatQty))
                grid.TextMatrix(i, bteColIncoming) = IIf(IsNull(RsStock!Incoming), "", Format(Trim(RsStock!Incoming), gs_formatQty))
                grid.TextMatrix(i, bteColReq) = IIf(IsNull(RsStock!req), "", Format(Trim(RsStock!req), gs_formatQty))
                grid.TextMatrix(i, bteColGrandTotal) = Format(CDbl(RsStock!StokBegin) + CDbl(RsStock!Incoming) - CDbl(RsStock!req), gs_formatQty)
                grid.TextMatrix(i, bteColGrandTotal2) = Format(CDbl(RsStock!StokBegin) + CDbl(RsStock!po) - CDbl(RsStock!req), gs_formatQty)
                RsStock.MoveNext
                
            Loop
 End Select


End Sub

Private Sub Grid_DblClick()
    If grid.Row = 1 Then
        If bytSort = 0 Then
            grid.Sort = flexSortGenericDescending
            bytSort = 1
        Else
            grid.Sort = flexSortGenericAscending
            bytSort = 0
        End If
    End If
End Sub

Private Sub grid_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) < 0 Or Chr(KeyAscii) > 9 And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Sub toExcel()
Dim xlapp As New Excel.application
Dim rsCek As New Recordset, Idx As Long, tempi As String, tempcust As String
Dim bolcust As Boolean, bolinv As Boolean
Dim rsCompany As New Recordset, sql_plus As String, sqlP As String
Dim sqlControl As String, RsInvControl As New ADODB.Recordset
Dim selisih As Double


Dim ls_prod As String
Dim ls_desc As String
Dim ls_unit As String
Dim ls_addr As String
Dim ls_premonth As String
Dim ls_receipt As String
Dim ls_supply As String
Dim ls_loss As String
Dim ls_end As String
Dim ls_inventory As String
Dim ls_diff As String
Dim ls_reason As String
Dim ls_stock As String
    
ls_prod = "A"
ls_desc = "B"
ls_unit = "C"
ls_addr = "D"
ls_premonth = "E"
ls_receipt = "F"
ls_supply = "G"
ls_loss = "H"
ls_end = "I"
ls_inventory = "J"
ls_diff = "K"
ls_reason = "L"
ls_stock = "M"
    
Me.MousePointer = vbHourglass

CboLocationCD = Trim(CboLocationCD)
If CboLocationCD.MatchFound Then
    LblLocationName = Trim(CboLocationCD.Column(1))
Else
    LblErrMsg = DisplayMsg(4018)
    Exit Sub
End If

sqlControl = "select * from inventory_control where fix_cls='1' order by inventory_year,inventory_month"

If RsInvControl.State <> adStateClosed Then RsInvControl.Close
RsInvControl.Open sqlControl, Db, adOpenKeyset, adLockOptimistic

If RsInvControl.EOF = True And RsInvControl.BOF = True Then LblErrMsg = "Inventory Stock hasn't been closed !": Exit Sub
RsInvControl.MoveLast

selisih = up_GetDateRange(DMonth)

If selisih = 0 Then
    sqlP = "lm_premonth PreMonth, LM_Receipt Receipt, LM_supply Supply, LM_lossReject LossReject, Lm_Current Currents, LM_Inventory Inventory, lm_reason reason--,Stock_Status "
ElseIf selisih = 1 Then
    sqlP = "TM_premonth PreMonth, TM_Receipt Receipt, TM_supply Supply, TM_lossReject LossReject, TM_Current Currents, TM_Inventory Inventory, tm_reason reason--,Stock_Status "
Else
    sqlP = "NM_premonth PreMonth, NM_Receipt Receipt, NM_supply Supply, NM_lossReject LossReject, NM_Current Currents, NM_Inventory Inventory, nm_reason reason--,Stock_Status "
End If

If selisih = 0 Or selisih = 1 Or selisih = 2 Then
    
    sql = "select gc.description groupDesc, uc.description Unit_Desc,sm.item_code, " & _
          vbLf & " descriptions=case isnull(im.sheetcoil_cls,0) when 0 then  im.item_name else rtrim(im.item_name) + ' (' + rtrim(sh.description) + ', T' + cast(im.thickness as varchar(15)) + ' x W' + cast(im.width as varchar(15)) + ' x L' + cast(im.length as varchar(15)) + ')'  end " & _
          vbLf & ",address, wh_name, makeritem_code, " & sqlP & _
          vbLf & " from stock_master sm left join warehouse_master wm on " & _
          vbLf & " sm.warehouse_code = wm.wh_code " & _
          vbLf & " left join item_master im on sm.item_code=im.item_code " & _
          vbLf & " left join sheetcoil_cls sh on im.sheetcoil_cls = sh.sheetcoil_cls " & _
          vbLf & " left join group_cls gc on im.group_cls=gc.group_cls " & _
          vbLf & " left join unit_cls uc on im.unit_cls=uc.unit_cls " & _
          vbLf & " where warehouse_code='" & Trim(CboLocationCD) & "' order by  warehouse_code,sm.item_Code"
          
Else

    sql = " select gc.description groupDesc, uc.description Unit_Desc, sm.*,item_name,address, " & vbCrLf & _
                "  description=(case isnull(im.sheetcoil_cls,'') when '' then rtrim(im.item_name)  else rtrim(cast(im.item_name as varchar(15))) + rtrim(cast(im.item_name as varchar(50)))+ rtrim(cast(',' as varchar(1))) + rtrim(cast('T' as varchar(1))) + rtrim(cast(IM.Thickness as varchar(18)))  + rtrim(cast('x' as varchar(1))) + rtrim(cast('W' as varchar(1))) + rtrim(cast(IM.width as varchar(18))) +  rtrim(cast(im.length as varchar(18))) end)   " & vbCrLf & _
                " ,wh_name, makeritem_code,  " & vbCrLf & _
                " PO = 0 ,--isnull((select sum(qty) from purchaseorder_detail PD inner join PurchaseOrder_master PM On PM.PO_no=PD.PO_NO where PM.Fix_cls='1' and month(PM.Delivery_date)='" & Format(DMonth.Month, "00") & "' and year(PM.Delivery_Date)='" & DMonth.Year & " ' and item_code=Sm.Item_Code),0), " & vbCrLf & _
                " StokBegin = 0, --isnull((select sum(Premonth) from stock_history where item_code=sm.item_code and stock_month='" & DMonth.Month & "' and stock_year ='" & DMonth.Year & "' ),0), " & vbCrLf & _
                " Incoming = 0, --isnull((select sum(Qty) from part_receipt where receipt_cls ='R' and Item_code=sm.item_code and month(receipt_date)='" & Format(DMonth.Month, "00") & "'),0), " & vbCrLf & _
                " Req = 0--, -- isnull((select case when sum(childRequirement_qty)-sum(childRequirementResult_qty)-sum(offchildrequirement_qty)<0 then 0 else sum(childRequirement_qty) end as SisaReqQty  " & vbCrLf & _
                " --from requirement  " & vbCrLf & _
                " --where year(childrequirement_date) ='" & Format(DMonth, "yyyy") & "'  " & vbCrLf & _
                "  --     and month(childrequirement_date) = '" & Format(DMonth.Month, "00") & "' " & vbCrLf & _
                "  --     and (complete_cls is null or complete_cls<>'1')  " & vbCrLf & _
                "  --     and childItem_code=sm.item_code),0) " & vbCrLf & _
                " from stock_history sm  left join warehouse_master wm  " & vbCrLf & _
                " on sm.warehouse_code=wm.wh_code  left join item_master im on im.item_code=sm.item_code   " & vbCrLf & _
                " left join group_cls gc on im.group_cls=gc.group_cls   " & vbCrLf & _
                " left join unit_cls uc on im.unit_cls=uc.unit_cls  " & vbCrLf & _
                " left join  sheetcoil_cls sh on im.sheetcoil_cls=sh.sheetcoil_cls  " & vbCrLf & _
                " where warehouse_code='" & Trim(CboLocationCD) & "' and stock_year='" & Format(DMonth, "yyyy") & "' and stock_month='" & DMonth.Month & "'"
    sql = sql & " order by  warehouse_code,sm.item_Code"
    
End If

If rsCek.State <> adStateClosed Then rsCek.Close
rsCek.CursorLocation = adUseClient
rsCek.Open sql, Db, adOpenDynamic, adLockOptimistic


If Not rsCek.EOF Then
Screen.MousePointer = vbHourglass
With xlapp

    sql = "select rtrim(company_name) company_name, rtrim(address1) Address1, rtrim(Address2) Address2, rtrim(Province) Province, rtrim(city) City, Rtrim(Postal_Code) POstal_Code, Rtrim(phone1) Phone1, Rtrim(phone2) Phone2,rtrim(fax) Fax  From company_profile "
    If rsCompany.State <> adStateClosed Then rsCompany.Close
    rsCompany.Open sql, Db, adOpenDynamic, adLockOptimistic
    If rsCompany.EOF Then Screen.MousePointer = vbDefault: Exit Sub
    .Workbooks.Add
    
    .Range("a2", ls_reason & "2").Merge
    .Range("a2") = rsCompany!company_name
    .Range("a3", ls_reason & "3").Merge
    .Range("a3") = rsCompany!address1 & " " & rsCompany!address2 & " " & rsCompany!City & " " & rsCompany!Province & " " & rsCompany!postal_code
    .Range("a4", ls_reason & "4").Merge
    .Range("a4") = "Phone: " & rsCompany!phone1 & " " & rsCompany!phone2 & " Fax: " & rsCompany!fax
    
    .Range("a6") = "Inventory Report"
    .Range("b6") = ""
    .Range("a6", "b6").Merge
    .Range("a6").horizontalAlignment = xlLeft
    .Range("a7") = "Warehouse Code"
    .Range("b7", ls_reason & "7").Merge
    .Range("b7") = ": " & Trim(CboLocationCD.Text) & " / " & Trim(LblLocationName)
    .Range("B7").horizontalAlignment = xlLeft
    .Range("a8") = "Date(Month)"
    .Range("b8") = ": " & Format(DMonth, "MMMM YYYY")
    
    Idx = 10
    
    .Range("A:A").NumberFormat = "@"
    
'    .Range(ls_prod & 10, ls_prod & rsCek.RecordCount).NumberFormat = "@"
    
    Do While Not rsCek.EOF
        If Idx = 10 Then
            .Range(ls_prod & Idx) = "Product Code"
            .Range(ls_desc & Idx) = "Description"
            .Range(ls_unit & Idx) = "Unit"
            .Range(ls_addr & Idx) = "Address"
            .Range(ls_premonth & Idx) = "Pre Month Stock"
            .Range(ls_receipt & Idx) = "Receipt Total"
            .Range(ls_supply & Idx) = "Supply Total"
            .Range(ls_loss & Idx) = "Loss/Reject"
            .Range(ls_end & Idx) = "End of Month Stock"
            .Range(ls_inventory & Idx) = "Inventory"
            .Range(ls_diff & Idx) = "Difference"
            .Range(ls_reason & Idx) = "Reason"
            '.Range(ls_stock & Idx) = "Stock Status"
            
            .Range(ls_prod & Idx, ls_stock & Idx).horizontalAlignment = xlCenter
            Idx = Idx + 1
        End If
        
        Idx = Idx
        'Content
        .Range(ls_prod & Idx) = Trim(rsCek!Item_Code)
        If selisih = 0 Or selisih = 1 Or selisih = 2 Then
            .Range(ls_desc & Idx) = Trim(rsCek!descriptions)
            .Range(ls_end & Idx) = Format(rsCek!Currents, gs_formatQty)
            .Range(ls_diff & Idx) = Format(rsCek!Currents - rsCek!inventory, gs_formatQty)
        Else
            .Range(ls_desc & Idx) = Trim(rsCek!item_name)
            .Range(ls_end & Idx) = Format(rsCek!current, gs_formatQty)
            .Range(ls_diff & Idx) = Format(rsCek!current - rsCek!inventory, gs_formatQty)
        End If
        .Range(ls_unit & Idx) = Trim(rsCek!Unit_Desc)
        .Range(ls_addr & Idx) = Trim(rsCek!Address)
        .Range(ls_premonth & Idx) = Format(rsCek!Premonth, gs_formatQty)
        .Range(ls_receipt & Idx) = Format(rsCek!Receipt, gs_formatQty)
        .Range(ls_supply & Idx) = Format(rsCek!supply, gs_formatQty)
        .Range(ls_loss & Idx) = Format(rsCek!LossReject, gs_formatQty)
        .Range(ls_inventory & Idx) = Format(rsCek!inventory, gs_formatQty)
        
        .Range(ls_reason & Idx) = IIf(IsNull(rsCek!reason), "", Trim(rsCek!reason))
        '.Range(ls_stock & Idx) = IIf(IsNull(rsCek!Stock_status), "", Trim(rsCek!Stock_status))
        Idx = Idx + 1
        rsCek.MoveNext
    Loop
        
    'Number Format
    
    .Range(ls_premonth & 10, ls_premonth & Idx - 1).NumberFormat = gs_formatQty
    .Range(ls_receipt & 10, ls_receipt & Idx - 1).NumberFormat = gs_formatQty
    .Range(ls_supply & 10, ls_supply & Idx - 1).NumberFormat = gs_formatQty
    .Range(ls_loss & 10, ls_loss & Idx - 1).NumberFormat = gs_formatQty
    .Range(ls_end & 10, ls_end & Idx - 1).NumberFormat = gs_formatQty
    .Range(ls_inventory & 10, ls_inventory & Idx - 1).NumberFormat = gs_formatQty
    .Range(ls_diff & 10, ls_diff & Idx - 1).NumberFormat = gs_formatQty
        
    'Border
    .Range(ls_prod & 10, ls_stock & Idx - 1).Borders(xlEdgeTop).LineStyle = xlContinuous
    .Range(ls_prod & 10, ls_stock & Idx - 1).Borders(xlEdgeBottom).LineStyle = xlContinuous
    .Range(ls_prod & 10, ls_stock & Idx - 1).Borders(xlEdgeLeft).LineStyle = xlContinuous
    .Range(ls_prod & 10, ls_stock & Idx - 1).Borders(xlEdgeRight).LineStyle = xlContinuous
    .Range(ls_prod & 10, ls_stock & Idx - 1).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    .Range(ls_prod & 10, ls_stock & Idx - 1).Borders(xlInsideVertical).LineStyle = xlContinuous
        
    'Alignment
    .Range(ls_reason & 11, ls_stock & Idx - 1).horizontalAlignment = xlLeft
                
    .Range("a1", ls_stock & Idx + 3).Columns.Font.Name = "Arial"
    .Range("a1", ls_stock & Idx + 3).Columns.Font.Size = 8
    
    .Range("a2", ls_stock & "2").Columns.Font.Name = "Arial"
    .Range("a2", ls_stock & "2").Columns.Font.Size = "10"
    .Range("a2", ls_stock & "2").Columns.Font.Bold = True
    .Range("a2", ls_stock & "4").horizontalAlignment = xlCenter
    .Range("a6", "b6").Columns.Font.Bold = True
    
    .ActiveSheet.PageSetup.Orientation = xlLandscape
    .Range("A:" & ls_stock).Columns.AutoFit
    .Range(ls_premonth & "11:" & ls_inventory & Idx).Select
    .Selection.NumberFormat = gs_formatQty
    .Range("A1").Select
    .WindowState = xlMaximized
    .Visible = True
End With
Else
    LblErrMsg = DisplayMsg(4006)
End If

Screen.MousePointer = vbDefault
Me.MousePointer = vbDefault
End Sub

Private Sub txtItemCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Cmd_Find_Click
End If
End Sub
