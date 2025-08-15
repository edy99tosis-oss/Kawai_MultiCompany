VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmUploadDetailItem 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Order Entry Upload"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13740
   Icon            =   "FrmUploadDetailItem.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8475
   ScaleWidth      =   13740
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtConsCode 
      Height          =   285
      Left            =   4560
      TabIndex        =   16
      Top             =   8160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtPoNo 
      Height          =   285
      Left            =   3360
      TabIndex        =   15
      Top             =   8160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtTradeCode 
      Height          =   285
      Left            =   2280
      TabIndex        =   14
      Top             =   8160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FDDFE3&
      Height          =   480
      Left            =   240
      TabIndex        =   11
      Top             =   7320
      Width           =   13275
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
         TabIndex        =   12
         Top             =   150
         Width           =   13020
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDDFE3&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   13245
      Begin VB.TextBox Text1 
         Height          =   360
         Left            =   1920
         TabIndex        =   8
         Top             =   240
         Width           =   3660
      End
      Begin VB.CommandButton cmdtemplate 
         BackColor       =   &H0080FFFF&
         Caption         =   "Template"
         Height          =   375
         Left            =   6480
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1140
      End
      Begin VB.CommandButton cmdLocation 
         BackColor       =   &H0080FFFF&
         Caption         =   "...."
         Height          =   375
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   660
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
         TabIndex        =   5
         Text            =   "0"
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Upload Location File"
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invalid [Item Code] or [Price]"
         Height          =   195
         Index           =   3
         Left            =   7800
         TabIndex        =   9
         Top             =   360
         Width           =   1995
      End
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
      Left            =   12375
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8040
      Width           =   1140
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8040
      Width           =   1125
   End
   Begin VSFlex8Ctl.VSFlexGrid grid 
      Height          =   5535
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   13260
      _cx             =   23389
      _cy             =   9763
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
         TabIndex        =   1
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
      Left            =   11400
      TabIndex        =   13
      Tag             =   "TTFF*/"
      Top             =   6960
      Width           =   2100
   End
End
Attribute VB_Name = "FrmUploadDetailItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
  
Dim sql As String, sqlGrid As String

Dim ls_PathExcel As String



Dim bteColItemCode As Byte
Dim bteColItemName As Byte
Dim bteColUnitCls As Byte
Dim bteColQty As Byte
Dim bteColCurr As Byte
Dim bteColPrice As Byte
Dim bteColAmount As Byte
Dim bteColDeliveryDate As Byte
Dim bteColRemark As Byte



Sub Header()
        
    Dim i As Long
    
  
    bteColItemCode = 0
    bteColItemName = 1
    bteColQty = 2
    bteColUnitCls = 3
    bteColCurr = 4
    bteColPrice = 5
    bteColAmount = 6
    bteColDeliveryDate = 7
    bteColRemark = 8
    
    With grid
        .clear
        .Rows = 1
        .ColS = 9
        
        
        .TextMatrix(0, bteColItemCode) = "Item Code"
        .TextMatrix(0, bteColItemName) = "Item Name"
        .TextMatrix(0, bteColQty) = "Qty"
        .TextMatrix(0, bteColUnitCls) = "Unit"
        .TextMatrix(0, bteColCurr) = "Curr"
        .TextMatrix(0, bteColPrice) = "Price"
        .TextMatrix(0, bteColAmount) = "Amount"
        .TextMatrix(0, bteColDeliveryDate) = "Delivery Date"
        .TextMatrix(0, bteColRemark) = "Remarks"
        
        .ColWidth(bteColItemCode) = 1250
        .ColWidth(bteColItemName) = 1800
        .ColWidth(bteColQty) = 1200
        .ColWidth(bteColCurr) = 800
        .ColWidth(bteColPrice) = 1500
        .ColWidth(bteColAmount) = 2000
        .ColWidth(bteColDeliveryDate) = 1300
        .ColWidth(bteColRemark) = 2000
   
        
        .EditMaxLength = 1
    End With
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
Dim TradeCode As String, PONO As String
Dim ItemCode As String, UnitCls As String, _
    Qty As Double, Curr As String, Price As Double, _
    Amount As Double, user As String, DeliveryDate As Date, Remarks As String

Me.MousePointer = vbHourglass


For X = 1 To grid.Rows - 1

TradeCode = Trim(txtTradeCode.Text)
PONO = Trim(txtPoNo.Text)
ItemCode = grid.TextMatrix(X, bteColItemCode)
UnitCls = grid.TextMatrix(X, bteColUnitCls)
Qty = grid.TextMatrix(X, bteColQty)
Curr = grid.TextMatrix(X, bteColCurr)
Price = grid.TextMatrix(X, bteColPrice)
Amount = grid.TextMatrix(X, bteColAmount)
user = userLogin
DeliveryDate = grid.TextMatrix(X, bteColDeliveryDate)
Remarks = grid.TextMatrix(X, bteColRemark)
        
        
If grid.Cell(flexcpBackColor, X, bteColItemCode, X, bteColRemark) <> &H8080FF Then
    
    
    grid.TextMatrix(X, bteColRemark) = uf_Save(TradeCode, PONO, ItemCode, UnitCls, Qty, Curr, Price, Amount, userLogin, DeliveryDate, Remarks)
    If grid.TextMatrix(X, bteColRemark) <> "" Then
        grid.Cell(flexcpBackColor, X, bteColItemCode, X, bteColRemark) = 16637923
    End If
    
End If
        
Next X
LblErrMsg.Caption = "Update Records Succes..!!"
Me.MousePointer = vbDefault
End Sub

Private Sub cmdtemplate_Click()
Call up_ExportOffline
End Sub

Private Sub Form_Load()
Header
End Sub

Private Sub up_ExportOffline()
Dim objExcel As New Excel.application
Dim RS As New ADODB.Recordset
Dim strSQL As String
Dim i As Double

'On Error GoTo errHandler
    
    If G_CekExcelApp = False Then LblErrMsg.Caption = "Excel Application is not found !": Exit Sub
    
    LblErrMsg.Caption = ""
    cdg.filter = "Excel Worksheets 2003 (*.xls)|*.xls|"
    cdg.filename = "Upload Order Entry Detail "
    cdg.CancelError = True
    
    On Error GoTo errCancel
    cdg.ShowSave
    
   On Error GoTo errHandler
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
    
        .Range("A1").Value = "Item Code"
        .Range("B1").Value = "Qty"
        .Range("C1").Value = "Delivery Date"
        .Range("D1").Value = "Remarks"

        
        .Range("A2").Value = "Char(15)"
        .Range("B2").Value = "Numeric(18,2)"
        .Range("C2").Value = "(yyyy-mm-dd)"
        .Range("D2").Value = "Char(35)"
        

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
        LblErrMsg.Caption = err.Description
        grid.FixedRows = 1
    End If
    If RS.State = adStateOpen Then
        RS.Close
        Set RS = Nothing
    End If
errCancel:

End Sub



Private Sub up_ImportOffline()

    Dim adoCmd As New Command
    Dim rsCheck As New ADODB.Recordset
    
    Dim objExcel As New Excel.application
    Dim objWorkSheet As New Worksheet
    Dim objWorkBook As Workbook
    Dim i As Long
    Dim iCol As Integer
    Dim colcount As Integer
    Dim RS As New ADODB.Recordset
    Dim strSQL As String
    Dim iGrdRow As Double
    Dim Year, Month, div, SubDiv, Block As String
    Dim HA As Double
    

    If G_CekExcelApp = False Then LblErrMsg.Caption = "Excel Application is not found": Exit Sub
    
    LblErrMsg.Caption = ""
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
                        grid.AddItem ""
                        
                        Dim ItemCode As String
                        Dim TradeCode As String
                        Dim ConsigneeCode As String
                        Dim Qty As Integer
                        Dim DeliveryDate As Date
                        Dim Remarks As String
                        
                        ItemCode = .Cells(i, 1)
                        TradeCode = Trim(txtTradeCode.Text)
                        ConsigneeCode = Trim(txtConsCode.Text)
                        Qty = .Cells(i, 2)
                        If IsDate(.Cells(i, 3)) Then
                            DeliveryDate = .Cells(i, 3)
                        End If
                        
                        Remarks = .Cells(i, 4)
                        
                        'Get Price
                        
                        adoCmd.ActiveConnection = Db.ConnectionString
                        adoCmd.CommandTimeout = 120
                        adoCmd.CommandType = adCmdStoredProc
                        adoCmd.CommandText = "SP_GetPriceOrderEntry"
                        adoCmd.Parameters(1) = Trim(ItemCode)
                        adoCmd.Parameters(2) = TradeCode
                        adoCmd.Parameters(3) = ConsigneeCode
                        adoCmd.Parameters(4) = Qty
                        
                                               
                        Set RS = adoCmd.Execute
                        If Not RS.EOF Then
                            grid.TextMatrix(iGrdRow, bteColItemCode) = RS("Item_Code")
                            grid.TextMatrix(iGrdRow, bteColItemName) = RS("Item_Name")
                            grid.TextMatrix(iGrdRow, bteColQty) = RS("Qty")
                            grid.TextMatrix(iGrdRow, bteColUnitCls) = RS("Unit_Cls")
                            
                            grid.TextMatrix(iGrdRow, bteColPrice) = RS("Price")
                            grid.TextMatrix(iGrdRow, bteColCurr) = RS("Curr")
                            grid.TextMatrix(iGrdRow, bteColAmount) = RS("Amount")
                            grid.TextMatrix(iGrdRow, bteColDeliveryDate) = Format(DeliveryDate, "yyyy-mm-dd")
                        
                            If CDbl(RS("Price")) = 0 Then
                                grid.TextMatrix(iGrdRow, bteColRemark) = "Invalid Price !"
                                grid.Cell(flexcpBackColor, iGrdRow, bteColItemCode, iGrdRow, bteColRemark) = &H8080FF
                                txtErr.Text = CDbl(txtErr.Text) + 1
                                
                            ElseIf Trim(RS("Price")) = "" Then
                                grid.TextMatrix(iGrdRow, bteColRemark) = "Invalid Item Code !"
                                grid.Cell(flexcpBackColor, iGrdRow, bteColItemCode, iGrdRow, bteColRemark) = &H8080FF
                                txtErr.Text = CDbl(txtErr.Text) + 1
                                
                            ElseIf DeliveryDate = "00:00:00" Then
                                grid.TextMatrix(iGrdRow, bteColRemark) = "Invalid Format Delivery Date (yyyy-mm-dd) !"
                                grid.Cell(flexcpBackColor, iGrdRow, bteColItemCode, iGrdRow, bteColRemark) = &H8080FF
                                txtErr.Text = CDbl(txtErr.Text) + 1
                                
                            Else
                            
                            
                                grid.TextMatrix(iGrdRow, bteColRemark) = Remarks
                            End If
                            
                        
                        End If
                        
                        
                        
                        iGrdRow = iGrdRow + 1
               
                    
                    
                LblErrMsg = "Reading row : " & i - 1
                DoEvents
                i = i + 1
            Loop
            
        End With
        
        ' clean object excel
        objExcel.Workbooks.Close
        Set objWorkSheet = Nothing
        
        Set objExcel = Nothing

        LblTotalRec = "Total : " & grid.Rows - 1 & " record (s)"
        LblErrMsg.Caption = "Reading Excel finish"
        
        Me.MousePointer = vbDefault
    End If
    Exit Sub
errCancel:
err:
    LblErrMsg.Caption = err.Description
    objExcel.Workbooks.Close
    Set objWorkBook = Nothing
    Set objWorkSheet = Nothing
    Set objWorkBook = Nothing
    Set objExcel = Nothing
    Me.MousePointer = vbDefault
    
End Sub

Private Function uf_Save(TradeCode As String, PONO As String, ItemCode As String, UnitCls As String, _
                      Qty As Double, Curr As String, Price As Double, Amount As Double, user As String, DeliveryDate As Date, Remarks As String) As String
        
       On Error GoTo err
        
        Dim adoCmd As New Command


        adoCmd.ActiveConnection = Db.ConnectionString
        adoCmd.CommandTimeout = 120
        adoCmd.CommandType = adCmdStoredProc
        adoCmd.CommandText = "SP_Save_OrderEntryDetail_Upload"
        adoCmd.Parameters(1) = TradeCode
        adoCmd.Parameters(2) = PONO
        adoCmd.Parameters(3) = ItemCode
        adoCmd.Parameters(4) = UnitCls
        adoCmd.Parameters(5) = Qty
        adoCmd.Parameters(6) = Curr
        adoCmd.Parameters(7) = Price
        adoCmd.Parameters(8) = Amount
        adoCmd.Parameters(9) = user
        adoCmd.Parameters(10) = DeliveryDate
        adoCmd.Parameters(11) = Remarks
        
        adoCmd.Execute
        uf_Save = ""
        Exit Function
                        
err:
      uf_Save = err.Description
      'lblErrMsg.Caption = err.Description
      Set adoCmd = Nothing
      Me.MousePointer = vbDefault
        


End Function

