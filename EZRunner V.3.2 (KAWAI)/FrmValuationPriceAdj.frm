VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmValuationPriceAdj 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Valuation Price Adjust"
   ClientHeight    =   10305
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14940
   Icon            =   "FrmValuationPriceAdj.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10305
   ScaleWidth      =   14940
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   600
      Left            =   180
      TabIndex        =   17
      Tag             =   "TFTT*/"
      Top             =   8610
      Width           =   14475
      Begin VB.Label LblErrMsg 
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
         Height          =   240
         Left            =   135
         TabIndex        =   18
         Tag             =   "TFTF*/"
         Top             =   210
         Width           =   14205
      End
   End
   Begin VB.CommandButton cmdSubMenu 
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   16
      Tag             =   "TFFT*/"
      Top             =   9510
      Width           =   1140
   End
   Begin VB.CommandButton cmdSubmit 
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
      Left            =   13515
      Style           =   1  'Graphical
      TabIndex        =   15
      Tag             =   "FFTT*/"
      Top             =   9510
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1695
      Left            =   180
      TabIndex        =   2
      Tag             =   "TTTF*/"
      Top             =   1260
      Width           =   14475
      Begin VB.CommandButton cmdSearch 
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
         Left            =   13050
         Style           =   1  'Graphical
         TabIndex        =   11
         Tag             =   "TTFF*/"
         Top             =   1080
         Width           =   1185
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   300
         Left            =   4320
         TabIndex        =   3
         Tag             =   "TTFF*/"
         Top             =   1005
         Width           =   300
      End
      Begin MSComCtl2.DTPicker DtPeriod 
         Height          =   330
         Left            =   2160
         TabIndex        =   4
         Tag             =   "TTFF*/"
         Top             =   480
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
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
         Format          =   137494531
         UpDown          =   -1  'True
         CurrentDate     =   37860
      End
      Begin VB.Label lblNm 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
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
         Left            =   8010
         TabIndex        =   10
         Tag             =   "TTFF*/"
         Top             =   1050
         Width           =   4410
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Code"
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
         Left            =   270
         TabIndex        =   9
         Tag             =   "TTFF*/"
         Top             =   1050
         Width           =   1155
      End
      Begin VB.Line Line8 
         Index           =   0
         X1              =   8010
         X2              =   12480
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Month Period"
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
         Index           =   1
         Left            =   270
         TabIndex        =   8
         Tag             =   "TTFF*/"
         Top             =   510
         Width           =   1350
      End
      Begin VB.Label Lbl_Make 
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
         Height          =   225
         Left            =   10200
         TabIndex        =   7
         Tag             =   "TTFF*/"
         Top             =   690
         Width           =   1845
      End
      Begin VB.Line Line8 
         Index           =   2
         X1              =   4740
         X2              =   7860
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label lblPartNumber 
         BackStyle       =   0  'Transparent
         Caption         =   "Part Number"
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
         Left            =   4740
         TabIndex        =   6
         Tag             =   "TTFF*/"
         Top             =   1050
         Width           =   3060
      End
      Begin MSForms.TextBox TxtProduk 
         Height          =   315
         Left            =   2160
         TabIndex        =   5
         Tag             =   "TTFF*/"
         Top             =   990
         Width           =   2055
         VariousPropertyBits=   746604571
         MaxLength       =   15
         Size            =   "3625;556"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   12810
      TabIndex        =   0
      Tag             =   "FTTF*/"
      Top             =   480
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   714
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   5160
      Left            =   180
      TabIndex        =   12
      Tag             =   "TTTT*/"
      Top             =   3060
      Width           =   14475
      _cx             =   25532
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
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   10
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
   Begin VB.Label LblRecord 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Record(s)"
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
      Left            =   12570
      TabIndex        =   14
      Tag             =   "FFTT*/"
      Top             =   8310
      Width           =   1065
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Record(s)"
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
      Left            =   13590
      TabIndex        =   13
      Tag             =   "FFTT*/"
      Top             =   8310
      Width           =   1065
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valuation Price Adjust"
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
      Left            =   6165
      TabIndex        =   1
      Tag             =   "TTTF*/"
      Top             =   465
      Width           =   2520
   End
End
Attribute VB_Name = "FrmValuationPriceAdj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dateUp As Date
Dim ColItemCode As Byte, ColDate As Byte, ColEndingBalance As Byte, ColEndingBalance_Hide As Byte
Dim ColPrice As Byte, ColPrice_Hide As Byte, ColAmount As Byte, ColValInventory As Byte

Dim SqlData As String

Private Sub SettingColumn()
    ColItemCode = 0
    ColDate = 1
    ColEndingBalance = 2
    ColPrice = 3
    ColAmount = 4
    ColValInventory = 5
    ColEndingBalance_Hide = 6
    ColPrice_Hide = 7
End Sub

Private Sub Header()
LblRecord = "0"
Dim RS As New ADODB.Recordset
Dim i As Long

    With grid
        .Rows = 1
        .ColS = 8
                        
        .TextMatrix(0, ColItemCode) = "Item Code"
        .TextMatrix(0, ColDate) = "Date"
        .TextMatrix(0, ColEndingBalance) = "Ending Balance"
        .TextMatrix(0, ColPrice) = "Price"
        .TextMatrix(0, ColAmount) = "Amount"
        .TextMatrix(0, ColValInventory) = "Value of Amount"
        .TextMatrix(0, ColEndingBalance_Hide) = "Ending Balance Hide"
        .TextMatrix(0, ColPrice_Hide) = "Price Hide"
        
        .ColWidth(ColItemCode) = 2300
        .ColWidth(ColDate) = 2000
        .ColWidth(ColEndingBalance) = 2000
        .ColWidth(ColPrice) = 2000
        .ColWidth(ColAmount) = 2500
        .ColWidth(ColValInventory) = 2500
        .ColWidth(ColEndingBalance_Hide) = 2000
        .ColWidth(ColPrice_Hide) = 2000
        
        .ColHidden(ColEndingBalance_Hide) = True
        .ColHidden(ColPrice_Hide) = True
        
        .Cell(flexcpAlignment, 0, ColItemCode, 0, ColValInventory) = flexAlignCenterCenter
    
    End With

End Sub

Private Sub cmdSearch_Click()
    Dim RS As New ADODB.Recordset
    Dim i As Long
    
    LblErrMsg = ""
    LblRecord = Format(0, "#,##0")
    
    Screen.MousePointer = vbHourglass
    
    SqlData = "SELECT Item_Code,Tgl,QtyEnding,AvgPrice,(QtyEnding * AvgPrice) Amount,(QtyEnding * AvgPrice) ValueInv " & vbCrLf & _
              "FROM dbo.Inventory_PriceDetail " & vbCrLf & _
              "WHERE Cls = 'Inv' AND Period = '" & Format(DtPeriod.Value, "yyyyMM") & "'" & vbCrLf
              
    If TxtProduk.Text <> "" Then
        SqlData = SqlData + "AND Item_Code = '" & Trim(TxtProduk.Text) & "'" & vbCrLf
    Else
        SqlData = SqlData + "AND QtyEnding < 0 " & vbCrLf
    End If
    
    SqlData = SqlData + "ORDER BY Item_Code,Tgl"
    
    If RS.State <> adStateClosed Then RS.Close
    RS.CursorLocation = adUseClient
    RS.Open SqlData, Db, adOpenKeyset, adLockOptimistic
    
    Header
    
    i = 0
    
    If RS.EOF <> True Then
        
        While RS.EOF = False
            i = i + 1
            grid.AddItem ""
            
            grid.TextMatrix(i, ColItemCode) = Trim(RS!Item_Code)
            grid.TextMatrix(i, ColDate) = Format(RS!Tgl, "dd MMM yyyy")
            grid.TextMatrix(i, ColEndingBalance) = Trim(RS!QtyEnding)
            grid.TextMatrix(i, ColPrice) = Trim(RS!AvgPrice)
            grid.TextMatrix(i, ColAmount) = Trim(RS!Amount)
            grid.TextMatrix(i, ColValInventory) = Trim(RS!ValueInv)
            
            grid.TextMatrix(i, ColEndingBalance_Hide) = Trim(RS!QtyEnding)
            grid.TextMatrix(i, ColPrice_Hide) = Trim(RS!AvgPrice)
            
            grid.Cell(flexcpAlignment, i, ColItemCode) = flexAlignLeftCenter
            grid.Cell(flexcpAlignment, i, ColDate) = flexAlignRightCenter
            grid.TextMatrix(i, ColEndingBalance) = Format(RS!QtyEnding, gs_formatQty)
            grid.TextMatrix(i, ColPrice) = Format(RS!AvgPrice, gs_formatAmountIDR)
            grid.TextMatrix(i, ColAmount) = Format(RS!Amount, gs_formatAmountIDR)
            grid.TextMatrix(i, ColValInventory) = Format(RS!ValueInv, gs_formatAmountIDR)
            
            grid.TextMatrix(i, ColEndingBalance_Hide) = Format(RS!QtyEnding, gs_formatQty)
            grid.TextMatrix(i, ColPrice_Hide) = Format(RS!AvgPrice, gs_formatAmountIDR)
            
            grid.Cell(flexcpBackColor, i, ColEndingBalance, i, ColPrice) = &H80000005
            
            RS.MoveNext
        Wend
    Else
'    If Not RS.EOF = False Then
        DtPeriod.Value = DateSerial(Year(DtPeriod.Value), Month(DtPeriod.Value) + 1, 0)
        i = i + 1
        grid.AddItem ""

        grid.TextMatrix(i, ColItemCode) = Trim(TxtProduk.Text)
        grid.TextMatrix(i, ColDate) = Format(DtPeriod.Value, "dd MMM yyyy")
        grid.TextMatrix(i, ColEndingBalance) = Format(0, gs_formatQty)
        grid.TextMatrix(i, ColPrice) = Format(0, gs_formatAmountIDR)
        grid.TextMatrix(i, ColAmount) = Format(0, gs_formatAmountIDR)
        grid.TextMatrix(i, ColValInventory) = Format(0, gs_formatAmountIDR)

        grid.TextMatrix(i, ColEndingBalance_Hide) = Format(0, gs_formatQty)
        grid.TextMatrix(i, ColPrice_Hide) = Format(0, gs_formatAmountIDR)

        grid.Cell(flexcpAlignment, i, ColItemCode) = flexAlignLeftCenter
        grid.Cell(flexcpAlignment, i, ColDate) = flexAlignRightCenter
        grid.TextMatrix(i, ColEndingBalance) = Format(0, gs_formatQty)
        grid.TextMatrix(i, ColPrice) = Format(0, gs_formatQty)
        grid.TextMatrix(i, ColAmount) = Format(0, gs_formatAmountIDR)
        grid.TextMatrix(i, ColValInventory) = Format(0, gs_formatAmountIDR)

        grid.TextMatrix(i, ColEndingBalance_Hide) = Format(0, gs_formatQty)
        grid.TextMatrix(i, ColPrice_Hide) = Format(0, gs_formatAmountIDR)

        grid.Cell(flexcpBackColor, i, ColEndingBalance, i, ColPrice) = &H80000005
   End If
    
    If RS.RecordCount >= 1 Then
        LblRecord = Format(RS.RecordCount - 1, "#,##0")
    Else
        LblRecord = Format("0", "#,##0")
    End If
    
    If RS.State <> adStateClosed Then RS.Close
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub CmdSubmit_Click()
    Dim adoLock As New ADODB.Connection
    Dim adoRs As New ADODB.Recordset
    Dim RS As New ADODB.Recordset
    Dim ubah As Boolean
    Dim Amount As Double
    
    On Error GoTo errHandler
    Me.MousePointer = vbHourglass
    
    LblErrMsg = ""
    
    If grid.Row = 0 Then
        LblErrMsg = DisplayMsg(8012): Me.MousePointer = vbDefault: Exit Sub
    End If
    
    If hakUpdate(Me.Name) = 0 Then _
            LblErrMsg = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
            
    SqlData = "SELECT Item_Code,Tgl,QtyEnding,AvgPrice,(QtyEnding * AvgPrice) Amount,(QtyEnding * AvgPrice) ValueInv " & vbCrLf & _
              "FROM dbo.Inventory_PriceDetail " & vbCrLf & _
              "WHERE Cls = 'Inv' AND Period = '" & Format(DtPeriod.Value, "yyyyMM") & "'" & vbCrLf
              
    If TxtProduk.Text <> "" Then
        SqlData = SqlData + "AND Item_Code = '" & Trim(TxtProduk.Text) & "'" & vbCrLf
    Else
        SqlData = SqlData + "AND QtyEnding < 0 " & vbCrLf
    End If
    
    SqlData = SqlData + "ORDER BY Item_Code,Tgl"
    
    If RS.State <> adStateClosed Then RS.Close
    RS.CursorLocation = adUseClient
    RS.Open SqlData, Db, adOpenKeyset, adLockOptimistic
            
    adoLock.ConnectionString = Db.ConnectionString
    adoLock.Open
    adoLock.BeginTrans
            
    
    If RS.EOF <> True Then
        With grid
        For i = 1 To .Rows - 1
            If CDbl(.TextMatrix(i, ColEndingBalance)) <> CDbl(.TextMatrix(i, ColEndingBalance_Hide)) Then
                ubah = True
            ElseIf CDbl(.TextMatrix(i, ColPrice)) <> CDbl(.TextMatrix(i, ColPrice_Hide)) Then
                ubah = True
            Else
                ubah = False
            End If
            
            If ubah = True Then
                SqlData = " UPDATE dbo.Inventory_Price " & vbCrLf & _
                            " SET Current_Price = " & Format(.TextMatrix(i, ColPrice), "####.00") & ", " & vbCrLf & _
                            "   Current_Stock = " & Format(.TextMatrix(i, ColEndingBalance), "####.00") & ", " & vbCrLf & _
                            "   Inventory_Price = " & Format(.TextMatrix(i, ColPrice), "####.00") & ", " & vbCrLf & _
                            "   Last_User = '" & userLogin & "', Last_Update = GETDATE() " & vbCrLf & _
                            " WHERE Item_Code = '" & .TextMatrix(i, ColItemCode) & "' " & vbCrLf & _
                            " AND Inventory_Year = " & Year(DtPeriod) & " AND Inventory_Month = " & Month(DtPeriod) & " "
                
                Db.Execute SqlData
                
                Amount = Format(CDbl(.TextMatrix(i, ColEndingBalance)) * CDbl(.TextMatrix(i, ColPrice)), "####.00")
                
                SqlData = " UPDATE dbo.Inventory_PriceDetail " & vbCrLf & _
                            " SET QtyEnding = " & Format(.TextMatrix(i, ColEndingBalance), "####.00") & ", " & vbCrLf & _
                            "   Price = " & Format(.TextMatrix(i, ColPrice), "####.00") & ", " & vbCrLf & _
                            "   AvgPrice = " & Format(.TextMatrix(i, ColPrice), "####.00") & ", " & vbCrLf & _
                            "   Amount = " & Amount & ", " & vbCrLf & _
                            "   AmountEnding = " & Amount & ", " & vbCrLf & _
                            "   Last_User = '" & userLogin & "', Last_Update = GETDATE() " & vbCrLf & _
                            " WHERE Item_Code = '" & .TextMatrix(i, ColItemCode) & "' AND Period = '" & Format(DtPeriod.Value, "yyyyMM") & "' AND Cls = 'Inv' "
                
                Db.Execute SqlData
                
            End If
        Next i
        End With
    Else
        With grid
        For i = 1 To .Rows - 1
            If CDbl(.TextMatrix(i, ColEndingBalance)) <> CDbl(.TextMatrix(i, ColEndingBalance_Hide)) Then
                ubah = True
            ElseIf CDbl(.TextMatrix(i, ColPrice)) <> CDbl(.TextMatrix(i, ColPrice_Hide)) Then
                ubah = True
            Else
                ubah = False
            End If
            
            If ubah = True Then
                SqlData = " INSERT INTO dbo.Inventory_Price " & vbCrLf & _
                            " (Current_Price, " & vbCrLf & _
                            " Current_Stock, " & vbCrLf & _
                            " Inventory_Price, " & vbCrLf & _
                            " Item_Code, " & vbCrLf & _
                            " Inventory_Year, " & vbCrLf & _
                            " Inventory_Month, " & vbCrLf & _
                            " Register_Date, " & vbCrLf & _
                            " Last_User, " & vbCrLf & _
                            " Duty_Status, Premonth_Stock, Premonth_Price, Incoming_Price, IncomingOther_Price, Outgoing_Price, OutgoingOther_Price) " & vbCrLf & _
                            " Values ( " & vbCrLf & _
                            " '" & Format(.TextMatrix(i, ColPrice), "####.00") & "', " & vbCrLf & _
                            " '" & Format(.TextMatrix(i, ColEndingBalance), "####.00") & "', " & vbCrLf & _
                            " '" & Format(.TextMatrix(i, ColPrice), "####.00") & "', " & vbCrLf & _
                            " '" & .TextMatrix(i, ColItemCode) & "', " & vbCrLf & _
                            " '" & Year(DtPeriod) & "', '" & Month(DtPeriod) & "', GetDate(), '" & userLogin & "', '" & 3 & "', '" & 0 & "','" & 0 & "','" & 0 & "','" & 0 & "','" & 0 & "','" & 0 & "') "
                
                Db.Execute SqlData
                
                Amount = Format(CDbl(.TextMatrix(i, ColEndingBalance)) * CDbl(.TextMatrix(i, ColPrice)), "####.00")
                
                SqlData = " INSERT INTO dbo.Inventory_PriceDetail " & vbCrLf & _
                            " (QtyEnding , " & vbCrLf & _
                            "  Price, " & vbCrLf & _
                            "  AvgPrice, " & vbCrLf & _
                            "  Amount, " & vbCrLf & _
                            "  AmountEnding, " & vbCrLf & _
                            "  Item_Code, " & vbCrLf & _
                            "  Period, " & vbCrLf & _
                            "  Cls, " & vbCrLf & _
                            " Register_Date, " & vbCrLf & _
                            " Last_User, Tgl, Idx, Seq_No, Duty_Status) " & vbCrLf & _
                            " Values ( " & vbCrLf & _
                            " '" & Format(.TextMatrix(i, ColEndingBalance), "####.00") & "', " & vbCrLf & _
                            " " & Format(.TextMatrix(i, ColPrice), "####.00") & ", " & vbCrLf & _
                            " " & Format(.TextMatrix(i, ColPrice), "####.00") & ", " & vbCrLf & _
                            " " & Amount & ", " & vbCrLf & _
                            " " & Amount & ", " & vbCrLf & _
                            " '" & .TextMatrix(i, ColItemCode) & "', " & vbCrLf & _
                            " '" & Format(DtPeriod.Value, "yyyyMM") & "', " & vbCrLf & _
                            " 'Inv', " & vbCrLf & _
                            " GetDate(), " & vbCrLf & _
                            " '" & userLogin & "',  DATEADD(dd, 0, DATEDIFF(dd, 0, GETDATE())), 0,1,3) "




                Db.Execute SqlData
                
            End If
        Next i
        End With
        End If
    
    
    adoLock.CommitTrans
    adoLock.Close
    
    cmdSearch_Click
    
    LblErrMsg.Caption = DisplayMsg(1101)
    
ErrExit:
    Me.MousePointer = vbDefault
    Set adoRs = Nothing
    Set adoLock = Nothing
    Exit Sub
errHandler:
    adoLock.RollbackTrans
    LblErrMsg.Caption = "[" & err.number & "] " & err.Description
    err.clear
    Resume ErrExit
End Sub

Private Sub Command1_Click()
    Me.MousePointer = vbHourglass
    
    frm_BrowseItem.getItemCode = TxtProduk.Text
    frm_BrowseItem.Show 1
    TxtProduk.Text = frm_BrowseItem.getItemCode
    TxtProduk.SetFocus
    
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    If gb_Simulation = True Then Call up_InitSimulation(Me)
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
    
    DtPeriod.Value = Format(Now, "MMM yyyy")
    TxtProduk = ""
    lblPartNumber.Caption = ""
    lblNm = ""
    
    SettingColumn
    Header
    
End Sub

Private Sub CmdSubMenu_Click()
    DoEvents
    frmMainMenu.Show
    DoEvents
    Unload Me
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If grid.TextMatrix(grid.Row, ColEndingBalance) = "" Then
        grid.TextMatrix(grid.Row, ColEndingBalance) = Format(0, gs_formatQty)
    Else
        grid.TextMatrix(grid.Row, ColEndingBalance) = Format(grid.TextMatrix(grid.Row, ColEndingBalance), gs_formatQty)
    End If
    
    grid.TextMatrix(grid.Row, ColAmount) = Format(CDbl(grid.TextMatrix(grid.Row, ColEndingBalance)) * CDbl(grid.TextMatrix(grid.Row, ColPrice)), gs_formatAmountIDR)
    grid.TextMatrix(grid.Row, ColValInventory) = Format(CDbl(grid.TextMatrix(grid.Row, ColEndingBalance)) * CDbl(grid.TextMatrix(grid.Row, ColPrice)), gs_formatAmountIDR)
    
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If grid.Col = ColEndingBalance Or grid.Col = ColPrice Then
        Cancel = False
    Else
        Cancel = True
    End If
End Sub

Private Sub grid_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
       KeyAscii = 0
    End If

    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
End Sub

Private Sub Grid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
       KeyAscii = 0
    End If

    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtProduk_Change()
    Dim sqlitem As String
    Dim RsItem As New ADODB.Recordset
    
    Header
    
    sqlitem = "Select * From Item_Master " & vbCrLf & _
              "Where Item_Code='" & Trim(Replace(TxtProduk, "'", "''") & "") & "' "
        
    If RsItem.State <> adStateClosed Then RsItem.Close
    
    Set RsItem = Db.Execute(sqlitem)
    
    If Not RsItem.EOF Then
        lblPartNumber.Caption = Trim(RsItem("MakerItem_Code") & "")
        lblNm.Caption = Trim(RsItem("Item_Name") & "")
    Else
        lblPartNumber.Caption = ""
        lblNm.Caption = ""
    End If
End Sub
