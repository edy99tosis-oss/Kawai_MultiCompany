VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmPOAdjustment_Price 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Purchase Order Adjustment Price"
   ClientHeight    =   10485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   Icon            =   "FrmPOAdjustment_Price.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   10485
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
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
      Left            =   13740
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   10020
      Width           =   1185
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   345
      TabIndex        =   18
      Top             =   9330
      Width           =   14595
      Begin VB.Label LblErr 
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
         Height          =   255
         Left            =   105
         TabIndex        =   19
         Top             =   195
         Width           =   14370
      End
   End
   Begin VB.CommandButton CmdCancel 
      BackColor       =   &H0080FFFF&
      Caption         =   "Ca&ncel"
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
      Left            =   12450
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   10020
      Width           =   1185
   End
   Begin VB.CommandButton cmdClear 
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
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   10020
      Width           =   1185
   End
   Begin VB.CommandButton CmdMenu 
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
      Left            =   345
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   10020
      Width           =   1185
   End
   Begin VB.TextBox txtamount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
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
      Left            =   12975
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   8865
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1695
      Left            =   345
      TabIndex        =   2
      Top             =   870
      Width           =   14595
      Begin VB.TextBox lblAddr 
         BackColor       =   &H00FDDFE3&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   8190
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   300
         Width           =   5835
      End
      Begin VB.CommandButton cmdsearch 
         BackColor       =   &H0080FFFF&
         Caption         =   "Sea&rch"
         Height          =   375
         Left            =   6270
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1200
         Width           =   1140
      End
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
         Left            =   3420
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   300
         Width           =   3510
      End
      Begin MSComCtl2.DTPicker PODate 
         Height          =   345
         Left            =   1785
         TabIndex        =   4
         Top             =   720
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
         Format          =   294780931
         CurrentDate     =   37868
      End
      Begin MSComCtl2.DTPicker PODate2 
         Height          =   345
         Left            =   3720
         TabIndex        =   5
         Top             =   720
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
         Format          =   294780931
         CurrentDate     =   37868
      End
      Begin VB.Line Line2 
         Index           =   1
         X1              =   8175
         X2              =   14055
         Y1              =   585
         Y2              =   585
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   255
         Index           =   4
         Left            =   7335
         TabIndex        =   22
         Top             =   330
         Width           =   840
      End
      Begin VB.Line Line4 
         X1              =   3420
         X2              =   6920
         Y1              =   585
         Y2              =   585
      End
      Begin MSForms.ComboBox CboCust 
         Height          =   315
         Left            =   1785
         TabIndex        =   11
         Top             =   255
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
         Left            =   285
         TabIndex        =   10
         Top             =   330
         Width           =   1275
      End
      Begin VB.Label Label1 
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
         Left            =   285
         TabIndex        =   9
         Top             =   795
         Width           =   705
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
         Left            =   3420
         TabIndex        =   8
         Top             =   795
         Width           =   165
      End
      Begin VB.Label LblPart 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PO No"
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
         Left            =   285
         TabIndex        =   7
         Top             =   1275
         Width           =   525
      End
      Begin MSForms.ComboBox CboPOnO 
         Height          =   315
         Left            =   1785
         TabIndex        =   6
         Top             =   1215
         Width           =   2895
         VariousPropertyBits=   746604571
         MaxLength       =   35
         DisplayStyle    =   3
         Size            =   "5106;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   420
      Left            =   13080
      TabIndex        =   1
      Top             =   270
      Width           =   1860
      _extentx        =   3281
      _extenty        =   741
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   5925
      Left            =   345
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   2685
      Width           =   14550
      _cx             =   25665
      _cy             =   10451
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
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
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
      Left            =   11610
      TabIndex        =   13
      Top             =   8925
      Width           =   1140
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Order Adjustment Price"
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
      Left            =   315
      TabIndex        =   0
      Top             =   300
      Width           =   14640
   End
End
Attribute VB_Name = "FrmPOAdjustment_Price"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset, rsGrid As New ADODB.Recordset
Dim bteColPartNo As Byte
Dim bteColPartName As Byte
Dim bteColQty As Byte
Dim bteColPrice As Byte
Dim bteColAdjPrice As Byte
Dim bteColAmount As Byte
Dim BtnCOmbo As Boolean
Dim bteColSelect As Byte
Dim DblJml As Double
Dim Baris As Integer
Dim BrsSama As Boolean


Function adtocboCust()
Dim sqlcust As String
Dim RsCust As New Recordset

    sqlcust = "SELECT  rtrim(Trade_Master.trade_Code) supp_code, rtrim(Trade_Master.Trade_Name) supp_name, " & _
        "rtrim(Trade_Master.Address1) address, country_Cls, POPayment_Day From Trade_Master where trade_cls in ('2','3')"
    Set RsCust = Db.Execute(sqlcust)

    With cboCust
        .clear
        .columnCount = 4
        .ColumnWidths = "80 pt;280 pt; 0 pt; 0 pt; 0 pt"
        .ListWidth = 360
        .ListRows = 15
        i = 0
        RsCust.Requery
        If Not RsCust.EOF And Not RsCust.BOF Then
            Do Until RsCust.EOF
                .AddItem ""
                .List(i, 0) = IIf(IsNull(Trim(RsCust!supp_code)), "", Trim(RsCust!supp_code))
                .List(i, 1) = IIf(IsNull(Trim(RsCust!supp_name)), "", Trim(RsCust!supp_name))
                .List(i, 2) = IIf(IsNull(Trim(RsCust!Address)), "", Trim(RsCust!Address))
                .List(i, 3) = IIf(IsNull(Trim(RsCust!country_cls)), "", Trim(RsCust!country_cls))
                .List(i, 4) = IIf(IsNull(Val(RsCust!POPayment_Day & "")), "", Val(RsCust!POPayment_Day & ""))
                i = i + 1
                RsCust.MoveNext
            Loop
        End If
    End With
    Set RsCust = Nothing
     
End Function
Sub Header()
  With grid
    .clear

    .Rows = 1
    .ColS = 6
    
   ' BteColSelect = 0
    bteColPartNo = 0
    bteColPartName = 1
    bteColQty = 2
    bteColPrice = 3
    bteColAdjPrice = 4
    bteColAmount = 5


    '.ColWidth(BteColSelect) = 300
    .ColWidth(bteColPartNo) = 3000
    .ColWidth(bteColPartName) = 4000
    .ColWidth(bteColQty) = 1500
    .ColWidth(bteColPrice) = 2000
    .ColWidth(bteColAdjPrice) = 2000
    .ColWidth(bteColAmount) = 2000
    
    
   ' .TextMatrix(0, BteColSelect) = ""
    .TextMatrix(0, bteColPartNo) = "Part No"
    .TextMatrix(0, bteColPartName) = "Part Name"
    .TextMatrix(0, bteColQty) = "Qty PO"
    
    .TextMatrix(0, bteColPrice) = "PO Price"
    .TextMatrix(0, bteColAdjPrice) = "Adj Price"
    .TextMatrix(0, bteColAmount) = "Amount"
    
    
    If hakPrice(Me.Name) <> 1 Then
    
        .ColHidden(bteColPrice) = True
        .ColHidden(bteColAdjPrice) = True
    End If
    
    
    
    .Cell(flexcpAlignment, 0, 0, 0, .ColS - 1) = flexAlignCenterCenter
    For i = colpoqty To ColAmountService
      .ColAlignment(i) = flexAlignRightCenter
    Next i
    
    .RowHeight(0) = 250
  End With
  
End Sub



Private Sub CboCust_Change()
    CboPOnO = ""
    BtnCOmbo = True
End Sub

Private Sub cboCust_Click()
Dim ketemu As Boolean

    LblErr = ""
    ketemu = False


    If cboCust.ListIndex <> -1 Then

        lblSupp.Text = cboCust.Column(1)
        lblAddr.Text = cboCust.Column(2)
        End If


End Sub




Sub adtocbopono()
Dim sqlno As String
Dim rsno As New Recordset
If Trim(lblSupp.Text) = "" Then Exit Sub
    sqlno = " select * From Purchaseorder_master " & vbCrLf & _
                      "     Where PO_Date>='" & Format(PODate, "dd-MMM-YYYY") & "' And PO_Date<='" & Format(PODate2, "dd-MMM-YYYY") & "' " & vbCrLf & _
                      "         And Supplier_Code='" & Trim(cboCust) & "' "
            
    Set rsno = Db.Execute(sqlno)
    With CboPOnO
        .clear
        Do While Not rsno.EOF
            .AddItem Trim(rsno("PO_No"))
            rsno.MoveNext
        Loop
        .ColumnWidths = "150pt"
        .ListWidth = 150
        .ListRows = 15
    End With
    Set rsno = Nothing
End Sub

Private Sub cboCust_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call cboCust_Click
End Sub

Private Sub cbocust_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = Asc("'") Then KeyAscii = 0
    Call cboCust_Click
End Sub

Private Sub cbocust_LostFocus()
    If Trim(cboCust) = "" Then LblErr = "": Exit Sub
    If cboCust.MatchFound = False Then
        lblAddr.Text = ""
        lblSupp.Text = ""
        LblErr = DisplayMsg("0032")
        'CboCust.SetFocus

        Exit Sub
    Else
         LblErr = ""
    End If
    
    Call cboCust_Click
        'Call adtocbopono
End Sub

Private Sub CboPOnO_Change()
    Call Header
End Sub



Private Sub CboPOnO_DropButtonClick()
If BtnCOmbo = True Then
Call adtocbopono
BtnCOmbo = False
End If
End Sub

Private Sub CboPOnO_GotFocus()
If BtnCOmbo = True Then
Call adtocbopono
BtnCOmbo = False
End If
End Sub

Private Sub CboPOnO_LostFocus()
      If Trim(CboPOnO) = "" Then LblErr = "": Exit Sub
    If CboPOnO.MatchFound = False Then
        LblErr = DisplayMsg("4015"): Exit Sub
    Else
         LblErr = ""
    End If
End Sub

Private Sub cmdCancel_Click()
   Call cmdSearch_Click
End Sub

Private Sub cmdClear_Click()
    Header
    Call adtocboCust
    lblAddr = ""
    lblSupp = ""
End Sub

Private Sub CmdMenu_Click()

    Unload Me
    frmMainMenu.Show
End Sub

Private Sub cmdSearch_Click()

If Trim(cboCust) = "" Then LblErr = DisplayMsg("1054"): Exit Sub 'Please Selecet Supplier Code
If Trim(CboPOnO) = "" Then LblErr = DisplayMsg("9001"): Exit Sub 'Please Select PO No.

Call Header
DblJml = 0
Dim StrOPen As String
StrOPen = "Select* , " & _
                " (Select Item_name from  Item_Master  Where Item_Code = A.item_Code) Item_Name2 " & _
                " From PurChaseOrder_Detail A " & _
                " Where Po_NO = '" & Trim(CboPOnO) & "'"
If rsGrid.State <> adStateClosed Then rsGrid.Close
rsGrid.Open StrOPen, Db, adOpenDynamic, adLockOptimistic
Do While Not rsGrid.EOF
    With grid
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, bteColPartNo) = IIf(IsNull(rsGrid("Item_Code")), "", Trim(rsGrid("Item_Code")))
    
        .TextMatrix(.Rows - 1, bteColPartName) = IIf(IsNull(rsGrid("Item_name2")), "", Trim(rsGrid("Item_name2")))
        .TextMatrix(.Rows - 1, bteColQty) = Format(rsGrid("Qty"), gs_formatQty)
        .TextMatrix(.Rows - 1, bteColPrice) = Format(rsGrid("Price"), gs_formatAmount)
        .TextMatrix(.Rows - 1, bteColAdjPrice) = Format(IIf(IsNull(rsGrid("Price_adj")), "0", Trim(rsGrid("Price_Adj"))), gs_formatAmount)
        .TextMatrix(.Rows - 1, bteColAmount) = Format(rsGrid("Amount"), gs_formatAmount)
        DblJml = DblJml + rsGrid("Amount")
      '  .ColAlignment(BteColSelect) = flexAlignCenterCenter
        
    End With
    rsGrid.MoveNext
Loop
If grid.Rows > 1 Then
    grid.Cell(flexcpAlignment, 1, bteColPartNo, grid.Rows - 1, bteColPartNo) = flexAlignLeftCenter
    grid.Cell(flexcpBackColor, 1, bteColAdjPrice, grid.Rows - 1, bteColAdjPrice) = &HFFFFFF
   ' Grid.Cell(flexcpBackColor, 1, BteColSelect, Grid.Rows - 1, BteColSelect) = &HFFFFFF

End If
rsGrid.Close
txtamount = Format(DblJml, gs_formatAmount)
End Sub

Private Sub CmdSubmit_Click()
    On Error GoTo ErrH
    If hakUpdate(Me.Name) <> 1 Then LblErr = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
    
    Db.BeginTrans
     For km = 1 To grid.Rows - 1
            Db.Execute "Update PurchaseOrder_detail set Amount= " & CDbl(Trim(grid.TextMatrix(km, bteColAmount))) & "  , Price_Adj = " & CDbl(Trim(grid.TextMatrix(km, bteColAdjPrice))) & ", " & _
                              " Last_User='" & userLogin & "', Last_Update=GetDate()  Where Po_No = '" & Trim(CboPOnO) & "' and Item_Code = '" & Trim(grid.TextMatrix(km, bteColPartNo)) & "'"
            
     Next km
     Db.Execute "Update PurchaseOrder_master Set Amount =" & CDbl(txtamount) & " , Total_Amount = " & CDbl(txtamount) & ", " & _
                              " Last_User='" & userLogin & "', Last_Update=GetDate()  Where Po_No ='" & Trim(CboPOnO) & "'"
     Db.CommitTrans
     LblErr = DisplayMsg("1000")
     Exit Sub
     
ErrH:
     Db.RollbackTrans
     LblErr = err.Description
     
End Sub

Private Sub Form_Load()

    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
    
    PODate = Now()
    PODate2 = Now()
    
    Call adtocboCust
    Call Header

End Sub
Sub Total()
DblJml = 0
    For km = 1 To grid.Rows - 1
        DblJml = DblJml + CDbl(grid.TextMatrix(km, bteColAmount))
    Next km
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RS = Nothing
    Set rsGrid = Nothing
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
lblerror = ""
If grid.Col = bteColAdjPrice Then
    If Trim(grid.Text) = "" Then grid.Text = "0"
    grid.TextMatrix(Row, bteColAdjPrice) = Val(grid.TextMatrix(Row, bteColAdjPrice))
    grid.TextMatrix(Row, bteColAdjPrice) = Format(grid.TextMatrix(Row, bteColAdjPrice), gs_formatAmount)

    If Val(Trim(grid.Text)) = 0 Then
        grid.TextMatrix(Row, bteColAmount) = Format((CDbl(grid.TextMatrix(Row, bteColQty)) * CDbl(grid.TextMatrix(Row, bteColPrice))), gs_formatAmount)
    Else
       grid.TextMatrix(Row, bteColAmount) = Format((CDbl(grid.TextMatrix(Row, bteColQty)) * CDbl(grid.TextMatrix(Row, bteColAdjPrice))), gs_formatAmount)
    End If
    
    Call Total
    txtamount.Text = Format(DblJml, gs_formatAmount)
End If

End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
LblErr.Caption = ""
 If grid.Col <> bteColAdjPrice Then
 Cancel = True
End If
End Sub

Private Sub grid_Click()
If grid.Col = bteColAdjPrice Then
    grid.FocusRect = flexFocusInset
   Else
    grid.FocusRect = flexFocusNone
End If
End Sub

Private Sub Grid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)

    If Col = bteColAdjPrice Then
    If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 46 And KeyAscii <> 8 Then KeyAscii = 0
    End If
    
    
End Sub

Private Sub Grid_KeyUp(KeyCode As Integer, Shift As Integer)
If grid.Col = bteColAdjPrice Then
    grid.FocusRect = flexFocusInset
   Else
    grid.FocusRect = flexFocusNone
End If
End Sub

Private Sub PODate_Change()
    CboPOnO = ""
    BtnCOmbo = True
End Sub

Private Sub PODate2_Change()
    CboPOnO = ""
    BtnCOmbo = True
End Sub

