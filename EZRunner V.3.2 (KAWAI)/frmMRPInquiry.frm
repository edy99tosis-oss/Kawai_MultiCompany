VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMRPInquiry 
   BackColor       =   &H00FDDFE3&
   Caption         =   "MRP Inquiry"
   ClientHeight    =   10980
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   15120
   Icon            =   "frmMRPInquiry.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   10980
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   510
      TabIndex        =   15
      Top             =   9210
      Width           =   14205
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
         Left            =   105
         TabIndex        =   16
         Top             =   195
         Width           =   13980
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   1695
      Left            =   495
      TabIndex        =   10
      Top             =   1080
      Width           =   14205
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   300
         Left            =   4485
         TabIndex        =   19
         Top             =   720
         Width           =   300
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H0000FFFF&
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
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1125
         Width           =   1140
      End
      Begin MSComCtl2.DTPicker dtAwal 
         Height          =   330
         Left            =   1650
         TabIndex        =   2
         Top             =   1140
         Width           =   1785
         _ExtentX        =   3149
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
         Format          =   151388163
         UpDown          =   -1  'True
         CurrentDate     =   37860
      End
      Begin MSComCtl2.DTPicker dtAkhir 
         Height          =   330
         Left            =   4200
         TabIndex        =   3
         Top             =   1140
         Width           =   1665
         _ExtentX        =   2937
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
         Format          =   151388163
         UpDown          =   -1  'True
         CurrentDate     =   37891
      End
      Begin VB.Line Line2 
         X1              =   4860
         X2              =   9390
         Y1              =   990
         Y2              =   990
      End
      Begin VB.Label lblNm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Name"
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
         Left            =   4860
         TabIndex        =   18
         Top             =   765
         Width           =   1245
      End
      Begin MSForms.ComboBox cbo 
         Height          =   315
         Index           =   1
         Left            =   1650
         TabIndex        =   1
         Top             =   720
         Width           =   2775
         VariousPropertyBits=   746604571
         MaxLength       =   25
         DisplayStyle    =   3
         Size            =   "4895;556"
         ListRows        =   15
         ShowDropButtonWhen=   2
         Value           =   "AAAAAAAAAAAAAAAAAAAAAAAAA"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
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
         Index           =   3
         Left            =   270
         TabIndex        =   17
         Top             =   780
         Width           =   1155
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
         Index           =   0
         Left            =   270
         TabIndex        =   14
         Top             =   360
         Width           =   1140
      End
      Begin MSForms.ComboBox cbo 
         Height          =   315
         Index           =   0
         Left            =   1650
         TabIndex        =   0
         Top             =   300
         Width           =   1875
         VariousPropertyBits=   746604571
         MaxLength       =   10
         DisplayStyle    =   3
         Size            =   "3307;556"
         ListRows        =   15
         ShowDropButtonWhen=   2
         Value           =   "AAAAAA"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
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
         Left            =   3735
         TabIndex        =   13
         Top             =   1215
         Width           =   165
      End
      Begin VB.Label Label1 
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
         Left            =   270
         TabIndex        =   12
         Top             =   1215
         Width           =   540
      End
      Begin VB.Label lblNm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
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
         Left            =   3825
         TabIndex        =   11
         Top             =   360
         Width           =   1395
      End
      Begin VB.Line Line1 
         X1              =   3825
         X2              =   9405
         Y1              =   600
         Y2              =   600
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
      Left            =   510
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9885
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   "To Inquiry"
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
      Left            =   13185
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9885
      Width           =   1530
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   6255
      Left            =   510
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2895
      Width           =   14205
      _cx             =   25056
      _cy             =   11033
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
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   1
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
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   12870
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   390
      Width           =   1845
      _extentx        =   3254
      _extenty        =   714
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MRP Inquiry"
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
      Left            =   510
      TabIndex        =   9
      Top             =   390
      Width           =   14205
   End
End
Attribute VB_Name = "frmMRPInquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i As Integer
Dim nilKosong As Boolean
Dim dealerCD As String
Dim tampungDtAwal As Byte, tampungDtAkhir As Byte

Dim bteColSelect As Byte
Dim bteColProduct As Byte
Dim bteColDesc As Byte
Dim bteColPartNo As Byte
Dim bteColLotNo As Byte
Dim bteColProdDate As Byte
Dim bteColPlan As Byte
Dim bteColOffQty As Byte

Private Sub headerGrid()
    
    bteColSelect = 0
    bteColProduct = 1
    bteColDesc = 2
    bteColPartNo = 3
    bteColLotNo = 4
    bteColProdDate = 5
    bteColPlan = 6
    bteColOffQty = 7
    
    With grid
        .clear
        .ColS = 8
        .Rows = 1
        
        .TextMatrix(0, bteColSelect) = "S"
        .TextMatrix(0, bteColProduct) = "Parent Code (Product Code)"
        .TextMatrix(0, bteColDesc) = "Description"
        .TextMatrix(0, bteColPartNo) = "Parts No."
        .TextMatrix(0, bteColLotNo) = "Lot No"
        .TextMatrix(0, bteColProdDate) = "Production Date"
        .TextMatrix(0, bteColPlan) = "Plan"
        .TextMatrix(0, bteColOffQty) = "Off Qty"
        
        .ColWidth(bteColSelect) = 300 'besar kolom
        .ColWidth(bteColProduct) = 2500
        .ColWidth(bteColDesc) = 3700
        .ColWidth(bteColPartNo) = 2500
        .ColWidth(bteColLotNo) = 1000
        .ColWidth(bteColProdDate) = 1750
        .ColWidth(bteColPlan) = 1500
        .ColWidth(bteColOffQty) = 1500
        
        .ColAlignment(bteColSelect) = flexAlignCenterCenter
        .ColAlignment(bteColProduct) = flexAlignLeftCenter
        .ColAlignment(bteColDesc) = flexAlignLeftCenter
        .ColAlignment(bteColPartNo) = flexAlignLeftCenter
        .ColAlignment(bteColLotNo) = flexAlignCenterCenter
        .ColAlignment(bteColProdDate) = flexAlignCenterCenter
        .ColAlignment(bteColPlan) = flexAlignRightCenter
        .ColAlignment(bteColOffQty) = flexAlignRightCenter
        
        .EditMaxLength = 1
    End With

End Sub

Sub Kosong(Optional mulai As Integer)
    nilKosong = True
    If mulai = 0 Then Cbo(0) = ""
    lblNm(0) = ""
    Cbo(1).ListIndex = 0: lblNm(1) = Cbo(1).Column(1)
    dtAwal = Format(Now, "MMM yyyy")
    dtAkhir = Format(Now, "MMM yyyy")
    nilKosong = False
End Sub

'******** Combo **********
Sub isiCboCust() 'Isi Combo Dealer CD dr Customer Master
Dim RsCust As New ADODB.Recordset 'Data Customer

With Cbo(0)
    .clear
    .columnCount = 2
    .TextColumn = 1
    
    '******** Ambil dr Customer Master utk Combo Dealer CD
    sql = "select Trade_Code,Trade_Name from Trade_Master " & _
        "where trade_code in (select distinct manufacture_code from manufacture_line) " & _
        "order by Trade_Code"
    Set RsCust = Db.Execute(sql)
    
    i = 0
    Do While Not (RsCust.EOF)
        .AddItem ""
        .List(i, 0) = Trim(RsCust(0))
        .List(i, 1) = Trim(RsCust(1))
        i = i + 1
        RsCust.MoveNext
    Loop
    
    .Text = ""
    .ListWidth = 380
    .ColumnWidths = "80 pt;300 pt"
    
    Set RsCust = Nothing
End With
End Sub

Sub isiCboItem()
Dim rscbo As New ADODB.Recordset

With Cbo(1)
    .clear
    .columnCount = 2
    .TextColumn = 1
    
    sql = "select a.Item_Code, Item_Name from Item_master a where a.use_endday > convert(char(8), getdate(), 112) order by Item_Code"
    Set rscbo = Db.Execute(sql)
    
    .AddItem ""
    .List(0, 0) = strAll
    .List(0, 1) = strAll
    i = 1
    Do While Not (rscbo.EOF)
        .AddItem ""
        .List(i, 0) = Trim(rscbo(0))
        .List(i, 1) = Trim(rscbo(1))
        i = i + 1
        rscbo.MoveNext
    Loop
    
    .Text = ""
    .ListWidth = 450
    .ColumnWidths = "150 pt;300 pt"
    
    Set rscbo = Nothing
End With
End Sub

Private Sub command2_Click()
    Me.MousePointer = vbHourglass
    frm_BrowseItem.getItemCode = Cbo(1).Text
    frm_BrowseItem.Show 1
    Cbo(1).Text = frm_BrowseItem.getItemCode
    Cbo(1).SetFocus
    Me.MousePointer = vbDefault
End Sub

'******************
Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
    
    Call isiCboCust
    Call isiCboItem
    Call Kosong
    Call headerGrid
End Sub

Function checkErr() As Boolean
checkErr = False
    If Trim(Cbo(0)) = "" Then
        LblErrMsg = DisplayMsg(1040) 'Please Input Factore Code
        Cbo(0).SetFocus: Exit Function
    ElseIf Cbo(0).MatchFound = False Then
        LblErrMsg = DisplayMsg(4060) 'Factory code not found!
        Cbo(0).SetFocus: Exit Function
    ElseIf Trim(Cbo(1)) = "" Then
        LblErrMsg = DisplayMsg(1009) 'Please Input Product Code
        Cbo(1).SetFocus: Exit Function
    ElseIf Cbo(1).MatchFound = False Then
        LblErrMsg = DisplayMsg(4061) 'Product Code not found!
        Cbo(1).SetFocus: Exit Function
    ElseIf Format(dtAwal, "yyyyMM") > Format(dtAkhir, "yyyyMM") Then
        LblErrMsg = DisplayMsg(4068)
        dtAwal.SetFocus: Exit Function
    End If
checkErr = True
End Function

Private Sub cmdSearch_Click()
    LblErrMsg = ""
    If checkErr Then Call IsiGrid
End Sub

Sub IsiGrid()
Dim rsProd As New ADODB.Recordset
Dim tglAwal As String, tglAkhir As String

If nilKosong = True Then Exit Sub
With grid
    Call headerGrid
    
    tglAwal = Year(dtAwal) & "-" & Format(Month(dtAwal), "00") & "-01"
    tglAkhir = Year(dtAkhir) & "-" & Format(Month(dtAkhir), "00") & "-" & DateDiff("d", dtAkhir, DateAdd("m", 1, dtAkhir))
    
    sql = "select a.Item_Code, Item_Name, MakerITem_Code,Lot_No,Actual_Date production_date, " & _
            "qty = sum(qty), isnull(sum(a.off_qty),0) as off_qty " & _
        "from actual_production a,Item_Master b " & _
        "where b.Item_Code = a.Item_Code " & _
            "And Actual_Date >= '" & tglAwal & "' and Actual_date<= '" & tglAkhir & "' " & _
            "And Manufacture_Code = '" & Cbo(0) & "' "
    
    If Cbo(1).ListIndex <> 0 Then sql = sql & "And a.Item_Code = '" & Cbo(1) & "' "
    sql = sql & _
        "Group by a.Item_Code, Item_Name, MakerITem_Code,Lot_No,Actual_Date " & _
        "order by a.Item_Code, Actual_Date"
    Set rsProd = Db.Execute(sql)
    
    i = 1
    If Not rsProd.EOF Then
        Do While Not rsProd.EOF
            .Rows = .Rows + 1
            .Cell(flexcpChecked, i, bteColSelect) = flexUnchecked
            .Cell(flexcpBackColor, i, bteColSelect) = vbWhite
            
            .TextMatrix(i, bteColProduct) = Trim(rsProd("Item_Code"))
            .TextMatrix(i, bteColDesc) = Trim(rsProd("Item_Name"))
            .TextMatrix(i, bteColPartNo) = Trim(rsProd("MakerITem_Code"))
            .TextMatrix(i, bteColLotNo) = Trim(rsProd("Lot_No"))
            .TextMatrix(i, bteColProdDate) = Format(Trim(rsProd("PRoduction_Date")), "dd MMM yyyy")
            .TextMatrix(i, bteColPlan) = Format(rsProd("Qty"), gs_formatQty)
            .TextMatrix(i, bteColOffQty) = Format(rsProd("Off_Qty"), gs_formatQty)
            
            i = i + 1
            rsProd.MoveNext
        Loop
    Else
        LblErrMsg = DisplayMsg(4006)
    End If
    Set rsProd = Nothing
End With
End Sub

'************** Tampilkan Data ***************
Private Sub cbo_Change(Index As Integer)
    LblErrMsg = ""
    Call cbo_Click(Index)
End Sub

Public Sub cbo_Click(Index As Integer)
    If nilKosong Then Exit Sub
    
    Cbo(Index) = Cbo(Index)
    If Cbo(Index).MatchFound Then lblNm(Index) = Cbo(Index).Column(1) Else lblNm(Index) = ""
    Call headerGrid
End Sub

Private Sub dtAwal_Change()
    Call dtAwal_Click
    tampungDtAwal = dtAwal.Month
    If Format(dtAwal, "yyyyMM") > Format(dtAkhir, "yyyyMM") Then
        LblErrMsg = DisplayMsg(4068)
    End If
End Sub

Private Sub dtAwal_Click()
If dtAwal.Month = 1 And Val(tampungDtAwal) = 12 Then dtAwal.Year = dtAwal.Year + 1
If dtAwal.Month = 12 And Val(tampungDtAwal) = 1 Then dtAwal.Year = dtAwal.Year - 1
End Sub

Private Sub dtAkhir_Change()
    Call dtAkhir_Click
    tampungDtAkhir = dtAkhir.Month
    If Format(dtAwal, "yyyyMM") > Format(dtAkhir, "yyyyMM") Then
        LblErrMsg = DisplayMsg(4066)
    End If
End Sub

Private Sub dtAkhir_Click()
If dtAkhir.Month = 1 And Val(tampungDtAkhir) = 12 Then dtAkhir.Year = dtAkhir.Year + 1
If dtAkhir.Month = 12 And Val(tampungDtAkhir) = 1 Then dtAkhir.Year = dtAkhir.Year - 1
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col <> bteColSelect Then Cancel = 1
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim tampung As Long
With grid
    If Row <> 0 And Col = bteColSelect Then
        If .Cell(flexcpChecked, Row, bteColSelect) = flexChecked Then
            tampung = Row
            For i = 1 To .Rows - 1
                .Cell(flexcpChecked, i, bteColSelect) = flexUnchecked
            Next i
            .Cell(flexcpChecked, tampung, bteColSelect) = flexChecked
        End If
    End If
End With
End Sub

Private Sub Command1_Click(Index As Integer)
Dim cek As Integer

Me.MousePointer = vbHourglass
Select Case Index
    Case 0: 'Submit
            With grid
                cek = 0
                For i = 1 To .Rows - 1
                    If .Cell(flexcpChecked, i, bteColSelect) = flexChecked Then
                        cek = i
                        Exit For
                    End If
                Next i
                
                If cek = 0 Then
                    LblErrMsg = DisplayMsg(8011)
                Else
                    LblErrMsg = ""
                    frmMRPInquiry2.lblNm(0) = Trim(.TextMatrix(cek, bteColProduct))
                    frmMRPInquiry2.lblNm(1) = Trim(.TextMatrix(cek, bteColDesc))
                    frmMRPInquiry2.lblNm(2) = Trim(.TextMatrix(cek, bteColPartNo))
                    frmMRPInquiry2.lblNm(3) = Trim(.TextMatrix(cek, bteColLotNo))
                    frmMRPInquiry2.lblNm(4) = Trim(.TextMatrix(cek, bteColProdDate))
                    frmMRPInquiry2.lblNm(5) = Trim(.TextMatrix(cek, bteColPlan))
                    frmMRPInquiry2.lblNm(6) = Trim(.TextMatrix(cek, bteColOffQty))
                    frmMRPInquiry2.Show
                    Call frmMRPInquiry2.IsiGrid(Trim(.TextMatrix(cek, bteColProduct)), Trim(.TextMatrix(cek, bteColLotNo)), Trim(Cbo(0)))
                    Me.Hide
                End If
            End With
End Select
Me.MousePointer = vbDefault
End Sub


'************ Unload **********
Private Sub CmdSubMenu_Click()
    DoEvents
    frmMainMenu.Show
    DoEvents
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then Unload Me Else LblErrMsg.Caption = ErrMsg
End Sub
'**************

