VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmMRPInquiry2 
   BackColor       =   &H00FDDFE3&
   Caption         =   "MRP Inquiry Detail"
   ClientHeight    =   10980
   ClientLeft      =   75
   ClientTop       =   420
   ClientWidth     =   15120
   Icon            =   "frmMRPInquiry2.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10980
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   1245
      Left            =   225
      TabIndex        =   6
      Top             =   1005
      Width           =   14805
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Off Qty"
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
         Index           =   6
         Left            =   11700
         TabIndex        =   20
         Top             =   765
         Width           =   615
      End
      Begin VB.Label lblNm 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sebango"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   12675
         TabIndex        =   19
         Top             =   765
         Width           =   750
      End
      Begin VB.Line Line1 
         Index           =   6
         X1              =   12375
         X2              =   13425
         Y1              =   975
         Y2              =   975
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   10395
         X2              =   11445
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label lblNm 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sebango"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   10695
         TabIndex        =   18
         Top             =   750
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plan"
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
         Index           =   5
         Left            =   9945
         TabIndex        =   17
         Top             =   750
         Width           =   360
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   8385
         X2              =   9735
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label lblNm 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sebango"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   8385
         TabIndex        =   16
         Top             =   750
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Production Date"
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
         Index           =   4
         Left            =   6855
         TabIndex        =   15
         Top             =   750
         Width           =   1365
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   5745
         X2              =   6585
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label lblNm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sebango"
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
         Left            =   5745
         TabIndex        =   14
         Top             =   750
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lot No"
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
         Left            =   4635
         TabIndex        =   13
         Top             =   750
         Width           =   540
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   1380
         X2              =   4000
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label lblNm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sebango"
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
         Left            =   1380
         TabIndex        =   12
         Top             =   750
         Width           =   2550
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Parts No."
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
         Left            =   195
         TabIndex        =   11
         Top             =   750
         Width           =   780
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   5745
         X2              =   11445
         Y1              =   510
         Y2              =   510
      End
      Begin VB.Label lblNm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sebango"
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
         Left            =   5745
         TabIndex        =   10
         Top             =   300
         Width           =   5040
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Index           =   1
         Left            =   4635
         TabIndex        =   9
         Top             =   300
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Parent Code "
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
         Left            =   195
         TabIndex        =   8
         Top             =   300
         Width           =   1125
      End
      Begin VB.Label lblNm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sebango"
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
         Left            =   1380
         TabIndex        =   7
         Top             =   300
         Width           =   2550
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   1380
         X2              =   4000
         Y1              =   510
         Y2              =   510
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   600
      Left            =   225
      TabIndex        =   4
      Top             =   9210
      Width           =   14805
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
         Height          =   285
         Left            =   105
         TabIndex        =   5
         Top             =   195
         Width           =   14580
      End
   End
   Begin VB.CommandButton cmdSubMenu 
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
      Left            =   225
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9915
      Width           =   1140
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   6795
      Left            =   225
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2370
      Width           =   14805
      _cx             =   26114
      _cy             =   11986
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
      Editable        =   0
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
      Left            =   13185
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   323
      Width           =   1845
      _extentx        =   3254
      _extenty        =   714
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MRP Inquiry Detail"
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
      Left            =   225
      TabIndex        =   3
      Top             =   330
      Width           =   14805
   End
End
Attribute VB_Name = "frmMRPInquiry2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer

Dim bteColProd As Byte
Dim bteColDesc As Byte
Dim bteColPartNo As Byte
Dim bteColWH As Byte
Dim bteColAddress As Byte
Dim bteColReqDate As Byte
Dim bteColPlan As Byte
Dim bteColOffQty As Byte
Dim bteColResult As Byte
Dim bteColUnit As Byte

Private Sub headerGrid()

    bteColProd = 0
    bteColDesc = 1
    bteColPartNo = 2
    bteColWH = 3
    bteColAddress = 4
    bteColReqDate = 5
    bteColPlan = 6
    bteColOffQty = 7
    bteColResult = 8
    bteColUnit = 9
    
    With grid
        .clear
        .ColS = 10
        .Rows = 1
        
        .TextMatrix(0, bteColProd) = "Product Code"
        .TextMatrix(0, bteColDesc) = "Description"
        .TextMatrix(0, bteColPartNo) = "Parts No."
        .TextMatrix(0, bteColWH) = "W/H"
        .TextMatrix(0, bteColAddress) = "Address"
        .TextMatrix(0, bteColReqDate) = "Requirement Date"
        .TextMatrix(0, bteColPlan) = "Plan (Req)"
        .TextMatrix(0, bteColOffQty) = "Off Qty"
        .TextMatrix(0, bteColResult) = "Result"
        .TextMatrix(0, bteColUnit) = "Unit"
        
        .ColWidth(bteColProd) = 2500 'besar kolom
        .ColWidth(bteColDesc) = 2750
        .ColWidth(bteColPartNo) = 2500
        .ColWidth(bteColWH) = 1250
        .ColWidth(bteColAddress) = 1250
        .ColWidth(bteColReqDate) = 1650
        .ColWidth(bteColPlan) = 1500
        .ColWidth(bteColOffQty) = 1500
        .ColWidth(bteColResult) = 1500
        .ColWidth(bteColUnit) = 600
        
        .ColAlignment(bteColProd) = flexAlignLeftCenter
        .ColAlignment(bteColDesc) = flexAlignLeftCenter
        .ColAlignment(bteColPartNo) = flexAlignLeftCenter
        .ColAlignment(bteColWH) = flexAlignLeftCenter
        .ColAlignment(bteColAddress) = flexAlignLeftCenter
        .ColAlignment(bteColReqDate) = flexAlignCenterCenter
        .ColAlignment(bteColPlan) = flexAlignRightCenter
        .ColAlignment(bteColOffQty) = flexAlignRightCenter
        .ColAlignment(bteColResult) = flexAlignRightCenter
        .ColAlignment(bteColUnit) = flexAlignLeftCenter
    End With

End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
    CtrlMenu1.FormName = Me.Name
End Sub

Sub IsiGrid(parent As String, lotno As String, factoryCD As String)
Dim rsRequirement As New ADODB.Recordset

With grid
    Call headerGrid
    
    '**** Utk DO Master
    sql = "select ChildItem_Code, b.Item_Name,b.MakerItem_Code," & _
            "b.Wh_Code,b.Address, ChildRequirement_Date,ChildRequirement_QTy,OffChildRequirement_Qty," & _
            "ChildREquirementResult_Qty,ChildUnit_Cls " & _
        "from Requirement a,Item_Master b " & _
        "where a.ChildItem_Code = b.Item_Code " & _
            "And Factory_Code = '" & factoryCD & "' " & _
            "And ParentItem_Code = '" & parent & "' and Lot_No = '" & lotno & "' " & _
            "And a.Production_Date = '" & Format(lblNm(4), "yyyy-MM-dd") & "' " & _
        "Order by ChildITem_Code"
    Set rsRequirement = Db.Execute(sql)
    
    If Not (rsRequirement.EOF) Then
        LblErrMsg = ""
        i = 1
        Do While Not rsRequirement.EOF
            .Rows = .Rows + 1
            .TextMatrix(i, bteColProd) = Trim(rsRequirement("ChildItem_Code"))
            .TextMatrix(i, bteColDesc) = Trim(rsRequirement("Item_Name"))
            .TextMatrix(i, bteColPartNo) = Trim(rsRequirement("MakerITem_Code"))
            .TextMatrix(i, bteColWH) = Trim(rsRequirement("Wh_Code"))
            .TextMatrix(i, bteColAddress) = IIf(IsNull(Trim(rsRequirement("Address"))), "", Trim(rsRequirement("Address")))
            .TextMatrix(i, bteColReqDate) = Format(Trim(rsRequirement("ChildRequirement_Date")), "dd MMM yyyy")
            .TextMatrix(i, bteColPlan) = Format(rsRequirement("ChildRequirement_Qty"), gs_formatQty)
            .TextMatrix(i, bteColOffQty) = Format(rsRequirement("OffChildRequirement_Qty"), gs_formatQty)
            .TextMatrix(i, bteColResult) = Format(rsRequirement("ChildRequirementResult_Qty"), gs_formatQty)
            .TextMatrix(i, bteColUnit) = uf_GetUnitDescription(rsRequirement("ChildUnit_Cls"))
            i = i + 1
            rsRequirement.MoveNext
        Loop
    Else
        LblErrMsg = DisplayMsg(4006)
    End If
    Set rsRequirement = Nothing
End With
End Sub

Private Sub CmdSubMenu_Click()
    DoEvents
    frmMRPInquiry.Show
    DoEvents
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
    Unload frmMRPInquiry
Else
    LblErrMsg.Caption = ErrMsg
End If
End Sub
