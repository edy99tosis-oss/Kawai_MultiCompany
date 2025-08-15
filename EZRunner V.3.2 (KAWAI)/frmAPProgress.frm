VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAPProgress 
   BackColor       =   &H00FDDFE3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AP Progress Control"
   ClientHeight    =   11040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   Icon            =   "frmAPProgress.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   11040
   ScaleWidth      =   15270
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   435
      Left            =   13200
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   120
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   767
   End
   Begin VSFlex8Ctl.VSFlexGrid grid2 
      Height          =   1200
      Left            =   6975
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   8220
      Width           =   8070
      _cx             =   14235
      _cy             =   2117
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   10932991
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16637923
      BackColorAlternate=   16777215
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
      GridLinesFixed  =   1
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
   Begin VB.CommandButton command1 
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
      Index           =   3
      Left            =   11550
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   10080
      Width           =   1125
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FDDFE3&
      Height          =   533
      Left            =   180
      TabIndex        =   12
      Top             =   9450
      Width           =   14895
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
         Left            =   120
         TabIndex        =   13
         Top             =   180
         Width           =   14625
      End
   End
   Begin VB.CommandButton command1 
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
      Index           =   1
      Left            =   13965
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   10080
      Width           =   1125
   End
   Begin VB.CommandButton cmdsubmenu 
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
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   10080
      Width           =   1125
   End
   Begin VB.CommandButton command1 
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
      Index           =   2
      Left            =   12765
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   10080
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1215
      Left            =   210
      TabIndex        =   14
      Top             =   780
      Width           =   14835
      Begin VB.CommandButton command1 
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
         Index           =   0
         Left            =   7350
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   675
         Width           =   1125
      End
      Begin VB.TextBox lblcust 
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
         Height          =   210
         Left            =   3825
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   315
         Width           =   4995
      End
      Begin MSComCtl2.DTPicker invdate1 
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Top             =   705
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
         Format          =   141230083
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker invdate2 
         Height          =   315
         Left            =   3750
         TabIndex        =   2
         Top             =   705
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
         Format          =   141230083
         CurrentDate     =   37798
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Closed"
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
         Left            =   5520
         TabIndex        =   19
         Top             =   765
         Width           =   585
      End
      Begin MSForms.ComboBox combo1 
         Height          =   315
         Left            =   6270
         TabIndex        =   3
         Top             =   705
         Width           =   975
         VariousPropertyBits=   746604571
         DisplayStyle    =   7
         Size            =   "1720;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label3 
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
         Index           =   4
         Left            =   3300
         TabIndex        =   18
         Top             =   765
         Width           =   165
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice Date"
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
         Left            =   120
         TabIndex        =   16
         Top             =   765
         Width           =   1095
      End
      Begin MSForms.ComboBox cbocust 
         Height          =   315
         Left            =   1560
         TabIndex        =   0
         Top             =   285
         Width           =   2040
         VariousPropertyBits=   612386843
         MaxLength       =   10
         DisplayStyle    =   3
         Size            =   "3598;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   3825
         X2              =   8865
         Y1              =   570
         Y2              =   570
      End
      Begin VB.Label LblCode 
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
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   330
         Width           =   1215
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid grid 
      Height          =   6015
      Left            =   210
      TabIndex        =   5
      Top             =   2100
      Width           =   14835
      _cx             =   26167
      _cy             =   10610
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
      FixedRows       =   2
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
      Begin MSComCtl2.DTPicker paiddate 
         Height          =   315
         Left            =   0
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
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
         Format          =   141230083
         CurrentDate     =   37798
      End
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "AP Progress Control"
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
      Left            =   5640
      TabIndex        =   11
      Top             =   225
      Width           =   3960
   End
End
Attribute VB_Name = "frmAPProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sql As String
Dim RS As New ADODB.Recordset
Dim i As Long, j As Integer, k As Boolean

Dim Curr(6) As String, InvoiceAmount(6) As Double, CollectionAmount(6) As Double
Dim RemainingAmount(6) As Double, RemainingPPN As Double
Dim InvoicePPN As Double, CollectionPPN As Double

Dim bteColInvDate As Byte
Dim bteColInvNo As Byte
Dim bteColCurr As Byte
Dim bteColCurrCode As Byte
Dim bteColInvAmount As Byte
Dim bteColInvPPn As Byte
Dim bteColPayAmount As Byte
Dim bteColPayPPn As Byte
Dim bteColRemAmount As Byte
Dim bteColRemPPn As Byte
Dim bteColDueDate As Byte
Dim bteColClosed As Byte
Dim bteColDate As Byte

Dim bteColTotCurr As Byte
Dim bteColTotInv As Byte
Dim bteColTotPay As Byte
Dim bteColTotRem As Byte

Sub Header()
    
    bteColInvDate = 0
    bteColInvNo = 1
    bteColCurr = 2
    bteColCurrCode = 3
    bteColInvAmount = 4
    bteColInvPPn = 5
    bteColPayAmount = 6
    bteColPayPPn = 7
    bteColRemAmount = 8
    bteColRemPPn = 9
    bteColDueDate = 10
    bteColClosed = 11
    bteColDate = 12
    
    With grid
        
        .clear
        .Rows = 2
        .ColS = 13
        
        .TextMatrix(0, bteColInvDate) = "Invoice Date"
        .TextMatrix(0, bteColInvNo) = "Invoice No"
        .TextMatrix(0, bteColCurr) = "CurrCls"
        .TextMatrix(0, bteColCurrCode) = "Curr"
        .TextMatrix(0, bteColInvAmount) = "Invoice"
        .TextMatrix(0, bteColInvPPn) = "Invoice"
        .TextMatrix(0, bteColPayAmount) = "Payment"
        .TextMatrix(0, bteColPayPPn) = "Payment"
        .TextMatrix(0, bteColRemAmount) = "Remaining"
        .TextMatrix(0, bteColRemPPn) = "Remaining"
        .TextMatrix(0, bteColDueDate) = "Due Date"
        .TextMatrix(0, bteColClosed) = "Closed"
        .TextMatrix(0, bteColDate) = "Closed"
        
        .TextMatrix(1, bteColInvDate) = "Invoice Date"
        .TextMatrix(1, bteColInvNo) = "Invoice No"
        .TextMatrix(1, bteColCurr) = "CurrCls"
        .TextMatrix(1, bteColCurrCode) = "Curr"
        .TextMatrix(1, bteColInvAmount) = "Amount"
        .TextMatrix(1, bteColInvPPn) = "PPN (IDR)"
        .TextMatrix(1, bteColPayAmount) = "Amount"
        .TextMatrix(1, bteColPayPPn) = "PPN (IDR)"
        .TextMatrix(1, bteColRemAmount) = "Amount"
        .TextMatrix(1, bteColRemPPn) = "PPN (IDR)"
        .TextMatrix(1, bteColDueDate) = "Due Date"
        .TextMatrix(1, bteColClosed) = "Closed"
        .TextMatrix(1, bteColDate) = "Date"
        
        .ColWidth(bteColInvDate) = 1250
        .ColWidth(bteColInvNo) = 2150
        .ColWidth(bteColCurrCode) = 500
        .ColWidth(bteColInvAmount) = 1250
        .ColWidth(bteColInvPPn) = 1250
        .ColWidth(bteColPayAmount) = 1250
        .ColWidth(bteColPayPPn) = 1250
        .ColWidth(bteColRemAmount) = 1250
        .ColWidth(bteColRemPPn) = 1250
        .ColWidth(bteColDueDate) = 1250
        .ColWidth(bteColClosed) = 750
        .ColWidth(bteColDate) = 1400
        
        '.ColHidden(bteColCurr) = True
        
        .MergeRow(bteColInvDate) = True
        .MergeRow(bteColInvNo) = True
        For i = 0 To 11
            .MergeCol(i) = True
        Next i
        .MergeCells = flexMergeFixedOnly
        
        .Cell(flexcpAlignment, 0, 0, 0, 10) = flexAlignCenterCenter
        .ColAlignment(bteColInvDate) = flexAlignCenterCenter
        .ColAlignment(bteColInvNo) = flexAlignLeftCenter
        .ColAlignment(bteColCurr) = flexAlignCenterCenter
        .ColAlignment(bteColCurrCode) = flexAlignRightCenter
        .ColAlignment(bteColInvAmount) = flexAlignRightCenter
        .ColAlignment(bteColInvPPn) = flexAlignRightCenter
        .ColAlignment(bteColPayPPn) = flexAlignRightCenter
        .ColAlignment(bteColRemAmount) = flexAlignRightCenter
        .ColAlignment(bteColRemPPn) = flexAlignRightCenter
        .ColAlignment(bteColDueDate) = flexAlignCenterCenter
        
    End With
    
    Dim ClsProc As New ClsProc
    Call ClsProc.AlignHeader(grid, 0)
    Call ClsProc.AlignHeader(grid, 1)
    Call ClsProc.AlignHeader(Grid2, 0)
    Call headerTotal
    paiddate.Visible = False
End Sub

Sub headerTotal()

    bteColTotCurr = 0
    bteColTotInv = 1
    bteColTotPay = 2
    bteColTotRem = 3
    
    With Grid2
        .clear
        .Rows = 1
        .ColS = 4
        
        .TextMatrix(0, bteColTotCurr) = "Curr"
        .TextMatrix(0, bteColTotInv) = "Grand Total Invoice"
        .TextMatrix(0, bteColTotPay) = "Grand Total Payment"
        .TextMatrix(0, bteColTotRem) = "Grand Total Remaining"
        
        .ColAlignment(bteColTotCurr) = flexAlignLeftCenter
        .ColAlignment(bteColTotInv) = flexAlignRightCenter
        .ColAlignment(bteColTotPay) = flexAlignRightCenter
        .ColAlignment(bteColTotRem) = flexAlignRightCenter
        
        .Cell(flexcpAlignment, 0, 0, 0, 3) = flexAlignCenterCenter
        
        .ColWidth(bteColTotCurr) = 1000
        .ColWidth(bteColTotInv) = 2200
        .ColWidth(bteColTotPay) = 2200
        .ColWidth(bteColTotRem) = 2200
    End With
End Sub

Sub Kosong()
    lblcust.Text = ""
    cboCust.Text = ""
    invdate1.Value = Format(Now, "yyyy-mm-01")
    invdate2.Value = Format(Now, "dd MMM yyyy")
    paiddate.Value = Format(Now, "dd MMM yyyy")
    combo1.ListIndex = 0
    LblErrMsg = ""
    Header
End Sub

Sub adtocboCust()
Dim sqlcust As String
Dim RsCust As New Recordset

    sqlcust = "select trade_code, trade_name, country_cls " & _
              "from trade_master where (trade_cls='2' ) or trade_cls='3' Order By Trade_Code"
              
    Set RsCust = Db.Execute(sqlcust)

    With cboCust
        .clear
        .columnCount = 2
        .ColumnWidths = "80pt;300pt"
        .ListWidth = 380
        .ListRows = 15

        i = 0
        Do While Not RsCust.EOF
            .AddItem
            .List(i, 0) = Trim(RsCust("Trade_code"))
            .List(i, 1) = IIf(IsNull(RsCust("trade_name")), " ", Trim(RsCust("Trade_Name")))
            
            RsCust.MoveNext
            i = i + 1

        Loop
    End With
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
    adtocboCust
    combo1.AddItem "All"
    combo1.AddItem "Yes"
    combo1.AddItem "No"

    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"

    Kosong
End Sub

Sub Browse()
Dim sqlpaid As String
    
    LblErrMsg = ""
    Header
    Call resetArr
    sqlpaid = ""
    If combo1.Text = "Yes" Then
        sqlpaid = " and isnull(im.paid_cls,0)='1' "
    ElseIf combo1.Text = "No" Then
        sqlpaid = " and isnull(im.paid_cls,0)='0' "
    End If
    


    sql = "select distinct im.supplier_code, im.invoice_no, im.invoice_Date, im.due_date, id.currency_code, " & _
            "total_amount = isnull(im.total_amount,0) + isnull(im.airfreight_amount,0) , " & _
            "totalapamount = isnull((Select sum(amount) from aP_detail where invoice_no= im.invoice_no and supplier_code=im.supplier_code),0) , " & _
            "paid_cls = isnull(paid_cls,0) , im.paid_Date, " & _
            "totPPN = ISNULL(IM.PPN,0) , " & _
            "totAPAmount = ISNULL((Select SUM(Amount) From AP_Detail Where Invoice_No = IM.Invoice_No),0), " & _
            "totAPPPN = ISNULL((Select SUM(PPN) From AP_Detail Where Invoice_No = IM.Invoice_No),0) " & _
          "from invoicesupplier_master im " & _
          "inner join invoicesupplier_detail id on id.invoice_no=im.invoice_no and id.supplier_code=im.supplier_code " & _
          "  left join trade_master tm on im.supplier_code=tm.trade_code " & _
          "where im.supplier_Code= '" & Trim(cboCust.Text) & "' and im.invoice_date >='" & Format(invdate1.Value, "yyyy-mm-dd") & _
          "' and im.invoice_date <= '" & Format(invdate2.Value, "yyyy-mm-dd") & "'  " & sqlpaid & _
          "order by im.invoice_no "

    If RS.State <> adStateClosed Then RS.Close
    RS.Open sql, Db, adOpenKeyset, adLockOptimistic
    If RS.EOF Then LblErrMsg.Caption = DisplayMsg(4006): Exit Sub

    With grid
        If Not (RS.BOF And RS.EOF) Then
            i = 2
        
            Do While Not RS.EOF
                .Rows = .Rows + 1
        
                .TextMatrix(i, bteColInvDate) = Format(RS("invoice_date"), "dd MMM yyyy")
                .TextMatrix(i, bteColInvNo) = IIf(IsNull(RS("invoice_no")), "", Trim(RS("invoice_no")))
                .TextMatrix(i, bteColCurr) = Trim(RS("currency_code"))
                .TextMatrix(i, bteColCurrCode) = uf_GetCurrencyDescription(Trim(RS("currency_code")))
                .TextMatrix(i, bteColInvAmount) = Format(RS("total_amount"), gs_formatAmount)
                .TextMatrix(i, bteColInvPPn) = Format(RS("totPPN"), gs_formatAmount)
                .TextMatrix(i, bteColPayAmount) = Format(RS("totAPAmount"), gs_formatAmount)
                .TextMatrix(i, bteColPayPPn) = Format(RS("totAPPPN"), gs_formatAmount)
                .TextMatrix(i, bteColRemAmount) = Format((RS("total_amount") - RS("totAPAmount")), gs_formatAmount)
                .TextMatrix(i, bteColRemPPn) = Format((RS("totPPN") - RS("totAPPPN")), gs_formatAmount)
                If IsNull(RS("due_date")) Then
                    .TextMatrix(i, bteColDueDate) = ""
                Else
                    .TextMatrix(i, bteColDueDate) = Format(RS("due_date"), "dd MMM yyyy")
                End If
        
                If RS("paid_cls") = 1 Then
                    .Cell(flexcpChecked, i, bteColClosed) = flexChecked
                    .TextMatrix(i, bteColRemAmount) = Format(0, gs_formatAmount)
                    .TextMatrix(i, bteColRemPPn) = Format(0, gs_formatAmount)
                    If IsNull(RS("paid_date")) Then
                        .TextMatrix(i, bteColDate) = ""
                    Else
                        .TextMatrix(i, bteColDate) = Format(RS("paid_date"), "dd MMM yyyy")
                    End If
                Else
                    .Cell(flexcpChecked, i, bteColClosed) = flexUnchecked
                    .TextMatrix(i, bteColDate) = ""
                End If
                .Cell(flexcpBackColor, i, bteColClosed) = vbWhite
                .Cell(flexcpBackColor, i, bteColDate) = vbWhite
                      
                '**************** Itung Total ****************
                If Curr(Val(.TextMatrix(i, bteColCurr))) = "" Then Curr(Val(.TextMatrix(i, bteColCurr))) = Trim(.TextMatrix(i, bteColCurrCode))
                InvoiceAmount(Val(.TextMatrix(i, bteColCurr))) = InvoiceAmount(Val(.TextMatrix(i, bteColCurr))) + CDbl(.TextMatrix(i, bteColInvAmount))
                InvoicePPN = InvoicePPN + CDbl(.TextMatrix(i, bteColInvPPn))
                                
                CollectionAmount(Val(.TextMatrix(i, bteColCurr))) = CollectionAmount(Val(.TextMatrix(i, bteColCurr))) + CDbl(.TextMatrix(i, bteColPayAmount))
                CollectionPPN = CollectionPPN + CDbl(.TextMatrix(i, bteColPayPPn))
                
                RemainingAmount(Val(.TextMatrix(i, bteColCurr))) = RemainingAmount(Val(.TextMatrix(i, bteColCurr))) + CDbl(.TextMatrix(i, bteColRemAmount))
                RemainingPPN = RemainingPPN + CDbl(.TextMatrix(i, bteColRemPPn))
                '************************************************
        
                RS.MoveNext
                i = i + 1
            Loop
        End If
    End With
    If RS.State <> adStateClosed Then RS.Close
    
    Call isiGridTotal
End Sub

Sub resetArr()
Dim j As Integer
    For j = 1 To 6
        Curr(j) = ""
        InvoiceAmount(j) = 0
        CollectionAmount(j) = 0
        RemainingAmount(j) = 0
    Next j
    InvoicePPN = 0
    CollectionPPN = 0
    RemainingPPN = 0
End Sub

Sub isiGridTotal()
With Grid2
    For i = 1 To 6
        If Curr(i) <> "" Then
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, bteColTotCurr) = Curr(i)
            .TextMatrix(.Rows - 1, bteColTotInv) = Format(InvoiceAmount(i), gs_formatAmount)
            .TextMatrix(.Rows - 1, bteColTotPay) = Format(CollectionAmount(i), gs_formatAmount)
            .TextMatrix(.Rows - 1, bteColTotRem) = Format(RemainingAmount(i), gs_formatAmount)
        End If
    Next i
    
    If InvoicePPN > 0 Then
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, bteColTotCurr) = "PPN (IDR)"
        .TextMatrix(.Rows - 1, bteColTotInv) = Format(InvoicePPN, gs_formatAmount)
        .TextMatrix(.Rows - 1, bteColTotPay) = Format(CollectionPPN, gs_formatAmount)
        .TextMatrix(.Rows - 1, bteColTotRem) = Format(RemainingPPN, gs_formatAmount)
    End If
End With
End Sub

Sub itungTot()
Call resetArr
Call headerTotal

With grid
    For i = 2 To .Rows - 1
        If Curr(Val(.TextMatrix(i, bteColCurr))) = "" Then Curr(Val(.TextMatrix(i, bteColCurr))) = Trim(.TextMatrix(i, bteColCurrCode))
        InvoiceAmount(Val(.TextMatrix(i, bteColCurr))) = InvoiceAmount(Val(.TextMatrix(i, bteColCurr))) + CDbl(.TextMatrix(i, bteColInvAmount))
        InvoicePPN = InvoicePPN + CDbl(.TextMatrix(i, bteColInvPPn))
                        
        CollectionAmount(Val(.TextMatrix(i, bteColCurr))) = CollectionAmount(Val(.TextMatrix(i, bteColCurr))) + CDbl(.TextMatrix(i, bteColPayAmount))
        CollectionPPN = CollectionPPN + CDbl(.TextMatrix(i, bteColPayPPn))
        
        RemainingAmount(Val(.TextMatrix(i, bteColCurr))) = RemainingAmount(Val(.TextMatrix(i, bteColCurr))) + CDbl(.TextMatrix(i, bteColRemAmount))
        RemainingPPN = RemainingPPN + CDbl(.TextMatrix(i, bteColRemPPn))
    Next i
End With

Call isiGridTotal
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With grid
    If Col = bteColClosed Then
        If .Cell(flexcpChecked, Row, bteColClosed) = flexChecked Then
            .TextMatrix(Row, bteColRemAmount) = Format(0, gs_formatAmount)
            .TextMatrix(Row, bteColRemPPn) = Format(0, gs_formatAmount)
        Else
            .TextMatrix(Row, bteColRemAmount) = Format((CDbl(.TextMatrix(Row, bteColInvAmount)) - CDbl(.TextMatrix(Row, bteColPayAmount))), gs_formatAmount)
            .TextMatrix(Row, bteColRemPPn) = Format((CDbl(.TextMatrix(Row, bteColInvPPn)) - CDbl(.TextMatrix(Row, bteColPayPPn))), gs_formatAmount)
        End If
        Call itungTot
    End If
End With
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If grid.Col <> bteColClosed Then Cancel = True
End Sub

Private Sub grid_Click()
    If grid.Row < 0 Then Exit Sub
    LblErrMsg.Caption = ""

    With grid
        If .Col = bteColClosed Or .Col = bteColDate Then
            If .TextMatrix(.Row, bteColDate) <> "" Then
                paiddate.Value = Format(.TextMatrix(.Row, bteColDate), "dd MMM yyyy")
            Else
                paiddate.Value = Format(Now, " dd mmm yyyy")
            End If

            If .Cell(flexcpChecked, .Row, bteColClosed) = flexChecked Then
                .TextMatrix(.Row, bteColDate) = Format(paiddate, "dd MMM YYYY")
                paiddate.Visible = True
                paiddate.Left = .Cell(flexcpLeft, .Row, bteColDate)
                paiddate.top = .Cell(flexcpTop, .Row, bteColDate)
                paiddate.Width = .Cell(flexcpWidth, .Row, bteColDate)
                paiddate.SetFocus
            Else
                .TextMatrix(.Row, bteColDate) = ""
                paiddate.Visible = False
            End If
        End If
    End With
End Sub

Private Sub Grid_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    paiddate.Visible = False
End Sub

Private Sub grid_LeaveCell()
    paiddate.Visible = False
End Sub

Private Sub PaidDate_Change()
With grid
    .TextMatrix(grid.Row, bteColDate) = Format(paiddate, "dd MMM YYYY")
    paiddate.Visible = False
End With
End Sub

Private Sub Combo1_Click()
    LblErrMsg = ""
    Header
End Sub

Private Sub combo1_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
  If KeyCode = 13 Then Combo1_Click
End Sub

Private Sub cboCust_Click()
LblErrMsg = ""

    If cboCust.ListIndex <> -1 Then
        lblcust.Text = cboCust.Column(1)
        
        Header
    Else
        Kosong
        LblErrMsg.Caption = DisplayMsg(4050) '"Record with this Supplier Code not Exist"
        cboCust.SetFocus
        Exit Sub
    End If
    
End Sub

Private Sub cboCust_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then cboCust_Click
End Sub

Private Sub cbocust_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub invdate1_Change()
   If CDate(invdate1) > CDate(invdate2) Then
      LblErrMsg.Caption = DisplayMsg(4029) & " " & Format(invdate2, "dd MMM yyyy") '"Invoice Date must be lower than "
      Exit Sub
   Else
      LblErrMsg.Caption = ""
   End If
    
    Header
End Sub

Private Sub invdate2_Change()
   If CDate(invdate2) < CDate(invdate1) Then
      LblErrMsg.Caption = DisplayMsg(4030) & " " & Format(invdate1, "dd MMM yyyy") '"invoice Date must be higher than "
      Exit Sub
   Else
      LblErrMsg.Caption = ""
   End If
    
    Header
End Sub

Private Sub Command1_Click(Index As Integer)
Dim db1 As New Connection
paiddate.Visible = False
LblErrMsg = ""
db1.ConnectionString = Db.ConnectionString

Select Case Index
  Case 0:
            If cboCust.Text = "" Then
              cboCust.SetFocus
              LblErrMsg = DisplayMsg(1064) '"Please Select Supplier Code"
              Exit Sub
            End If
        
                If cboCust.Text <> "" Then
                  cboCust.MatchEntry = 1
                  cboCust.Text = cboCust.Text
                  If cboCust.MatchFound = False Then
                      LblErrMsg = DisplayMsg(4050)
                      cboCust.SetFocus
                      cboCust.MatchEntry = 2
                      Exit Sub
                  End If
                  cboCust.MatchEntry = 2
                End If
                cboCust.MatchEntry = 1
            
            Browse
            
  Case 1:   If hakUpdate(Me.Name) = 0 Then LblErrMsg = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub

            If cboCust.Text = "" Then
              cboCust.SetFocus
              LblErrMsg = DisplayMsg(1064) '"Please Select Supplier Code"
              Exit Sub
            End If
        
                If cboCust.Text <> "" Then
                  cboCust.MatchEntry = 1
                  cboCust.Text = cboCust.Text
                  If cboCust.MatchFound = False Then
                      LblErrMsg = DisplayMsg(4050)
                      cboCust.SetFocus
                      cboCust.MatchEntry = 2
                      Exit Sub
                  End If
                  cboCust.MatchEntry = 2
                End If
                cboCust.MatchEntry = 1
                
                                     
            If grid.Rows = 1 Then
                LblErrMsg.Caption = DisplayMsg(5012)
                Command1(0).SetFocus
                Exit Sub
            End If
            
            db1.Open
            db1.BeginTrans
            
            With grid
                For i = 1 To .Rows - 1
                    If .Cell(flexcpChecked, i, bteColClosed) = flexChecked Then
                        sql = "update invoicesupplier_master set paid_cls='1', paid_date='" & Format(.TextMatrix(i, bteColDate), "yyyy-mm-dd") & "', " & _
                                "Last_Update = getdate(), Last_User = '" & userLogin & "' " & _
                                "where invoice_no='" & Trim(.TextMatrix(i, bteColInvNo)) & "' " & _
                                "and supplier_code='" & Trim(cboCust.Text) & "' "
                    Else
                        sql = "update invoicesupplier_master set paid_cls='0', paid_date=Null, " & _
                                "Last_Update = getdate(), Last_User = '" & userLogin & "' " & _
                                "where invoice_no='" & Trim(.TextMatrix(i, bteColInvNo)) & "' " & _
                                "and supplier_code='" & Trim(cboCust.Text) & "' "
                    End If
                    db1.Execute sql
                Next i
            End With
            
            If err.number = 0 Then
                db1.CommitTrans
            Else
                db1.RollbackTrans
                err.clear
                Set db1 = Nothing
                Exit Sub
            End If
            
            Browse
            LblErrMsg = DisplayMsg(1101)
                    
    Case 2: Kosong
            cboCust.SetFocus
    
    Case 3: If cboCust.Text <> "" Then Browse
            
End Select
Set db1 = Nothing
End Sub

Private Sub CmdSubMenu_Click()
    Unload Me
    frmMainMenu.Show
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    LblErrMsg.Caption = ErrMsg
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If RS.State <> adStateClosed Then RS.Close
End Sub

