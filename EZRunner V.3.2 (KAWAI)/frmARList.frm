VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmARList 
   BackColor       =   &H00FDDFE3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AR List"
   ClientHeight    =   10425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   Icon            =   "frmARList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10425
   ScaleWidth      =   15270
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   435
      Left            =   13230
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   150
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   767
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1155
      Left            =   180
      TabIndex        =   17
      Top             =   720
      Width           =   14895
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
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   275
         Width           =   4995
      End
      Begin MSComCtl2.DTPicker invdate1 
         Height          =   315
         Left            =   1560
         TabIndex        =   1
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
         Format          =   141230083
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker invdate2 
         Height          =   315
         Left            =   3600
         TabIndex        =   2
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
         Format          =   141230083
         CurrentDate     =   37798
      End
      Begin VB.Label LblCode 
         BackStyle       =   0  'Transparent
         Caption         =   "Cust CD"
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
         TabIndex        =   21
         Top             =   285
         Width           =   1215
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   3240
         X2              =   8280
         Y1              =   525
         Y2              =   525
      End
      Begin MSForms.ComboBox cbocust 
         Height          =   315
         Left            =   1560
         TabIndex        =   0
         Top             =   240
         Width           =   1500
         VariousPropertyBits=   612386843
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2646;556"
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
         TabIndex        =   20
         Top             =   750
         Width           =   1095
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
         Left            =   3240
         TabIndex        =   19
         Top             =   750
         Width           =   165
      End
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
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9360
      Width           =   1125
   End
   Begin VB.CommandButton command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Create"
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
      Left            =   8970
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1980
      Width           =   1125
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FDDFE3&
      Height          =   533
      Left            =   165
      TabIndex        =   13
      Top             =   8745
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
         TabIndex        =   14
         Top             =   180
         Width           =   14685
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
      Left            =   13935
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9360
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
      TabIndex        =   11
      Top             =   9360
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
      Left            =   12735
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   9360
      Width           =   1125
   End
   Begin MSComCtl2.DTPicker apdate 
      Height          =   315
      Left            =   7320
      TabIndex        =   5
      Top             =   2010
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
   Begin VSFlex8Ctl.VSFlexGrid grid 
      Height          =   4935
      Left            =   180
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2460
      Width           =   14895
      _cx             =   26273
      _cy             =   8705
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
   End
   Begin VSFlex8Ctl.VSFlexGrid grid2 
      Height          =   1200
      Left            =   4935
      TabIndex        =   22
      Top             =   7500
      Width           =   10110
      _cx             =   17833
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
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Collection Voucher No"
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
      Left            =   1740
      TabIndex        =   16
      Top             =   2040
      Width           =   2040
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher Date"
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
      Index           =   2
      Left            =   6000
      TabIndex        =   15
      Top             =   2040
      Width           =   1305
   End
   Begin MSForms.ComboBox cboapno 
      Height          =   315
      Left            =   3780
      TabIndex        =   4
      Top             =   2010
      Width           =   2025
      VariousPropertyBits=   612386843
      DisplayStyle    =   3
      Size            =   "3572;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox combo1 
      Height          =   315
      Left            =   180
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2010
      Width           =   1215
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "2143;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "AR List"
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
      Left            =   5647
      TabIndex        =   12
      Top             =   225
      Width           =   3960
   End
End
Attribute VB_Name = "frmARList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sql As String, sqlGrid As String
Dim RS As New ADODB.Recordset
Dim rsGrid As New ADODB.Recordset
Dim i As Long, ubah As Boolean, ada As Boolean, j As Integer, k As Boolean

Dim Curr(6) As String, InvoiceAmount(6) As Double, CollectionAmount(6) As Double
Dim GrandTotalCollection(6) As Double, GrandTotalCollectionPPN As Double
Dim InvoicePPN As Double, CollectionPPN As Double
Dim RemainingAmount(6) As Double, RemainingPPN As Double

Dim bteColSelect As Byte
Dim bteColInvDate As Byte
Dim bteColInvNo As Byte
Dim bteColCurrCls As Byte
Dim bteColCurr As Byte
Dim bteColInvAmount As Byte
Dim bteColInvPPn As Byte
Dim bteColInvTotal As Byte
Dim bteColCollectAmount As Byte
Dim bteColCollectPPn As Byte
Dim bteColRemAmount As Byte
Dim bteColRemPPn As Byte
Dim bteColDueDate As Byte
Dim bteColARAmount As Byte
Dim bteColARPPn As Byte
Dim bteColARRemAmount As Byte
Dim bteColARRemPPn As Byte
Dim bteColPaidCls As Byte
Dim bteColARTotAmount As Byte
Dim bteColARTotPPn As Byte

Dim bteColCurr2 As Byte
Dim bteColCollect2 As Byte
Dim bteColTotInv As Byte
Dim bteColTotCollect As Byte
Dim bteColTotRem As Byte

Sub Header()
    
    bteColSelect = 0
    bteColInvDate = 1
    bteColInvNo = 2
    bteColCurrCls = 3
    bteColCurr = 4
    bteColInvAmount = 5
    bteColInvPPn = 6
    bteColInvTotal = 7
    bteColCollectAmount = 8
    bteColCollectPPn = 9
    bteColRemAmount = 10
    bteColRemPPn = 11
    bteColDueDate = 12
    bteColARAmount = 13
    bteColARPPn = 14
    bteColARRemAmount = 15
    bteColARRemPPn = 16
    bteColPaidCls = 17
    bteColARTotAmount = 18
    bteColARTotPPn = 19
    
    With grid
        .clear
        
        .Rows = 2
        .ColS = 20
        
        .TextMatrix(0, bteColSelect) = ""
        .TextMatrix(0, bteColInvDate) = "Invoice Date"
        .TextMatrix(0, bteColInvNo) = "Invoice No"
        .TextMatrix(0, bteColCurrCls) = "CurrCls"
        .TextMatrix(0, bteColCurr) = "Curr"
        .TextMatrix(0, bteColInvAmount) = "Invoice"
        .TextMatrix(0, bteColInvPPn) = "Invoice"
        .TextMatrix(0, bteColInvTotal) = "Invoice"
        .TextMatrix(0, bteColCollectAmount) = "Collection"
        .TextMatrix(0, bteColCollectPPn) = "Collection"
        .TextMatrix(0, bteColRemAmount) = "Remaining"
        .TextMatrix(0, bteColRemPPn) = "Remaining"
        .TextMatrix(0, bteColDueDate) = "Due Date"
        
        .TextMatrix(1, bteColSelect) = ""
        .TextMatrix(1, bteColInvDate) = "Invoice Date"
        .TextMatrix(1, bteColInvNo) = "Invoice No"
        .TextMatrix(1, bteColCurrCls) = "CurrCls"
        .TextMatrix(1, bteColCurr) = "Curr"
        .TextMatrix(1, bteColInvAmount) = "Amount"
        .TextMatrix(1, bteColInvPPn) = "PPN (IDR)"
        .TextMatrix(1, bteColInvTotal) = "Total Amount"
        .TextMatrix(1, bteColCollectAmount) = "Amount"
        .TextMatrix(1, bteColCollectPPn) = "PPN (IDR)"
        .TextMatrix(1, bteColRemAmount) = "Amount"
        .TextMatrix(1, bteColRemPPn) = "PPN (IDR)"
        .TextMatrix(1, bteColDueDate) = "Due Date"
        
        .ColWidth(bteColSelect) = 300
        .ColWidth(bteColInvDate) = 1365
        .ColWidth(bteColInvNo) = 2220
        .ColWidth(bteColCurr) = 500
        .ColWidth(bteColInvAmount) = 1500
        .ColWidth(bteColInvPPn) = 1200
        .ColWidth(bteColInvTotal) = 1500
        .ColWidth(bteColCollectAmount) = 1700
        .ColWidth(bteColCollectPPn) = 1400
        .ColWidth(bteColRemAmount) = 1700
        .ColWidth(bteColRemPPn) = 1400
        .ColWidth(bteColDueDate) = 1335
        
        .ColHidden(bteColCurrCls) = True
        .ColHidden(bteColARAmount) = True
        .ColHidden(bteColARPPn) = True
        .ColHidden(bteColARRemAmount) = True
        .ColHidden(bteColARRemPPn) = True
        .ColHidden(bteColPaidCls) = True
        .ColHidden(bteColARTotAmount) = True
        .ColHidden(bteColARTotPPn) = True
                
        .MergeRow(0) = True
        .MergeRow(1) = True
        For i = 0 To bteColDueDate
            .MergeCol(i) = True
        Next i
        .MergeCells = flexMergeFixedOnly
        
        .Cell(flexcpAlignment, 0, 0, 1, bteColDueDate) = flexAlignCenterCenter
        .ColAlignment(bteColSelect) = flexAlignCenterCenter
        .ColAlignment(bteColInvDate) = flexAlignCenterCenter
        .ColAlignment(bteColInvNo) = flexAlignLeftCenter
        .ColAlignment(bteColCurr) = flexAlignCenterCenter
        .ColAlignment(bteColInvAmount) = flexAlignRightCenter
        .ColAlignment(bteColInvPPn) = flexAlignRightCenter
        .ColAlignment(bteColInvTotal) = flexAlignRightCenter
        .ColAlignment(bteColCollectAmount) = flexAlignRightCenter
        .ColAlignment(bteColCollectPPn) = flexAlignRightCenter
        .ColAlignment(bteColRemAmount) = flexAlignRightCenter
        .ColAlignment(bteColRemPPn) = flexAlignRightCenter
        .ColAlignment(bteColDueDate) = flexAlignCenterCenter
    End With
    headerTotal
End Sub

Sub headerTotal()
    
    bteColCurr2 = 0
    bteColCollect2 = 1
    bteColTotInv = 2
    bteColTotCollect = 3
    bteColTotRem = 4
    
    With Grid2
        .clear
        .Rows = 1
        .ColS = 5
        
        .TextMatrix(0, bteColCurr2) = "Curr"
        .TextMatrix(0, bteColCollect2) = "Total Collection"
        .TextMatrix(0, bteColTotInv) = "Grand Total Invoice"
        .TextMatrix(0, bteColTotCollect) = "Grand Total Collection"
        .TextMatrix(0, bteColTotRem) = "Grand Total Remaining"
        
        .ColAlignment(bteColCurr2) = flexAlignLeftCenter
        .ColAlignment(bteColCollect2) = flexAlignRightCenter
        .ColAlignment(bteColTotInv) = flexAlignRightCenter
        .ColAlignment(bteColTotCollect) = flexAlignRightCenter
        .ColAlignment(bteColTotRem) = flexAlignRightCenter
        
        .Cell(flexcpAlignment, 0, 0, 0, 3) = flexAlignCenterCenter
        
        .ColWidth(bteColCurr2) = 1000
        .ColWidth(bteColCollect2) = 2000
        .ColWidth(bteColTotInv) = 2200
        .ColWidth(bteColTotCollect) = 2200
        .ColWidth(bteColTotRem) = 2200
    End With
End Sub

Sub Kosong()
    lblcust.Text = ""
    cboCust.Text = ""
    invdate1.Value = Format(Now, "yyyy-mm-01")
    invdate2.Value = Format(Now, "dd MMM yyyy")
    cboapno.clear
    cboapno = ""
    apdate.Value = Format(Now, "dd MMM yyyy")
    ubah = False
    LblErrMsg = ""
    Header
End Sub

Sub adtocboCust()
Dim sqlcust As String
Dim RsCust As New Recordset

    sqlcust = "select trade_code, trade_name, country_cls " & _
              "from trade_master where (trade_cls='2') or trade_cls='2' Order By Trade_Code"
    Set RsCust = Db.Execute(sqlcust)

    With cboCust
        .clear
        .columnCount = 2
        .ColumnWidths = "50pt;300pt"
        .ListWidth = 350
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

Sub adtocboapno()
Dim sqlno As String
Dim rsno As New Recordset
    
    sqlno = "(select distinct ar_no from ar_master where ar_no not in (select ar_no from ar_detail) " & _
            "and cust_code='" & Trim(cboCust.Text) & "') " & _
            "Union " & _
            "(select distinct ar_no from ar_master where cust_code='" & Trim(cboCust.Text) & "' " & _
            "and ar_date>='" & Format(invdate1.Value, "yyyy-mm-dd") & "' " & _
            "and ar_date<='" & Format(invdate2.Value, "yyyy-mm-dd") & "') "
    Set rsno = Db.Execute(sqlno)

    With cboapno
        .clear
        .ColumnWidths = "150pt"
        .ListWidth = 150
        .ListRows = 15

        i = 0
        Do While Not rsno.EOF
            .AddItem Trim(rsno("ar_No"))
            rsno.MoveNext
            i = i + 1
        Loop
    End With

End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
    adtocboCust
    combo1.AddItem "Create"
    combo1.AddItem "Update"
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
    Kosong
    combo1.ListIndex = 1
End Sub

Sub browseitem()
Dim sqlbrowse As String
Dim rsbrowse As New Recordset

    sqlbrowse = "Select dt1.*, " & _
                    "remainingamount = amount - TotARAmount, " & _
                    "remainingppn = PPNIDR - TotARPPN " & _
                "From (" & _
                    "select distinct stAR = 0, ism.cust_code, ism.invoice_no, ism.invoice_Date, ism.due_Date, " & _
                        "isd.currency_code, ism.amount, " & _
                        "Paid_Cls = ISNULL(ISM.Paid_Cls,0), " & _
                        "PPNIDR = ISNULL((Select Top 1 PPN_IDR From FakturPajak_Master FPM, FakturPajak_Detail FPD " & _
                                "Where FPM.FakturPajak_No = FPD.FakturPajak_No " & _
                                    "And Invoice_No = ISM.Invoice_No),0), " & _
                        "aramount = isnull((select sum(amount) from ar_detail where cust_code=ism.cust_code and invoice_no=ism.invoice_no and ar_no='" & Trim(cboapno.Text) & "'),0), " & _
                        "arppn =  isnull((select sum(ppn) from ar_detail where cust_code=ism.cust_code and invoice_no=ism.invoice_no and ar_no='" & Trim(cboapno.Text) & "'),0), " & _
                        "TotARAmount = isnull((select sum(amount) from ar_detail where cust_code=ism.cust_code and invoice_no=ism.invoice_no),0), " & _
                        "TotARPPN = isnull((select sum(ppn) from ar_detail where cust_code=ism.cust_code and invoice_no=ism.invoice_no),0) " & _
                    "from invoice_master ism " & _
                        "inner join invoice_detail isd on isd.invoice_no=ism.invoice_no " & _
                        "INNER JOIN Trade_Master T ON T.Trade_Code = ISM.Cust_Code " & _
                    "where ism.cust_code='" & Trim(cboCust.Text) & "' and ism.invoice_date>='" & Format(invdate1.Value, "yyyy-mm-dd") & "' and " & _
                        "ism.invoice_date<='" & Format(invdate2.Value, "yyyy-mm-dd") & "' and ism.fix_cls='1' " & _
                        "and isnull(ism.paid_cls,0)<>'1' " & _
                        "and ism.invoice_no not in (select invoice_no from ar_detail where cust_code=ism.cust_code and ar_no='" & Trim(cboapno.Text) & "') " & _
                    ")dt1 " & _
                "Where (amount - TotARAmount > 0) OR (PPNIDR - TotARPPN > 0 ) "
    
    sqlbrowse = sqlbrowse & _
                "UNION Select dt2.*, " & _
                    "remainingamount = (Case Paid_Cls When 1 then 0 Else amount - TotARAmount End), " & _
                    "remainingppn = (Case Paid_Cls When 1 then 0 Else PPNIDR - TotARPPN End)" & _
                "From (" & _
                    "select distinct stAR = 1, ism.cust_code, ism.invoice_no, ism.invoice_Date, ism.due_Date, isd.currency_code, ism.amount, " & _
                        "Paid_Cls = ISNULL(ISM.Paid_Cls,0), " & _
                        "PPNIDR = ISNULL((Select Top 1 PPN_IDR From FakturPajak_Master FPM, FakturPajak_Detail FPD " & _
                                "Where FPM.FakturPajak_No = FPD.FakturPajak_No " & _
                                    "And Invoice_No = ISM.Invoice_No),0) " & _
                        ",aramount = isnull((select sum(amount) from ar_detail where cust_code=ism.cust_code and invoice_no=ism.invoice_no And AR_No = '" & Trim(cboapno.Text) & "'),0), " & _
                        "arppn = isnull((select sum(ppn) from ar_detail where cust_code=ism.cust_code and invoice_no=ism.invoice_no And AR_No = '" & Trim(cboapno.Text) & "'),0) , " & _
                        "TotARAmount = isnull((select sum(amount) from ar_detail where cust_code=ism.cust_code and invoice_no=ism.invoice_no),0), " & _
                        "TotARPPN = isnull((select sum(ppn) from ar_detail where cust_code=ism.cust_code and invoice_no=ism.invoice_no),0) " & _
                    "from invoice_master ism " & _
                        "inner join invoice_detail isd on isd.invoice_no=ism.invoice_no " & _
                        "INNER JOIN Trade_Master T ON T.Trade_Code = ISM.Cust_Code " & _
                    "where ism.cust_code='" & Trim(cboCust.Text) & "' and ism.invoice_date>='" & Format(invdate1.Value, "yyyy-mm-dd") & "' and " & _
                        "ism.invoice_date<='" & Format(invdate2.Value, "yyyy-mm-dd") & "' and ism.fix_cls='1' " & _
                        "and ism.invoice_no in (select invoice_no from ar_detail where cust_code=ism.cust_code and ar_no='" & Trim(cboapno.Text) & "') " & _
                    ")dt2 " & _
                "order by invoice_date, invoice_no "
    If rsbrowse.State <> adStateClosed Then rsbrowse.Close
    rsbrowse.Open sqlbrowse, Db, adOpenDynamic, adLockOptimistic
    
    Call resetArr
    With grid
        If Not (rsbrowse.BOF And rsbrowse.EOF) Then
            Header
            i = 2
            
            Do While Not rsbrowse.EOF
                .Rows = .Rows + 1
                
                .Cell(flexcpChecked, i, bteColSelect) = IIf(rsbrowse("stAR") = 1, flexChecked, flexUnchecked)
                .Cell(flexcpBackColor, i, bteColSelect) = vbWhite
                .TextMatrix(i, bteColInvDate) = Format(rsbrowse("invoice_date"), "dd MMM yyyy")
                .TextMatrix(i, bteColInvNo) = IIf(IsNull(rsbrowse("invoice_no")), "", Trim(rsbrowse("invoice_no")))
                .TextMatrix(i, bteColCurrCls) = Trim(rsbrowse("currency_code"))
                .TextMatrix(i, bteColCurr) = uf_GetCurrencyDescription(Trim(rsbrowse("currency_code")))
                .TextMatrix(i, bteColInvAmount) = IIf(IsNull(rsbrowse("amount")), 0, Format(rsbrowse("amount"), gs_formatAmountIDR))
                .TextMatrix(i, bteColInvPPn) = IIf(IsNull(rsbrowse("PPNIDR")), 0, Format(rsbrowse("PPNIDR"), gs_formatAmountIDR))
                If .TextMatrix(i, bteColCurrCls) = "03" Then
                    .TextMatrix(i, bteColInvTotal) = Format(CDbl(.TextMatrix(i, bteColInvAmount)) + CDbl(.TextMatrix(i, bteColInvPPn)), gs_formatAmountIDR)
                Else
                    .TextMatrix(i, bteColInvTotal) = ""
                End If
                .TextMatrix(i, bteColCollectAmount) = IIf(IsNull(rsbrowse("aramount")), 0, Format(rsbrowse("aramount"), gs_formatAmountIDR))
                .TextMatrix(i, bteColCollectPPn) = IIf(IsNull(rsbrowse("arppn")), 0, Format(rsbrowse("arppn"), gs_formatAmountIDR))
                .Cell(flexcpBackColor, i, bteColCollectAmount) = vbWhite
                .Cell(flexcpBackColor, i, bteColCollectPPn) = vbWhite
                .TextMatrix(i, bteColRemAmount) = IIf(IsNull(rsbrowse("remainingamount")), 0, Format(rsbrowse("remainingamount"), gs_formatAmountIDR))
                .TextMatrix(i, bteColRemPPn) = IIf(IsNull(rsbrowse("remainingppn")), 0, Format(rsbrowse("remainingppn"), gs_formatAmountIDR))
                .TextMatrix(i, bteColDueDate) = IIf(IsNull(rsbrowse("due_date")), "", Format(rsbrowse("due_Date"), "dd MMM yyyy"))
                
                .TextMatrix(i, bteColARAmount) = IIf(IsNull(rsbrowse("aramount")), 0, Format(rsbrowse("aramount"), gs_formatAmountIDR))
                .TextMatrix(i, bteColARPPn) = IIf(IsNull(rsbrowse("arppn")), 0, Format(rsbrowse("arppn"), gs_formatAmountIDR))
                .TextMatrix(i, bteColARRemAmount) = IIf(IsNull(rsbrowse("remainingamount")), 0, Format(rsbrowse("remainingamount"), gs_formatAmountIDR))
                .TextMatrix(i, bteColARRemPPn) = IIf(IsNull(rsbrowse("remainingppn")), 0, Format(rsbrowse("remainingppn"), gs_formatAmountIDR))
                .TextMatrix(i, bteColPaidCls) = rsbrowse!Paid_Cls
                .TextMatrix(i, bteColARTotAmount) = IIf(IsNull(rsbrowse("TotARAmount")), 0, Format(rsbrowse("TotARAmount"), gs_formatAmountIDR))
                .TextMatrix(i, bteColARTotPPn) = IIf(IsNull(rsbrowse("TotARPPN")), 0, Format(rsbrowse("TotARPPN"), gs_formatAmountIDR))
                
                '**************** Itung Total ****************
                If Curr(Val(.TextMatrix(i, bteColCurrCls))) = "" Then Curr(Val(.TextMatrix(i, bteColCurrCls))) = Trim(.TextMatrix(i, bteColCurrCls))
                InvoiceAmount(Val(.TextMatrix(i, bteColCurrCls))) = InvoiceAmount(Val(.TextMatrix(i, bteColCurrCls))) + CDbl(.TextMatrix(i, bteColInvAmount))
                
                InvoicePPN = InvoicePPN + CDbl(.TextMatrix(i, bteColInvPPn))
                
                CollectionAmount(Val(.TextMatrix(i, bteColCurrCls))) = CollectionAmount(Val(.TextMatrix(i, bteColCurrCls))) + CDbl(.TextMatrix(i, bteColCollectAmount))
                
                CollectionPPN = CollectionPPN + CDbl(.TextMatrix(i, bteColCollectPPn))
                
                GrandTotalCollection(Val(.TextMatrix(i, bteColCurrCls))) = GrandTotalCollection(Val(.TextMatrix(i, bteColCurrCls))) + rsbrowse("TotARAmount")
                
                GrandTotalCollectionPPN = GrandTotalCollectionPPN + rsbrowse("TotARPPN")
                
                RemainingAmount(Val(.TextMatrix(i, bteColCurrCls))) = RemainingAmount(Val(.TextMatrix(i, bteColCurrCls))) + CDbl(.TextMatrix(i, bteColRemAmount))
                
                RemainingPPN = RemainingPPN + CDbl(.TextMatrix(i, bteColRemPPn))
                '************************************************
                
                rsbrowse.MoveNext
                i = i + 1
            Loop
        End If
    End With
    If rsbrowse.State <> adStateClosed Then rsbrowse.Close
    
    Call isiGridTotal
End Sub

Sub resetArr()
    Dim j As Integer
    For j = 1 To 6
        Curr(j) = ""
        CollectionAmount(j) = 0
        InvoiceAmount(j) = 0
        GrandTotalCollection(j) = 0
        RemainingAmount(j) = 0
    Next j
    CollectionPPN = 0
    InvoicePPN = 0
    GrandTotalCollectionPPN = 0
    RemainingPPN = 0
End Sub

Sub isiGridTotal()
    With Grid2
        Call headerTotal
        For i = 1 To 6
            If Curr(i) <> "" Then
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, bteColCurr2) = uf_GetCurrencyDescription(Format((i), "00"))
                .TextMatrix(.Rows - 1, bteColCollect2) = Format(CollectionAmount(i), gs_formatAmountIDR)
                .TextMatrix(.Rows - 1, bteColTotInv) = Format(InvoiceAmount(i), gs_formatAmountIDR)
                .TextMatrix(.Rows - 1, bteColTotCollect) = Format(GrandTotalCollection(i), gs_formatAmountIDR)
                .TextMatrix(.Rows - 1, bteColTotRem) = Format(RemainingAmount(i), gs_formatAmountIDR)
            End If
        Next i
        If InvoicePPN > 0 Then
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, bteColCurr2) = "PPN (IDR)"
            .TextMatrix(.Rows - 1, bteColCollect2) = Format(CollectionPPN, gs_formatAmountIDR)
            .TextMatrix(.Rows - 1, bteColTotInv) = Format(InvoicePPN, gs_formatAmountIDR)
            .TextMatrix(.Rows - 1, bteColTotCollect) = Format(GrandTotalCollectionPPN, gs_formatAmountIDR)
            .TextMatrix(.Rows - 1, bteColTotRem) = Format(RemainingPPN, gs_formatAmountIDR)
        End If
    End With
End Sub

Sub itungTot()
Call resetArr
With grid
    For i = 2 To .Rows - 1
        If Curr(Val(.TextMatrix(i, bteColCurrCls))) = "" Then Curr(Val(.TextMatrix(i, bteColCurrCls))) = Trim(.TextMatrix(i, bteColCurrCls))
        If .Cell(flexcpChecked, i, bteColSelect) = flexChecked Then
            CollectionAmount(Val(.TextMatrix(i, bteColCurrCls))) = CollectionAmount(Val(.TextMatrix(i, bteColCurrCls))) + CDbl(.TextMatrix(i, bteColCollectAmount))
            CollectionPPN = CollectionPPN + CDbl(.TextMatrix(i, bteColCollectPPn))
        End If
        
        InvoiceAmount(Val(.TextMatrix(i, bteColCurrCls))) = InvoiceAmount(Val(.TextMatrix(i, bteColCurrCls))) + CDbl(.TextMatrix(i, bteColInvAmount))
        InvoicePPN = InvoicePPN + CDbl(.TextMatrix(i, bteColInvPPn))
        GrandTotalCollection(Val(.TextMatrix(i, bteColCurrCls))) = GrandTotalCollection(Val(.TextMatrix(i, bteColCurrCls))) + CDbl(.TextMatrix(i, bteColARTotAmount)) - CDbl(.TextMatrix(i, bteColARAmount)) + CDbl(.TextMatrix(i, bteColCollectAmount))
        GrandTotalCollectionPPN = GrandTotalCollectionPPN + CDbl(.TextMatrix(i, bteColARTotPPn)) - CDbl(.TextMatrix(i, bteColARPPn)) + CDbl(.TextMatrix(i, bteColCollectPPn))
        RemainingAmount(Val(.TextMatrix(i, bteColCurrCls))) = RemainingAmount(Val(.TextMatrix(i, bteColCurrCls))) + CDbl(.TextMatrix(i, bteColRemAmount))
        RemainingPPN = RemainingPPN + CDbl(.TextMatrix(i, bteColRemPPn))
        
    Next i
End With

Call isiGridTotal
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)

With grid
    If .Col = bteColCollectAmount Or .Col = bteColCollectPPn Then
        If .TextMatrix(Row, bteColCollectAmount) = "" Then .TextMatrix(Row, bteColCollectAmount) = 0
        If IsNumeric(.TextMatrix(Row, bteColCollectAmount)) = False Then .TextMatrix(Row, bteColCollectAmount) = 0
        
        If Col = bteColCollectAmount Then
        '"Quantity must be lower or equal than 99,999,999,999,999,999.99"
            If CDbl(.TextMatrix(Row, bteColCollectAmount)) > gd_MaxAmount Then LblErrMsg = DisplayMsg("0047") & " " & gd_MaxAmount: .SetFocus: Exit Sub
            .TextMatrix(Row, bteColCollectAmount) = Format(.TextMatrix(Row, bteColCollectAmount), gs_formatAmountIDR)
            .TextMatrix(Row, bteColRemAmount) = Format((CDbl(.TextMatrix(Row, bteColARRemAmount)) + CDbl(.TextMatrix(Row, bteColARAmount)) - CDbl(.TextMatrix(Row, bteColCollectAmount))), gs_formatAmountIDR)
        ElseIf Col = bteColCollectPPn Then
            If CDbl(.TextMatrix(Row, bteColCollectPPn)) > gd_MaxAmount Then LblErrMsg = DisplayMsg("0047") & " " & gd_MaxAmount: .SetFocus: Exit Sub
            .TextMatrix(Row, bteColCollectPPn) = Format(.TextMatrix(Row, bteColCollectPPn), gs_formatAmountIDR)
            .TextMatrix(Row, bteColRemPPn) = Format((CDbl(.TextMatrix(Row, bteColARRemPPn)) + CDbl(.TextMatrix(Row, bteColARPPn)) - CDbl(.TextMatrix(Row, bteColCollectPPn))), gs_formatAmountIDR)
        End If
    End If
    
    If .Col = bteColSelect Or .Col = bteColCollectAmount Or .Col = bteColCollectPPn Then Call itungTot
End With
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If grid.TextMatrix(Row, bteColPaidCls) = "1" Then
        LblErrMsg = DisplayMsg(1139) '"Collection Already Closed"
        Cancel = 1
    Else
        If grid.Cell(flexcpChecked, Row, bteColSelect) <> flexChecked Then
            If grid.Col <> bteColSelect Then Cancel = True
        Else
            If grid.Col <> bteColSelect And grid.Col <> bteColCollectAmount And grid.Col <> bteColCollectPPn Then Cancel = True
        End If
    End If
End Sub

Private Sub Grid_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If grid.Col = bteColCollectAmount Or grid.Col = bteColCollectPPn Then _
    If InStr(1, grid.TextMatrix(Row, Col), ",") = 1 Then grid.TextMatrix(Row, Col) = Right(grid.TextMatrix(Row, Col), Len(grid.TextMatrix(Row, Col)) - 1)
End Sub

Private Sub grid_Click()
LblErrMsg.Caption = ""

    If grid.Row > 0 Then
        If grid.Cell(flexcpChecked, grid.Row, bteColSelect) = flexChecked And grid.Cell(flexcpBackColor, grid.Row, grid.Col) = vbWhite Then
            grid.FocusRect = flexFocusInset
        Else
            grid.FocusRect = flexFocusNone
        End If
    End If
End Sub

Private Sub Grid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
LblErrMsg = ""
  If grid.Col = bteColCollectAmount Or grid.Col = bteColCollectPPn Then
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then
      KeyAscii = 0
    End If
  End If
End Sub

Sub autoapno(ByVal Tgl As String)
Dim sqlno As String
Dim rsno As New Recordset
    
    ' Kawai Format need count Voucher per Month with Format --> S/B/MM/999
    'sqlno = "select max(right(ar_no,4)) from ar_master " & _
            "where left(rtrim(ar_no),4)='" & year(Tgl) & "' " & _
            "group by left(rtrim(ar_no),4)"
    
    sqlno = "Select isnull(max(right(ar_no,3)),0) as ar_no from AR_Master where year(ar_date) = '" _
                & Year(apdate.Value) & "' And Month(ar_Date)= '" & Month(apdate.Value) & "'"

    Set rsno = Db.Execute(sqlno)
    
    
    If Not (rsno.BOF And rsno.EOF) Then
        'cboapno.Text = Format(Tgl, "yyyymmdd") & Format(Val(rsno(0)) + 1, "000#")
        cboapno.Text = "S/B/" & Format(Tgl, "mm") & "/" & Format(Val(rsno(0)) + 1, "00#")
    Else
        'cboapno.Text = Format(Tgl, "yyyymmdd") & "0001"
        cboapno.Text = "S/B/" & Format(Tgl, "mm") & "/" & "001"
    End If
    cboapno.locked = True
End Sub

Private Sub Combo1_Click()

    LblErrMsg = ""
    Header

    If combo1.ListIndex = 0 Then
        Command1(0).Caption = "Create"
        ubah = False
        cboapno.clear
        autoapno (apdate.Value)
    Else
        ubah = True
        Command1(0).Caption = "Update"
        adtocboapno
        cboapno.locked = False
    End If
End Sub

Private Sub combo1_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
  If KeyCode = 13 Then Combo1_Click
End Sub

Private Sub CboApNo_Change()
    LblErrMsg = ""
    Header
End Sub

Private Sub cboapno_Click()
LblErrMsg = ""
    
    sql = "select * from ar_master where ar_no='" & cboapno.Text & "' "
    If RS.State <> adStateClosed Then RS.Close
    RS.Open sql, Db, adOpenKeyset, adLockOptimistic

    If Not (RS.BOF And RS.EOF) Then
        apdate.Value = Format(RS("ar_date"), "dd MMM yyyy")
    End If
    
End Sub

Private Sub cboapno_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
  If KeyCode = 13 Then cboapno_Click
End Sub

Private Sub cboCust_Click()
LblErrMsg = ""

    If cboCust.ListIndex <> -1 Then
        lblcust.Text = cboCust.Column(1)
        If combo1.ListIndex = 1 Then
            adtocboapno
        End If
        Header
        cboapno.Text = ""
    Else
        Kosong
        LblErrMsg.Caption = DisplayMsg(4011) '"Record with this Customer Code not Exist"
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

Private Sub apdate_Change()
    If combo1.ListIndex = 0 Then
        autoapno (apdate)
    End If
        
End Sub

Private Sub invdate1_Change()
   If CDate(invdate1) > CDate(invdate2) Then
      LblErrMsg.Caption = DisplayMsg(4029) & " " & Format(invdate2, "dd MMM yyyy") '"Invoice Date must be lower than "
      Exit Sub
   Else
      LblErrMsg.Caption = ""
   End If
    
    If combo1.ListIndex = 1 And cboCust.Text <> "" Then
        adtocboapno
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
    
    If combo1.ListIndex = 1 And cboCust.Text <> "" Then
        adtocboapno
    End If
    Header
End Sub

Sub Browse()
Dim sqltgl As String, rstgl As New Recordset

    LblErrMsg = ""
    ada = False
    
    sql = "select * from ar_master where ar_no='" & cboapno.Text & "' and cust_code='" & Trim(cboCust.Text) & "' "
    If RS.State <> adStateClosed Then RS.Close
    RS.Open sql, Db, adOpenKeyset, adLockOptimistic

    If Not (RS.BOF And RS.EOF) Then
        ada = True
        
        sqltgl = "select max(invoice_date) as max, min(invoice_date) as min from invoice_master ism " & _
                 "inner join ar_detail ad on ad.invoice_no=ism.invoice_no " & _
                 "where ad.ar_no='" & cboapno.Text & "' "
        Set rstgl = Db.Execute(sqltgl)
        
        If Not IsNull(rstgl("max")) Or IsDate(rstgl("max")) = True Then
            If CDate(rstgl("max")) > CDate(invdate2.Value) Then
                invdate2.Value = Format(rstgl("max"), "dd MMM yyyy")
            End If
        End If
        
        If Not IsNull(rstgl("min")) Or IsDate(rstgl("min")) = True Then
            If CDate(rstgl("min")) < CDate(invdate1.Value) Then
                invdate1.Value = Format(rstgl("min"), "dd MMM yyyy")
            End If
        End If
        Set rstgl = Nothing
        
        browseitem
        apdate.Value = Format(RS("ar_Date"), "dd MMM yyyy")
    
    End If
    If RS.State <> adStateClosed Then RS.Close
End Sub

Private Sub Command1_Click(Index As Integer)
Dim sql4 As String, rs4 As New Recordset
Dim db1 As New Connection, eror As Boolean

LblErrMsg = ""
db1.ConnectionString = Db.ConnectionString
eror = False

Select Case Index
    Case 0:
        If hakUpdate(Me.Name) = 0 Then LblErrMsg = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
        
        If cboCust.Text = "" Then
          cboCust.SetFocus
          LblErrMsg = DisplayMsg(1027) '"Please Select cust Code"
          Exit Sub
        End If

        If cboCust.Text <> "" Then
          cboCust.MatchEntry = 1
          cboCust.Text = cboCust.Text
          If cboCust.MatchFound = False Then
              LblErrMsg = DisplayMsg(4011)
              cboCust.SetFocus
              cboCust.MatchEntry = 2
              Exit Sub
          End If
          cboCust.MatchEntry = 2
        End If
        cboCust.MatchEntry = 1
    
        If combo1.ListIndex = 0 Then
            
            db1.Open
            db1.BeginTrans
                            
            If ubah = False Then
                sql = "select * from ar_master where ar_no='" & cboapno.Text & "' "
                If RS.State <> adStateClosed Then RS.Close
                RS.Open sql, db1, adOpenKeyset, adLockOptimistic
              
                If Not (RS.BOF And RS.EOF) Then
                    autoapno (apdate)
                End If
                  RS.AddNew
                  RS("ar_no") = cboapno.Text
                  RS("cust_code") = cboCust.Text
            End If
            
            RS("ar_date") = Format(apdate.Value, "YYYY-MM-DD")
            RS("last_update") = Now
            RS("last_user") = userLogin
            RS.update
            
            If err.number = 0 Then
                db1.CommitTrans
            Else
                db1.RollbackTrans
                err.clear
            End If
            
            If InStr(1, err.Description, "Violation of PRIMARY KEY constraint") > 0 Then
                autoapno (apdate)
                RS("ar_no") = cboapno.Text
                RS("last_update") = Now
                RS("last_user") = userLogin
                RS.update
            End If

            combo1.Text = "Update"
            browseitem
            LblErrMsg.Caption = DisplayMsg(1000)
            ubah = True
                
        Else
        
            If cboapno.Text = "" Then
              cboapno.SetFocus
              LblErrMsg = DisplayMsg(1095) '"Please Select AR No"
              Exit Sub
            End If
              
             Browse
    
             If ada = False Then
                'kosongBwh
                
                LblErrMsg.Caption = DisplayMsg(4133)    'Record with this AR No not found
                cboapno.SetFocus
                Exit Sub
             End If

        End If

  Case 1:   If hakUpdate(Me.Name) = 0 Then LblErrMsg = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub

            If cboCust.Text = "" Then
              cboCust.SetFocus
              LblErrMsg = DisplayMsg(1027) '"Please Select Cust Code"
              Exit Sub
            ElseIf cboapno.Text = "" Then
              cboapno.SetFocus
              LblErrMsg = DisplayMsg(1095) '"Please Select AR No"
              Exit Sub
            End If
        
                If cboCust.Text <> "" Then
                  cboCust.MatchEntry = 1
                  cboCust.Text = cboCust.Text
                  If cboCust.MatchFound = False Then
                      LblErrMsg = DisplayMsg(4011)
                      cboCust.SetFocus
                      cboCust.MatchEntry = 2
                      Exit Sub
                  End If
                  cboCust.MatchEntry = 2
                End If
                cboCust.MatchEntry = 1
                
                If cboapno.Text <> "" Then
                  cboapno.MatchEntry = 1
                  cboapno.Text = cboapno.Text
                  If cboapno.MatchFound = False Then
                      LblErrMsg = DisplayMsg(4133)
                      cboapno.SetFocus
                      cboapno.MatchEntry = 2
                      Exit Sub
                  End If
                  cboapno.MatchEntry = 2
                End If
                cboapno.MatchEntry = 1
                                      
            If grid.Rows = 1 Then
                LblErrMsg.Caption = DisplayMsg(5012)
                Command1(0).SetFocus
                Exit Sub
            End If
            
            db1.Open
            db1.BeginTrans
            
            sql = "select * from ar_master where ar_no='" & cboapno.Text & "' "
            If RS.State <> adStateClosed Then RS.Close
            RS.Open sql, db1, adOpenKeyset, adLockOptimistic

            If RS.BOF And RS.EOF Then
              LblErrMsg.Caption = DisplayMsg(4133)
              cboapno.SetFocus
              Set db1 = Nothing
              Exit Sub
            End If

            If ubah = True Then
                RS("ar_date") = Format(apdate.Value, "YYYY-MM-DD")
                RS("last_update") = Now
                RS("last_user") = userLogin
                RS.update
                
                With grid
                                                                                                    
                    For i = 1 To .Rows - 1
                        If .Cell(flexcpChecked, i, bteColSelect) = flexChecked Then
                            If .TextMatrix(i, bteColCollectAmount) = 0 And .TextMatrix(i, bteColCollectPPn) = 0 Then
                                If .TextMatrix(i, bteColCollectAmount) = 0 Then .Col = bteColCollectAmount Else .Col = bteColCollectPPn
                                .Row = i
                                .SetFocus
                                LblErrMsg = DisplayMsg(1094) '"Please Input Amount"
                                eror = True
                                GoTo herefirst
                                Exit Sub
                            ElseIf CDbl(.TextMatrix(i, bteColCollectAmount)) > gd_MaxAmount Then
                                LblErrMsg = DisplayMsg("0047") & " " & gd_MaxAmount  '"Quantity must be lower or equal than 99,999,999,999,999,999.99"
                                .Col = bteColCollectAmount
                                .Row = i
                                .SetFocus
                                eror = True
                                GoTo herefirst
                                Exit Sub
                            ElseIf CDbl(.TextMatrix(i, bteColCollectAmount)) > (CDbl(.TextMatrix(i, bteColARRemAmount)) + CDbl(.TextMatrix(i, bteColARAmount))) Then
                                LblErrMsg = DisplayMsg("0047") & " " & Format((CDbl(.TextMatrix(i, bteColARRemAmount)) + CDbl(.TextMatrix(i, bteColARAmount))), gs_formatAmountIDR)
                                .Col = bteColCollectAmount
                                .Row = i
                                .SetFocus
                                eror = True
                                GoTo herefirst
                                Exit Sub
                            ElseIf CDbl(.TextMatrix(i, bteColCollectPPn)) > (CDbl(.TextMatrix(i, bteColARRemPPn)) + CDbl(.TextMatrix(i, bteColARPPn))) Then
                                LblErrMsg = DisplayMsg("0048") & " " & Format((CDbl(.TextMatrix(i, bteColARPPn)) + CDbl(.TextMatrix(i, bteColARRemPPn))), gs_formatAmountIDR)
                                .Col = bteColCollectPPn
                                .Row = i
                                .SetFocus
                                eror = True
                                GoTo herefirst
                                Exit Sub
                            End If
                        End If
                    Next i
                                                                                                                        
                                                                                                    
                    sql4 = "delete from ar_detail where ar_no='" & cboapno.Text & "' "
                    db1.Execute sql4
                    
                    sqlGrid = "select * from ar_detail"
                    If rsGrid.State <> adStateClosed Then rsGrid.Close
                    rsGrid.Open sqlGrid, db1, adOpenKeyset, adLockOptimistic
                    Dim xrate As Double, noCek As Boolean
                    noCek = True
                    For i = 1 To .Rows - 1
                        If .Cell(flexcpChecked, i, bteColSelect) = flexChecked Then
                            noCek = False
                            rsGrid.AddNew
                            rsGrid("ar_no") = cboapno.Text
                            rsGrid("cust_code") = cboCust.Text
                            rsGrid("invoice_no") = .TextMatrix(i, bteColInvNo)
                            rsGrid("currency_code") = .TextMatrix(i, bteColCurrCls)
                            rsGrid("amount") = .TextMatrix(i, bteColCollectAmount)
                            rsGrid("ppn") = .TextMatrix(i, bteColCollectPPn)
                            xrate = exchangeRate(Format(apdate, "YYYY-MM-DD"), .TextMatrix(i, bteColCurrCls))
                            rsGrid("exchange_rate") = xrate
                            rsGrid("exchange_amount") = .TextMatrix(i, bteColCollectAmount) * xrate
                            rsGrid("last_update") = Now
                            rsGrid("last_user") = userLogin
                            rsGrid.update
                        End If
                    Next i
herefirst:
                    If err.number = 0 And eror = False Then
                        db1.CommitTrans
                    Else
                        db1.RollbackTrans
                        err.clear
                        Set db1 = Nothing
                        Exit Sub
                    End If
                    
                    Browse
                    LblErrMsg = DisplayMsg(1101)
                End With
                    
          End If

    Case 2: Kosong
            combo1.ListIndex = 1
            Call Combo1_Click
            cboCust.SetFocus
    
    Case 3:
            If cboapno.Text <> "" And cboCust.Text <> "" Then
                Browse
            End If
            
End Select
Set db1 = Nothing
End Sub

Private Sub CmdSubMenu_Click()
Dim sqlapus As String
    sqlapus = "delete from ar_master where ar_no not in (select ar_no from ar_detail)"
    Db.Execute sqlapus
    Unload Me
    frmMainMenu.Show
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
    If ErrMsg = "" Then Unload Me Else LblErrMsg.Caption = ErrMsg
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If RS.State <> adStateClosed Then RS.Close
    If rsGrid.State <> adStateClosed Then rsGrid.Close
End Sub

Function exchangeRate(Tgl As String, Curr As String) As Double
Dim rstrate As New Recordset
If Curr = "03" Then exchangeRate = 1: Exit Function
sql = "select daily_ExchangeRate from daily_exchangeRate where ExchangeRate_Date = '" & Format(Tgl, "YYYY-MM-DD") & "' and currency_Code = '" & Curr & "'"
If rstrate.State <> adStateClosed Then rstrate.Close
rstrate.Open sql, Db, adOpenStatic, adLockOptimistic
If Not rstrate.EOF Then
    exchangeRate = rstrate!daily_ExchangeRate
Else
    exchangeRate = 0
End If
rstrate.Close
Set rstrate = Nothing
End Function
