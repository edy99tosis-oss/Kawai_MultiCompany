VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form FrmBomCostReport 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Bom Cost Report"
   ClientHeight    =   10140
   ClientLeft      =   2400
   ClientTop       =   1680
   ClientWidth     =   15120
   Icon            =   "FrmBomCostReport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11501.07
   ScaleMode       =   0  'User
   ScaleWidth      =   25788.43
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   600
      Left            =   90
      TabIndex        =   6
      Tag             =   "TFTT*/"
      Top             =   8880
      Width           =   14760
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
         Left            =   210
         TabIndex        =   8
         Tag             =   "TTTF*/"
         Top             =   180
         Width           =   14205
      End
   End
   Begin VB.CommandButton Cmd_Clear 
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
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "FFTT*/"
      Top             =   9660
      Width           =   1125
   End
   Begin VB.CommandButton Cmd_SubMenu 
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
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   7
      Tag             =   "TFFT*/"
      Top             =   9660
      Width           =   1125
   End
   Begin VB.CommandButton Cmd_Excel 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Excel"
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
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "FFTT*/"
      Top             =   9660
      Width           =   1125
   End
   Begin InetCtlsObjects.Inet Inetftp 
      Left            =   3180
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   6390
      Left            =   105
      TabIndex        =   9
      Tag             =   "TTTT*/"
      Top             =   2175
      Width           =   14760
      _cx             =   26035
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
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   13080
      TabIndex        =   10
      Tag             =   "FTTF*/"
      Top             =   120
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   714
   End
   Begin MSComCtl2.DTPicker DTEffective 
      Height          =   330
      Left            =   1680
      TabIndex        =   0
      Tag             =   "TTFF*/"
      Top             =   1080
      Width           =   1395
      _ExtentX        =   2461
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
      Format          =   149684227
      UpDown          =   -1  'True
      CurrentDate     =   37860
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1395
      Left            =   120
      TabIndex        =   14
      Tag             =   "TTTF*/"
      Top             =   720
      Width           =   14760
      Begin VB.CommandButton cmdBrowser 
         BackColor       =   &H00FFFFFF&
         Caption         =   "..."
         Height          =   300
         Index           =   0
         Left            =   3120
         TabIndex        =   2
         Top             =   840
         Width           =   300
      End
      Begin VB.CommandButton CmdSearch 
         BackColor       =   &H80000004&
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
         Height          =   380
         Left            =   8520
         Style           =   1  'Graphical
         TabIndex        =   3
         Tag             =   "TTFF*/"
         Top             =   840
         Width           =   1245
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
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   1
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label txtItemName 
         AutoSize        =   -1  'True
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
         Index           =   1
         Left            =   3480
         TabIndex        =   18
         Tag             =   "TTFF*/"
         Top             =   840
         Width           =   1980
      End
      Begin VB.Line Line8 
         Index           =   0
         X1              =   3480
         X2              =   8400
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label3 
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
         Index           =   2
         Left            =   240
         TabIndex        =   16
         Tag             =   "TTFF*/"
         Top             =   405
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item Code"
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
         TabIndex        =   15
         Tag             =   "TTFF*/"
         Top             =   840
         Width           =   915
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid2 
      Height          =   1335
      Left            =   2400
      TabIndex        =   13
      Tag             =   "TTTT*/"
      Top             =   3000
      Visible         =   0   'False
      Width           =   12240
      _cx             =   21590
      _cy             =   2355
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
   End
   Begin VB.Label lblNm 
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
      Index           =   0
      Left            =   4080
      TabIndex        =   17
      Top             =   840
      Width           =   2565
   End
   Begin MSForms.Label LblRecord 
      Height          =   255
      Left            =   13560
      TabIndex        =   12
      Tag             =   "FFTT*/"
      Top             =   8640
      Width           =   1305
      ForeColor       =   16711935
      BackColor       =   16637923
      VariousPropertyBits=   8388627
      Size            =   "2302;450"
      FontName        =   "Verdana"
      FontEffects     =   1073741827
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   34
      ParagraphAlign  =   2
      FontWeight      =   700
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BOM Cost Report"
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
      Index           =   0
      Left            =   0
      TabIndex        =   11
      Tag             =   "TTTF*/"
      Top             =   240
      Width           =   14610
   End
End
Attribute VB_Name = "FrmBomCostReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dateUp As Date
Dim dbTransfer As New ADODB.Connection

Dim ColA, colNo, ColDesc, ColHSNumber, ColItemCode, ColDescription, ColRegion, ColOriginCountry, ColInvoiceNumber, ColDate As Byte
Dim ColValue, ColPercent As Byte, ColBCNO As Byte, colBCDate As Byte, ColAmount As Byte, ColUSD As Byte


Dim sqlfg As String, sqlmb As String, FGCls As String
Dim SqlData As String

Private Sub up_Header()
Dim X As Integer
    
    colNo = 0
    ColHSNumber = 1
    ColItemCode = 2
    ColDescription = 3
    ColRegion = 4
    ColOriginCountry = 5
    ColInvoiceNumber = 6
    ColDate = 7
    ColValue = 8
    ColPercent = 9
    
    
    With Grid
        .clear
        .ColS = 10
        .Rows = 1
        
        .TextMatrix(0, colNo) = "No"
        .TextMatrix(0, ColHSNumber) = "HS Number"
        .TextMatrix(0, ColItemCode) = "Item Code"
        .TextMatrix(0, ColDescription) = "Description"
        .TextMatrix(0, ColRegion) = "Region"
        .TextMatrix(0, ColOriginCountry) = "Origin Country"
        .TextMatrix(0, ColInvoiceNumber) = "Invoice Number"
        .TextMatrix(0, ColDate) = "Date"
        .TextMatrix(0, ColValue) = "Value (USD)"
        .TextMatrix(0, ColPercent) = "Percent (%)"
        
        .ColAlignment(colNo) = flexAlignCenterCenter
        .ColAlignment(ColHSNumber) = flexAlignCenterCenter
        .ColAlignment(ColItemCode) = flexAlignCenterCenter
        .ColAlignment(ColDescription) = flexAlignLeftCenter
        .ColAlignment(ColRegion) = flexAlignCenterCenter
        .ColAlignment(ColOriginCountry) = flexAlignCenterCenter
        .ColAlignment(ColInvoiceNumber) = flexAlignCenterCenter
        .ColAlignment(ColValue) = flexAlignLeftCenter
        .ColAlignment(ColDate) = flexAlignCenterCenter
        .ColAlignment(ColPercent) = flexAlignCenterCenter
        
        .ColHidden(colNo) = True
'
        .ColWidth(colNo) = 1000
        .ColWidth(ColHSNumber) = 1500
        .ColWidth(ColItemCode) = 2000
        .ColWidth(ColDescription) = 6500
        .ColWidth(ColRegion) = 2000
        .ColWidth(ColOriginCountry) = 2000
        .ColWidth(ColInvoiceNumber) = 2000
        .ColWidth(ColDate) = 2000
        .ColWidth(ColValue) = 2000
        .ColWidth(ColPercent) = 2000
    End With
   
End Sub

Private Sub up_Header2()
Dim X As Integer
    
    colNo = 0
    ColHSNumber = 1
    ColItemCode = 2
    ColDescription = 3
    ColRegion = 4
    ColOriginCountry = 5
    ColInvoiceNumber = 6
    ColDate = 7
    ColValue = 8
    ColPercent = 9
    
    
    With Grid2
        .clear
        .ColS = 10
        .Rows = 1
        
        .TextMatrix(0, colNo) = "No"
        .TextMatrix(0, ColHSNumber) = "HS Number"
        .TextMatrix(0, ColItemCode) = "Item Code"
        .TextMatrix(0, ColDescription) = "Description"
        .TextMatrix(0, ColRegion) = "Region"
        .TextMatrix(0, ColOriginCountry) = "Origin Country"
        .TextMatrix(0, ColInvoiceNumber) = "Invoice Number"
        .TextMatrix(0, ColDate) = "Date"
        .TextMatrix(0, ColValue) = "Value (USD)"
        .TextMatrix(0, ColPercent) = "Percent (%)"
        
        .ColAlignment(colNo) = flexAlignCenterCenter
        .ColAlignment(ColHSNumber) = flexAlignCenterCenter
        .ColAlignment(ColItemCode) = flexAlignLeftCenter
        .ColAlignment(ColDescription) = flexAlignLeftCenter
        .ColAlignment(ColRegion) = flexAlignCenterCenter
        .ColAlignment(ColOriginCountry) = flexAlignCenterCenter
        .ColAlignment(ColInvoiceNumber) = flexAlignCenterCenter
        .ColAlignment(ColValue) = flexAlignLeftCenter
        .ColAlignment(ColDate) = flexAlignCenterCenter
        .ColAlignment(ColPercent) = flexAlignCenterCenter
        
        .ColWidth(colNo) = 500
        .ColWidth(ColHSNumber) = 1500
        .ColWidth(ColItemCode) = 2000
        .ColWidth(ColDescription) = 6500
        .ColWidth(ColRegion) = 2000
        .ColWidth(ColOriginCountry) = 2000
        .ColWidth(ColInvoiceNumber) = 2000
        .ColWidth(ColDate) = 2000
        .ColWidth(ColValue) = 2000
        .ColWidth(ColPercent) = 2000
    End With
   
End Sub

Private Sub cmdBrowser_Click(Index As Integer)
    Me.MousePointer = vbHourglass
 Select Case Index
  Case 0:
   frm_BrowseItemCode.getItemCode = txtItemCode.Text
   frm_BrowseItemCode.Show 1
   txtItemCode.Text = frm_BrowseItemCode.getItemCode
   txtItemName(1) = frm_BrowseItemCode.getItemName
  Case 1:
   If txtItemCode.Enabled = True Then
    frm_BrowseItemCode.getItemCode = txtItemCode.Text
    frm_BrowseItemCode.Show 1
    txtItemCode.Text = frm_BrowseItemCode.getItemCode
    txtItemName(1) = frm_BrowseItemCode.getItemName
   End If
  End Select
 Me.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
'LblRecord = "0"
    If gb_Simulation = True Then Call up_InitSimulation(Me) 'Editan

    CtrlMenu1.FormName = Me.Name    'Editan
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"  'Editan
    DTEffective.Value = Now
        
    Call up_Header
    Call up_Header2
    
End Sub

Private Sub Cmd_Excel_Click()
    Dim lb_OFFICE_UNDER2010 As Boolean
    Dim xlapp As New Excel.application
    Dim a As Integer, X As Integer
    Dim TglEnd1 As String, strSQL As String
    Dim Region As String
    Dim Country As String
    Dim RsSearch As New ADODB.Recordset
    
    Dim Idx As Integer
    Dim IdxA As Integer
    Dim IdxB As Integer
    Dim IdxC As Integer
    Dim TotalAsean As Integer
    Dim TotalAseanV As Integer
    Dim xlColA As String
    Dim xlColNo As String
    Dim xlColItemCode As String
    Dim xlColHSNumber As String
    Dim xlColDescription As String
    Dim xlColRegion As String
    Dim xlColOriginCountry As String
    Dim xlColInvoiceNumber As String
    Dim xlColDate As String
    Dim xlColValue As String
    Dim xlColPercent As String
    Dim xlcolUSD As String
    Dim xlColBCNo As String
    Dim xlColBCDate As String
    Dim xlColAmount As String
    
    
    
    If Grid.Rows > 1 Then
    
    FGCls = "01"
    
    For X = 1 To 3
    
            LblErrMsg = ""
            Me.MousePointer = vbHourglass
            
               
             xlColA = "a"
             xlColNo = "b"
             xlColHSNumber = "c"
             xlColItemCode = "d"
             xlColDescription = "e"
             xlColRegion = "f"
             xlColOriginCountry = "g"
             xlColInvoiceNumber = "h"
             xlColDate = "i"
             xlColValue = "j"
             xlColPercent = "k"
             xlcolUSD = "l"
             xlColBCNo = "m"
             xlColBCDate = "n"
             xlColAmount = "o"
                          
    
            With xlapp
                
'
                If FGCls = "01" Then
                    .Workbooks.Add
                    .Sheets.Add
                End If
                                
                If FGCls = "01" Then
                    '.Sheets.Add
                    '.Sheets.Add
                    .Sheets("Sheet2").Select
                    .Sheets("Sheet2").Name = "CS"
                    
                    .Range(xlColA & "1:" & xlColPercent & "2").Merge
                    .Range(xlColA & "1") = "PT KAWAI INDONESIA"
                    .Range(xlColA & "1").Font.Size = 18
                    .Range(xlColA & "1").HorizontalAlignment = xlLeft
                    .Range(xlColA & "1").VerticalAlignment = xlCenter
                    .Range(xlColA & "1").Font.Bold = True
                         
                    .Range(xlColA & "3:" & xlColPercent & "3").Merge
                    .Range(xlColA & "3") = "COST STRUCTURE / BILL OF MATERIALS PER UNIT"
                    .Range(xlColA & "3").HorizontalAlignment = xlCenter
                    .Range(xlColA & "3").VerticalAlignment = xlCenter
                    .Range(xlColA & "3").Font.Bold = False
             
                    .Range(xlColA & "4:" & xlColPercent & "4").Merge
                    .Range(xlColA & "4") = "BASED ON EX-WORKS PRICE"
                    .Range(xlColA & "4").HorizontalAlignment = xlCenter
                    .Range(xlColA & "4").VerticalAlignment = xlCenter
                    .Range(xlColA & "4").Font.Bold = False
                    
                    .Range(xlColA & "7:" & xlColNo & "7").Merge
                    .Range(xlColA & "7") = "Model"
                    .Range(xlColA & "7").HorizontalAlignment = xlLeft
                    .Range(xlColA & "7").HorizontalAlignment = xlCenter
                    
                    .Range(xlColHSNumber & "7") = txtItemCode
                    .Range(xlColHSNumber & "7").HorizontalAlignment = xlLeft
                    .Range(xlColHSNumber & "7").VerticalAlignment = xlCenter
                    .Range(xlColHSNumber & "7").Font.Bold = True
                    
                    .Range(xlColItemCode & "7") = txtItemName(1)
                    .Range(xlColItemCode & "7").HorizontalAlignment = xlLeft
                    .Range(xlColItemCode & "7").VerticalAlignment = xlCenter
                    .Range(xlColItemCode & "7").Font.Bold = True
                    
                    '------------Components Imported or Unknown Origin
                    .Range(xlColA & "9:" & xlColDescription & "9").Merge
                    .Range(xlColA & "9") = "A. Components Imported or Unknown Origin"
                    .Range(xlColA & "9").HorizontalAlignment = xlLeft
                    .Range(xlColA & "9").VerticalAlignment = xlCenter
                    
                    .Range(xlColNo & "10") = "No."
                    .Range(xlColHSNumber & "10") = "HS Number"
                    .Range(xlColItemCode & "10") = "Item Code"
                    .Range(xlColDescription & "10") = "Description"
                    .Range(xlColRegion & "10") = "Region"
                    .Range(xlColOriginCountry & "10") = "Origin Country"
                    .Range(xlColInvoiceNumber & "10") = "Invoice Number"
                    .Range(xlColDate & "10") = "Date"
                    .Range(xlColValue & "10") = "Value (USD)"
                    .Range(xlColPercent & "10") = "%"
                    
                    .Range(xlColHSNumber & "12") = "Total"
                    .Range(xlColHSNumber & "12").Font.Bold = True
                    .Range(xlColHSNumber & "12:" & xlColPercent & "12").Interior.ColorIndex = 6
                    .Range(xlColValue & "12").Font.Bold = True
                    
                    Idx = 12
                    IdxA = Idx
                    
                    TglEnd1 = DTEffective
                    TglEnd1 = Format(TglEnd1, "mm/dd/yyyy")
                                        
                    strSQL = " Declare @TglEnd as date=(SELECT DATEADD(s,-1,DATEADD(M, DATEDIFF(MONTH,0,'" & TglEnd1 & "')+1,0))) " & vbCrLf & _
                             " Select * From Bom_Cost a left join Region_Cls b on a.Region=b.Description " & vbCrLf & _
                             " where Parent_ItemCode='" & txtItemCode & "' and Period = @TglEnd and b.Region_Cls = '09'"

                    If RsSearch.State <> adStateClosed Then RsSearch.Close
                    RsSearch.CursorLocation = adUseClient
                    RsSearch.Open strSQL, Db, adOpenForwardOnly, adLockOptimistic
                    
                    
                    i = 0
                    '***
                    Do While Not RsSearch.EOF
                        i = i + 1
                        Grid2.AddItem i
                        Grid2.TextMatrix(i, colNo) = Grid2.TextMatrix(i, colNo)
                        Grid2.TextMatrix(i, ColHSNumber) = Trim(RsSearch!HS_Code)
                        Grid2.TextMatrix(i, ColItemCode) = Trim(RsSearch!Item_Code)
                        Grid2.TextMatrix(i, ColDescription) = Trim(RsSearch!item_name)
                        Grid2.TextMatrix(i, ColRegion) = IIf(IsNull(Trim(RsSearch("Region"))), "", Trim(RsSearch("Region")))
                        Grid2.TextMatrix(i, ColOriginCountry) = IIf(IsNull(Trim(RsSearch("Origin_Country"))), "", Trim(RsSearch("Origin_Country")))
                        Grid2.TextMatrix(i, ColInvoiceNumber) = Trim(RsSearch!BC40_No)
                        Grid2.TextMatrix(i, ColDate) = Format(RsSearch!BC40_Date, "dd MMM yyyy")
                        Grid2.TextMatrix(i, ColValue) = Trim(RsSearch!Value)
                        Grid2.TextMatrix(i, ColPercent) = Format(RsSearch!Percentage, gs_formatPercentage)

                        RsSearch.MoveNext
                    Loop
                    
                    '#Fill Grid
                                   
                    For i = 1 To Grid2.Rows - 1
                        Idx = Idx + 1
                        
                        .Range(xlColNo & Idx) = Grid2.TextMatrix(i, colNo)
                        .Range(xlColHSNumber & Idx) = "'" + Grid2.TextMatrix(i, ColHSNumber)
                        .Range(xlColItemCode & Idx) = Grid2.TextMatrix(i, ColItemCode)
                        .Range(xlColDescription & Idx) = Grid2.TextMatrix(i, ColDescription)
                        .Range(xlColRegion & Idx) = Grid2.TextMatrix(i, ColRegion)
                        .Range(xlColOriginCountry & Idx) = Grid2.TextMatrix(i, ColOriginCountry)
                        .Range(xlColInvoiceNumber & Idx) = "'" + Grid2.TextMatrix(i, ColInvoiceNumber)
                        .Range(xlColDate & Idx) = Grid2.TextMatrix(i, ColDate)
                        .Range(xlColValue & Idx) = Grid2.TextMatrix(i, ColValue)
                        .Range(xlColPercent & Idx) = Grid2.TextMatrix(i, ColPercent)
                    Next
                
                    .Columns(xlColA & ":" & xlColNo).ColumnWidth = 2.57
                    .Columns(xlColNo & ":" & xlColNo).ColumnWidth = 5
                    .Columns(xlColHSNumber & ":" & xlColHSNumber).ColumnWidth = 13
                    .Columns(xlColItemCode & ":" & xlColItemCode).ColumnWidth = 13
                    .Columns(xlColDescription & ":" & xlColDescription).ColumnWidth = 62
                    .Columns(xlColRegion & ":" & xlColRegion).ColumnWidth = 16
                    .Columns(xlColOriginCountry & ":" & xlColOriginCountry).ColumnWidth = 16
                    .Columns(xlColInvoiceNumber & ":" & xlColInvoiceNumber).ColumnWidth = 16
                    .Columns(xlColDate & ":" & xlColDate).ColumnWidth = 11
                    .Columns(xlColValue & ":" & xlColValue).ColumnWidth = 13
                    .Columns(xlColPercent & ":" & xlColPercent).ColumnWidth = 9
                  
                    .Range(xlColNo & "10:" & xlColPercent & Idx).Select
                    
                    If i = 1 Then
                        .Range(xlColValue & "12").Formula = 0
                        .Range(xlColValue & "12").HorizontalAlignment = xlRight
                        .Range(xlColPercent & "12").HorizontalAlignment = xlRight
                        .Range(xlColPercent & "12").Formula = 0
                        .Range(xlColPercent & "12:" & xlColPercent & Idx).NumberFormat = "0.00 %"
                    Else
                        .Range(xlColValue & "12").Formula = "=Sum(" & xlColValue & "13:" & xlColValue & Idx & ")"
                        .Range(xlColValue & "12").HorizontalAlignment = xlRight
                        .Range(xlColPercent & "12").HorizontalAlignment = xlRight
                        .Range(xlColPercent & "12").Formula = "=Sum(" & xlColPercent & "13:" & xlColPercent & Idx & ")"
                        .Range(xlColPercent & "12:" & xlColPercent & Idx).NumberFormat = "0.00 %"
                    End If
                    
                    
                    .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
                    .Selection.Borders(xlDiagonalUp).LineStyle = xlNone

                    With .Selection.Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With

                    With .Selection.Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    With .Selection.Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    With .Selection.Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    With .Selection.Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    With .Selection.Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    
                    .Range(xlColNo & "10:" & xlColPercent & "10").Select
                    .Range(xlColNo & "10:" & xlColPercent & "10").Font.Bold = True
                    
                    With .Selection
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                    End With
                    
                    .Range(xlColValue & "11:" & xlColValue & Idx).Select
                    With .Selection
                        .NumberFormat = gs_formatPriceIDR
                    End With
                    
                    .Range(xlColNo & "12:" & xlColPercent & "12").Select
                    .Range(xlColNo & "12:" & xlColPercent & "12").Font.Bold = True
                    
                    Call up_Header2
                                         
        '-----------Template From ASEAN Countries
                     
                     Idx = Idx + 2
                     a = Idx + 1
                     
                    .Range(xlColA & (Idx & ":") & xlColDescription & Idx).Merge
                    .Range(xlColA & Idx) = "B. Components From ASEAN Countries"
                    .Range(xlColA & Idx).HorizontalAlignment = xlLeft
                    .Range(xlColA & Idx).VerticalAlignment = xlCenter
                    
                    Idx = Idx + 1
                    
                    .Range(xlColNo & Idx) = "No."
                    .Range(xlColHSNumber & Idx) = "HS Number"
                    .Range(xlColItemCode & Idx) = "Item Code"
                    .Range(xlColDescription & Idx) = "Description"
                    .Range(xlColRegion & Idx) = "Region"
                    .Range(xlColOriginCountry & Idx) = "Origin Country"
                    .Range(xlColInvoiceNumber & Idx) = "Invoice Number"
                    .Range(xlColDate & Idx) = "Date"
                    .Range(xlColValue & Idx) = "Value (USD)"
                    .Range(xlColPercent & Idx) = "%"
                    
                    .Range(xlColNo & (Idx & ":") & xlColPercent & Idx).Select
                    .Range(xlColNo & (Idx & ":") & xlColPercent & Idx).Font.Bold = True
                    With .Selection
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                    End With
                    
                    TotalAsean = Idx + 2
                    
                    .Range(xlColHSNumber & TotalAsean) = "Total"
                    .Range(xlColHSNumber & TotalAsean).Font.Bold = True
                    .Range(xlColHSNumber & (TotalAsean & ":") & xlColPercent & TotalAsean).Interior.ColorIndex = 6
                    .Range(xlColPercent & TotalAsean).Font.Bold = True
                    
                    IdxB = Idx
                    
                    TglEnd1 = DTEffective
                    TglEnd1 = Format(TglEnd1, "mm/dd/yyyy")
                    Region = "08"
                        
                                                                        
                    strSQL = " Declare @TglEnd as date=(SELECT DATEADD(s,-1,DATEADD(M, DATEDIFF(MONTH,0,'" & TglEnd1 & "')+1,0))) " & vbCrLf & _
                             " Select * From Bom_Cost a left join Region_Cls b on a.Region=b.Description " & vbCrLf & _
                             " where Parent_ItemCode='" & txtItemCode & "' and Period = @TglEnd and b.Region_Cls='" & Region & "' "
                                                                                
                    If RsSearch.State <> adStateClosed Then RsSearch.Close
                    RsSearch.CursorLocation = adUseClient
                    RsSearch.Open strSQL, Db, adOpenForwardOnly, adLockOptimistic
                        
                    '#Fill Grid
                    i = 0
                    Do While Not RsSearch.EOF
                        i = i + 1
                        Grid2.AddItem i
                        Grid2.TextMatrix(i, colNo) = Grid2.TextMatrix(i, colNo)
                        Grid2.TextMatrix(i, ColHSNumber) = Trim(RsSearch!HS_Code)
                        Grid2.TextMatrix(i, ColItemCode) = Trim(RsSearch!Item_Code)
                        Grid2.TextMatrix(i, ColDescription) = Trim(RsSearch!item_name)
                        Grid2.TextMatrix(i, ColRegion) = IIf(IsNull(Trim(RsSearch("Region"))), "", Trim(RsSearch("Region")))
                        Grid2.TextMatrix(i, ColOriginCountry) = IIf(IsNull(Trim(RsSearch("Origin_Country"))), "", Trim(RsSearch("Origin_Country")))
                        Grid2.TextMatrix(i, ColInvoiceNumber) = Trim(RsSearch!BC40_No)
                        Grid2.TextMatrix(i, ColDate) = Format(RsSearch!BC40_Date, "dd MMM yyyy")
                        Grid2.TextMatrix(i, ColValue) = Trim(RsSearch!Value)
                        Grid2.TextMatrix(i, ColPercent) = Format(RsSearch!Percentage, gs_formatPercentage)
                        
                        RsSearch.MoveNext
                    Loop
                    
                    Idx = Idx + 2
                    
                    For i = 1 To Grid2.Rows - 1
                        Idx = Idx + 1
                        .Range(xlColNo & Idx) = Grid2.TextMatrix(i, colNo)
                        .Range(xlColHSNumber & Idx) = "'" + Grid2.TextMatrix(i, ColHSNumber)
                        .Range(xlColItemCode & Idx) = Grid2.TextMatrix(i, ColItemCode)
                        .Range(xlColDescription & Idx) = Grid2.TextMatrix(i, ColDescription)
                        .Range(xlColRegion & Idx) = Grid2.TextMatrix(i, ColRegion)
                        .Range(xlColOriginCountry & Idx) = Grid2.TextMatrix(i, ColOriginCountry)
                        .Range(xlColInvoiceNumber & Idx) = "'" + Grid2.TextMatrix(i, ColInvoiceNumber)
                        .Range(xlColDate & Idx) = Grid2.TextMatrix(i, ColDate)
                        .Range(xlColValue & Idx) = Grid2.TextMatrix(i, ColValue)
                        .Range(xlColPercent & Idx) = Grid2.TextMatrix(i, ColPercent)
                    Next
                        
                    .Columns(xlColA & ":" & xlColNo).ColumnWidth = 2.57
                    .Columns(xlColNo & ":" & xlColNo).ColumnWidth = 5
                    .Columns(xlColItemCode & ":" & xlColItemCode).ColumnWidth = 13
                    .Columns(xlColHSNumber & ":" & xlColHSNumber).ColumnWidth = 13
                    .Columns(xlColDescription & ":" & xlColDescription).ColumnWidth = 62
                    .Columns(xlColRegion & ":" & xlColRegion).ColumnWidth = 16
                    .Columns(xlColOriginCountry & ":" & xlColOriginCountry).ColumnWidth = 16
                    .Columns(xlColInvoiceNumber & ":" & xlColInvoiceNumber).ColumnWidth = 16
                    .Columns(xlColDate & ":" & xlColDate).ColumnWidth = 11
                    .Columns(xlColValue & ":" & xlColValue).ColumnWidth = 13
                    .Columns(xlColPercent & ":" & xlColPercent).ColumnWidth = 9
                    
                    If i = 1 Then
                        .Range(xlColPercent & TotalAsean).HorizontalAlignment = xlRight
                        .Range(xlColPercent & TotalAsean).Formula = 0
                        .Range(xlColPercent & (TotalAsean & ":") & xlColPercent & Idx).NumberFormat = "0.00 %"
                        .Range(xlColValue & TotalAsean).Formula = 0
                    Else
                        TotalAseanV = TotalAsean + 1
                        .Range(xlColPercent & TotalAsean).HorizontalAlignment = xlRight
                        .Range(xlColPercent & TotalAsean).Formula = "=Sum(" & xlColPercent & (TotalAseanV & ":") & xlColPercent & Idx & ")"
                        .Range(xlColPercent & (TotalAsean & ":") & xlColPercent & Idx).NumberFormat = "0.00 %"
                        .Range(xlColValue & TotalAsean).Formula = "=Sum(" & xlColValue & (TotalAseanV & ":") & xlColValue & Idx & ")"
                    End If
                    
                    .Range(xlColNo & (a & ":") & xlColPercent & Idx).Select
                    
                    .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
                    .Selection.Borders(xlDiagonalUp).LineStyle = xlNone


                    With .Selection.Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    With .Selection.Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    With .Selection.Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    With .Selection.Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    With .Selection.Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    With .Selection.Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    
                    .Range(xlColValue & (a & ":") & xlColValue & Idx).Select
                    With .Selection
                        .NumberFormat = gs_formatPriceIDR
                    End With
                    
                    Call up_Header2
                                        
                    '------Components From ASEAN Indonesia
                    Idx = Idx + 2
                     a = Idx + 1
                     
                    .Range(xlColA & (Idx & ":") & xlColDescription & Idx).Merge
                    .Range(xlColA & Idx) = "C. Components From ASEAN Indonesia"
                    .Range(xlColA & Idx).HorizontalAlignment = xlLeft
                    .Range(xlColA & Idx).VerticalAlignment = xlCenter
                    
                    Idx = Idx + 1
                    
                    .Range(xlColNo & Idx) = "No."
                    .Range(xlColHSNumber & Idx) = "HS Number"
                    .Range(xlColItemCode & Idx) = "Item Code"
                    .Range(xlColDescription & Idx) = "Description"
                    .Range(xlColRegion & Idx) = "Region"
                    .Range(xlColOriginCountry & Idx) = "Origin Country"
                    .Range(xlColInvoiceNumber & Idx) = "Invoice Number"
                    .Range(xlColDate & Idx) = "Date"
                    .Range(xlColValue & Idx) = "Value (USD)"
                    .Range(xlColPercent & Idx) = "%"
                    .Range(xlColNo & (Idx & ":") & xlColPercent & Idx).Select
                    .Range(xlColNo & (Idx & ":") & xlColPercent & Idx).Font.Bold = True
                     With .Selection
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                    End With
                                        
                    Idx = Idx + 2
                    
                    .Range(xlColDescription & Idx) = "ATTACHED"
                    .Range(xlColDescription & (Idx & ":") & xlColPercent & Idx).Font.Bold = True
                    .Range(xlColDescription & (Idx & ":") & xlColPercent & Idx).Interior.ColorIndex = 6
                    .Range(xlColOriginCountry & Idx) = "INDONESIA"
                    .Range(xlColValue & Idx).NumberFormat = gs_formatPriceIDR
                    .Range(xlColPercent & Idx).NumberFormat = "0.00 %"
                 
                    .Columns(xlColA & ":" & xlColNo).ColumnWidth = 2.57
                    .Columns(xlColNo & ":" & xlColNo).ColumnWidth = 5
                    .Columns(xlColHSNumber & ":" & xlColHSNumber).ColumnWidth = 13
                    .Columns(xlColItemCode & ":" & xlColItemCode).ColumnWidth = 13
                    .Columns(xlColDescription & ":" & xlColDescription).ColumnWidth = 62
                    .Columns(xlColRegion & ":" & xlColRegion).ColumnWidth = 16
                    .Columns(xlColOriginCountry & ":" & xlColOriginCountry).ColumnWidth = 16
                    .Columns(xlColInvoiceNumber & ":" & xlColInvoiceNumber).ColumnWidth = 16
                    .Columns(xlColDate & ":" & xlColDate).ColumnWidth = 11
                    .Columns(xlColValue & ":" & xlColValue).ColumnWidth = 13
                    .Columns(xlColPercent & ":" & xlColPercent).ColumnWidth = 9
                    
                    a = Idx - 2
                    
                    IdxC = Idx
                    
                    .Range(xlColNo & (a & ":") & xlColPercent & Idx).Select
                    
                    .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
                    .Selection.Borders(xlDiagonalUp).LineStyle = xlNone

                    With .Selection.Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    With .Selection.Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    With .Selection.Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    With .Selection.Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    With .Selection.Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    With .Selection.Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    
                    .Range(xlColValue & (a & ":") & xlColValue & Idx).Select
                    With .Selection
                        .NumberFormat = gs_formatPriceIDR
                    End With
                    
                    '-----------Direct Production Cost
                     Idx = Idx + 2
                     a = Idx + 1
                     
                    .Range(xlColA & (Idx & ":") & xlColDescription & Idx).Merge
                    .Range(xlColA & Idx) = "D. Direct Production Cost"
                    .Range(xlColA & Idx).HorizontalAlignment = xlLeft
                    .Range(xlColA & Idx).VerticalAlignment = xlCenter
                    
                    Idx = Idx + 1
                    
                    .Range(xlColNo & Idx) = "No."
                    .Range(xlColHSNumber & (Idx & ":") & xlColDate & Idx).Merge
                    .Range(xlColHSNumber & Idx) = "Direct Production Description"
                    .Range(xlColValue & Idx) = "Value (USD)"
                    .Range(xlColPercent & Idx) = "%"
                    
                    .Columns(xlColA & ":" & xlColNo).ColumnWidth = 2.57
                    .Columns(xlColNo & ":" & xlColNo).ColumnWidth = 5
                    .Columns(xlColHSNumber & ":" & xlColHSNumber).ColumnWidth = 13
                    .Columns(xlColItemCode & ":" & xlColItemCode).ColumnWidth = 13
                    .Columns(xlColDescription & ":" & xlColDescription).ColumnWidth = 62
                    .Columns(xlColRegion & ":" & xlColRegion).ColumnWidth = 16
                    .Columns(xlColOriginCountry & ":" & xlColOriginCountry).ColumnWidth = 16
                    .Columns(xlColInvoiceNumber & ":" & xlColInvoiceNumber).ColumnWidth = 16
                    .Columns(xlColDate & ":" & xlColDate).ColumnWidth = 11
                    .Columns(xlColValue & ":" & xlColValue).ColumnWidth = 13
                    .Columns(xlColPercent & ":" & xlColPercent).ColumnWidth = 9
                    
                    a = Idx
                    
                    .Range(xlColNo & (a & ":") & xlColPercent & Idx).Select

                    .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
                    .Selection.Borders(xlDiagonalUp).LineStyle = xlNone


                    With .Selection.Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    With .Selection.Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    With .Selection.Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    With .Selection.Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    With .Selection.Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    With .Selection.Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    
                    '--------Profit
                    Idx = Idx + 2
                     a = Idx + 1
                     
                    .Range(xlColA & (Idx & ":") & xlColDescription & Idx).Merge
                    .Range(xlColA & Idx) = "E. Profit"
                    .Range(xlColA & Idx).HorizontalAlignment = xlLeft
                    .Range(xlColA & Idx).VerticalAlignment = xlCenter
                    
                    Idx = Idx + 1
                    
                    .Range(xlColNo & Idx) = "No."
                    .Range(xlColHSNumber & (Idx & ":") & xlColDate & Idx).Merge
                    .Range(xlColHSNumber & Idx) = "Direct Production Description"
                    .Range(xlColValue & Idx) = "Value (USD)"
                    .Range(xlColPercent & Idx) = "%"
                    
                    .Columns(xlColA & ":" & xlColNo).ColumnWidth = 2.57
                    .Columns(xlColNo & ":" & xlColNo).ColumnWidth = 5
                    .Columns(xlColHSNumber & ":" & xlColHSNumber).ColumnWidth = 13
                    .Columns(xlColItemCode & ":" & xlColItemCode).ColumnWidth = 13
                    .Columns(xlColDescription & ":" & xlColDescription).ColumnWidth = 62
                    .Columns(xlColRegion & ":" & xlColRegion).ColumnWidth = 16
                    .Columns(xlColOriginCountry & ":" & xlColOriginCountry).ColumnWidth = 16
                    .Columns(xlColInvoiceNumber & ":" & xlColInvoiceNumber).ColumnWidth = 16
                    .Columns(xlColDate & ":" & xlColDate).ColumnWidth = 11
                    .Columns(xlColValue & ":" & xlColValue).ColumnWidth = 13
                    .Columns(xlColPercent & ":" & xlColPercent).ColumnWidth = 9
                    
                    a = Idx
                    
                    .Range(xlColNo & (a & ":") & xlColPercent & Idx).Select

                    .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
                    .Selection.Borders(xlDiagonalUp).LineStyle = xlNone


                    With .Selection.Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    With .Selection.Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    With .Selection.Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    With .Selection.Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    With .Selection.Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    With .Selection.Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    
                    '----Total
                    
                     Idx = Idx + 2
                     a = Idx + 1
                     
                     
                    .Range(xlColA & (Idx & ":") & xlColDescription & Idx).Merge
                    .Range(xlColA & Idx) = "Total"
                    .Range(xlColA & Idx).HorizontalAlignment = xlLeft
                    .Range(xlColA & Idx).VerticalAlignment = xlCenter
                                        
                    Idx = Idx + 1
                    
                    .Range(xlColNo & (Idx & ":") & xlColDate & Idx).Merge
                    .Range(xlColNo & Idx) = "Total Value / Ex-Works Price Value"
                    
                    .Range(xlColValue & Idx).Formula = "=Sum(" & xlColValue & (IdxA & "+") & xlColValue & (TotalAsean & "+") & xlColValue & IdxC & ")"
                                        
                    .Range(xlColPercent & Idx).Formula = "=Sum(" & xlColPercent & (IdxA & "+") & xlColPercent & (TotalAsean & "+") & xlColPercent & IdxC & ")"
                    .Range(xlColPercent & Idx).NumberFormat = "0.00 %"
       
                    .Columns(xlColValue & ":" & xlColValue).ColumnWidth = 13
                    .Columns(xlColPercent & ":" & xlColPercent).ColumnWidth = 9
                    
                    .Range(xlColNo & (Idx & ":") & xlColPercent & Idx).Font.Bold = True
                    .Range(xlColPercent & Idx).Select
                    
                    With .Selection
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                    End With
                    
                    a = Idx
                    
                    .Range(xlColNo & (a & ":") & xlColPercent & Idx).Select
                                        
                    .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
                    .Selection.Borders(xlDiagonalUp).LineStyle = xlNone


                    With .Selection.Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    With .Selection.Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    With .Selection.Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    With .Selection.Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    With .Selection.Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    With .Selection.Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                    
                    .Range(xlColValue & (Idx & ":") & xlColValue & Idx).Select
                    With .Selection
                        .NumberFormat = gs_formatPriceIDR
                    End With
                    
    '--- Sheet 2 Components from INDONESIA
                    ElseIf FGCls = "02" Then
                    
                        .Sheets("Sheet1").Select
                        .Sheets("Sheet1").Name = "C"
                        
                        .Range(xlColA & "4:" & xlColDescription & "4").Merge
                        .Range(xlColA & "4") = "C. Components from INDONESIA"
                        .Range(xlColA & "4").HorizontalAlignment = xlLeft
                        .Range(xlColA & "4").VerticalAlignment = xlCenter
                
                        .Range(xlColNo & "5") = "No."
                        .Range(xlColHSNumber & "5") = "HS Number"
                        .Range(xlColItemCode & "5") = "Item Code"
                        .Range(xlColDescription & "5") = "Description"
                        .Range(xlColRegion & "5") = "Region"
                        .Range(xlColOriginCountry & "5") = "Origin Country"
                        .Range(xlColInvoiceNumber & "5") = "Invoice Number"
                        .Range(xlColDate & "5") = "Date"
                        .Range(xlColValue & "5") = "Value (USD)"
                        .Range(xlColPercent & "5") = "%"
                        
                        
                        .Range(xlColHSNumber & "7") = "Total"
                        .Range(xlColHSNumber & "7").Font.Bold = True
                        .Range(xlColHSNumber & "7:" & xlColPercent & "7").Interior.ColorIndex = 6
                        .Range(xlColValue & "7").Font.Bold = True
                        
                        
                        .Range(xlColNo & "5:" & xlColPercent & "5").Select
                        .Range(xlColNo & "5:" & xlColPercent & "5").Font.Bold = True
                        
                        With .Selection
                            .HorizontalAlignment = xlCenter
                            .VerticalAlignment = xlCenter
                        End With
                        
                        Idx = 7
                        
                        .Range(xlColNo & (Idx & ":") & xlColPercent & Idx).Select
                        .Range(xlColNo & (Idx & ":") & xlColPercent & Idx).Font.Bold = True
                        
                        With .Selection
                            .HorizontalAlignment = xlCenter
                            .VerticalAlignment = xlCenter
                        End With
                        
                        TglEnd1 = DTEffective
                        TglEnd1 = Format(TglEnd1, "mm/dd/yyyy")
                        'Region = "01"
                        
                                                                        
                        strSQL = " Declare @TglEnd as date=(SELECT DATEADD(s,-1,DATEADD(M, DATEDIFF(MONTH,0,'" & TglEnd1 & "')+1,0))) " & vbCrLf & _
                             " Select * From Bom_Cost a left join Region_Cls b on a.Region=b.Description " & vbCrLf & _
                             " where Parent_ItemCode='" & txtItemCode & "' and Period = @TglEnd and b.Region_Cls='01' "
                                                                                
                        If RsSearch.State <> adStateClosed Then RsSearch.Close
                        RsSearch.CursorLocation = adUseClient
                        RsSearch.Open strSQL, Db, adOpenForwardOnly, adLockOptimistic
                        
                        '#Fill Grid
                        i = 0
                        Do While Not RsSearch.EOF
                             i = i + 1
                            Grid2.AddItem i
                            Grid2.TextMatrix(i, colNo) = Grid2.TextMatrix(i, colNo)
                            Grid2.TextMatrix(i, ColHSNumber) = "'" + Trim(RsSearch!HS_Code)
                            Grid2.TextMatrix(i, ColItemCode) = Trim(RsSearch!Item_Code)
                            Grid2.TextMatrix(i, ColDescription) = Trim(RsSearch!item_name)
                            Grid2.TextMatrix(i, ColRegion) = IIf(IsNull(Trim(RsSearch("Region"))), "", Trim(RsSearch("Region")))
                            Grid2.TextMatrix(i, ColOriginCountry) = IIf(IsNull(Trim(RsSearch("Origin_Country"))), "", Trim(RsSearch("Origin_Country")))
                            Grid2.TextMatrix(i, ColInvoiceNumber) = "'" + Trim(RsSearch!BC40_No)
                            Grid2.TextMatrix(i, ColDate) = Format(RsSearch!BC40_Date, "dd MMM yyyy")
                            Grid2.TextMatrix(i, ColValue) = Trim(RsSearch!Value)
                            Grid2.TextMatrix(i, ColPercent) = Format(RsSearch!Percentage, gs_formatPercentage)
                            
                        RsSearch.MoveNext
                        Loop
                        
                        For i = 1 To Grid2.Rows - 1
                            Idx = Idx + 1
                            .Range(xlColNo & Idx) = Grid2.TextMatrix(i, colNo)
                            .Range(xlColHSNumber & Idx) = Grid2.TextMatrix(i, ColHSNumber)
                            .Range(xlColItemCode & Idx) = Grid2.TextMatrix(i, ColItemCode)
                            .Range(xlColDescription & Idx) = Grid2.TextMatrix(i, ColDescription)
                            .Range(xlColRegion & Idx) = Grid2.TextMatrix(i, ColRegion)
                            .Range(xlColOriginCountry & Idx) = Grid2.TextMatrix(i, ColOriginCountry)
                            .Range(xlColInvoiceNumber & Idx) = Grid2.TextMatrix(i, ColInvoiceNumber)
                            .Range(xlColDate & Idx) = Grid2.TextMatrix(i, ColDate)
                            .Range(xlColValue & Idx) = Grid2.TextMatrix(i, ColValue)
                            .Range(xlColPercent & Idx) = Grid2.TextMatrix(i, ColPercent)
                        Next
                        
                        .Columns(xlColA & ":" & xlColNo).ColumnWidth = 2.57
                        .Columns(xlColNo & ":" & xlColNo).ColumnWidth = 5
                        .Columns(xlColHSNumber & ":" & xlColHSNumber).ColumnWidth = 13
                        .Columns(xlColItemCode & ":" & xlColItemCode).ColumnWidth = 13
                        .Columns(xlColDescription & ":" & xlColDescription).ColumnWidth = 62
                        .Columns(xlColRegion & ":" & xlColRegion).ColumnWidth = 16
                        .Columns(xlColOriginCountry & ":" & xlColOriginCountry).ColumnWidth = 16
                        .Columns(xlColInvoiceNumber & ":" & xlColInvoiceNumber).ColumnWidth = 16
                        .Columns(xlColDate & ":" & xlColDate).ColumnWidth = 11
                        .Columns(xlColValue & ":" & xlColValue).ColumnWidth = 13
                        .Columns(xlColPercent & ":" & xlColPercent).ColumnWidth = 9
                      
                        .Range(xlColNo & "5:" & xlColPercent & Idx).Select
                        
                        If i = 1 Then
                            .Range(xlColValue & "7").Formula = 0
                        Else
                            .Range(xlColValue & "7").Formula = "=Sum(" & xlColValue & "8:" & xlColValue & Idx & ")"
                        End If
                        
                        If i = 1 Then
                            .Range(xlColValue & "7").HorizontalAlignment = xlRight
                            .Range(xlColValue & "7").Formula = 0
                            .Range(xlColPercent & "7:" & xlColPercent & Idx).NumberFormat = "0.00 %"
                        Else
                            .Range(xlColValue & "7").HorizontalAlignment = xlRight
                            .Range(xlColPercent & "7").Formula = "=Sum(" & xlColPercent & "8:" & xlColPercent & Idx & ")"
                            .Range(xlColPercent & "7:" & xlColPercent & Idx).NumberFormat = "0.00 %"
                        End If
                        
                        .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
                        .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
                        

                        With .Selection.Borders(xlEdgeLeft)
                            .LineStyle = xlContinuous
                            .Weight = xlThin
                        End With
                        With .Selection.Borders(xlEdgeTop)
                            .LineStyle = xlContinuous
                            .Weight = xlThin
                        End With
                        With .Selection.Borders(xlEdgeBottom)
                            .LineStyle = xlContinuous
                            .Weight = xlThin
                        End With
                        With .Selection.Borders(xlEdgeRight)
                            .LineStyle = xlContinuous
                            .Weight = xlThin
                        End With
                        With .Selection.Borders(xlInsideVertical)
                            .LineStyle = xlContinuous
                            .Weight = xlThin
                        End With
                        With .Selection.Borders(xlInsideHorizontal)
                            .LineStyle = xlContinuous
                            .Weight = xlThin
                        End With
                        
                        .Range(xlColValue & "7:" & xlColValue & Idx).Select
                        With .Selection
                            .NumberFormat = gs_formatPriceIDR
                        End With
                        
                        Call up_Header2
                        
                        .Sheets("CS").Select
                        .Range(xlColValue & IdxC).Formula = "='C'!J7"
                        .Range(xlColPercent & IdxC).Formula = "='C'!K7"
                        
                    ElseIf FGCls = "03" Then
'                        .Sheets("Sheet2").Select
'                        .Sheets("Sheet2").Name = "BOM"
'
'                        .Range(xlColA & "4:" & xlColDescription & "4").Merge
'                        .Range(xlColA & "4") = "BOM : "
'                        .Range(xlColA & "4").HorizontalAlignment = xlLeft
'                        .Range(xlColA & "4").VerticalAlignment = xlCenter
'
'                        .Range(xlColA & "6") = "HS CODE"
'                        .Range(xlColNo & "6") = "Code"
'                        .Range(xlColHSNumber & "6") = "Part Name"
'                        .Range(xlColDescription & "6") = "Qty"
'                        .Range(xlColOriginCountry & "6") = "Unit"
'                        .Range(xlColInvoiceNumber & "6") = "Vendor" '
'                        .Range(xlColDate & "6") = "Origin"
'                        .Range(xlColValue & "6") = "Unit Price"
'                        .Range(xlColPercent & "6") = " " '
'                        .Range(xlcolUSD & "6") = "USD" '
'                        .Range(xlColBCNo & "6") = "BC No" '
'                        .Range(xlColBCDate & "6") = "BC Date"
'                        .Range(xlColAmount & "6") = "Amount"
'
'                        .Range(xlColHSNumber & "7") = "Total"
'                        .Range(xlColHSNumber & "7").Font.Bold = True
'                        .Range(xlColHSNumber & "7").Font.Bold = True
'                        '.Range(xlColHSNumber & "7:" & xlColPercent & "7").Interior.ColorIndex = 6
'
'                        Idx = 7
'
'                        '#Fill Grid
''                        For i = 2 To Grid.Rows - 1
''                            Idx = Idx + 1
''                            .Range(xlColA & Idx) = Grid.TextMatrix(i, ColA)
''                            .Range(xlColNo & Idx) = Grid.TextMatrix(i, ColNo)
''                            .Range(xlColHSNumber & Idx) = Grid.TextMatrix(i, ColHSNumber)
''                            .Range(xlColDescription & Idx) = Grid.TextMatrix(i, ColDescription)
''                            .Range(xlColOriginCountry & Idx) = Grid.TextMatrix(i, ColOriginCountry)
''                            .Range(xlColInvoiceNumber & Idx) = Grid.TextMatrix(i, ColInvoiceNumber)
''                            .Range(xlColDate & Idx) = Grid.TextMatrix(i, ColDate)
''                            .Range(xlColValue & Idx) = Grid.TextMatrix(i, ColValue)
''                            .Range(xlColPercent & Idx) = Grid.TextMatrix(i, ColPercent)
''                            .Range(xlcolUSD & Idx) = Grid.TextMatrix(i, ColUSD)
''                            .Range(xlColBCNo & Idx) = Grid.TextMatrix(i, ColBCNo)
''                            .Range(xlColBCDate & Idx) = Grid.TextMatrix(i, ColBCDate)
''                            .Range(xlColAmount & Idx) = Grid.TextMatrix(i, ColAmount)
''
''                        Next
'
'                        .Columns(xlColA & ":" & xlColNo).ColumnWidth = 10
'                        .Columns(xlColNo & ":" & xlColNo).ColumnWidth = 8.86
'                        .Columns(xlColHSNumber & ":" & xlColHSNumber).ColumnWidth = 68
'                        .Columns(xlColDescription & ":" & xlColDescription).ColumnWidth = 15.43
'                        .Columns(xlColOriginCountry & ":" & xlColOriginCountry).ColumnWidth = 8.43
'                        .Columns(xlColInvoiceNumber & ":" & xlColInvoiceNumber).ColumnWidth = 20
'                        .Columns(xlColDate & ":" & xlColDate).ColumnWidth = 11
'                        .Columns(xlColValue & ":" & xlColValue).ColumnWidth = 13 '--
'                        .Columns(xlColPercent & ":" & xlColPercent).ColumnWidth = 8.43
'                        .Columns(xlcolUSD & ":" & xlcolUSD).ColumnWidth = 11
'                        .Columns(xlColBCNo & ":" & xlColBCNo).ColumnWidth = 7.29
'                        .Columns(xlColBCDate & ":" & xlColBCDate).ColumnWidth = 9.43
'                        .Columns(xlColAmount & ":" & xlColAmount).ColumnWidth = 8.45
'
'                        .Range(xlColA & "6:" & xlColAmount & Idx).Select
'
'                        .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
'                        .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
'
'
'                        With .Selection.Borders(xlEdgeLeft)
'                            .LineStyle = xlContinuous
'                            .Weight = xlThin
'                        End With
'                        With .Selection.Borders(xlEdgeTop)
'                            .LineStyle = xlContinuous
'                            .Weight = xlThin
'                        End With
'                        With .Selection.Borders(xlEdgeBottom)
'                            .LineStyle = xlContinuous
'                            .Weight = xlThin
'                        End With
'                        With .Selection.Borders(xlEdgeRight)
'                            .LineStyle = xlContinuous
'                            .Weight = xlThin
'                        End With
'                        With .Selection.Borders(xlInsideVertical)
'                            .LineStyle = xlContinuous
'                            .Weight = xlThin
'                        End With
'                        With .Selection.Borders(xlInsideHorizontal)
'                            .LineStyle = xlContinuous
'                            .Weight = xlThin
'                        End With
                    End If
                
               
                If FGCls = "03" Then
                    .Sheets("CS").Select
                    .Visible = True
                ElseIf FGCls = "02" Then
                    FGCls = "03"
                ElseIf FGCls = "01" Then
                    FGCls = "02"
                End If

                .WindowState = xlMaximized
                .ActiveWindow.Zoom = 80
                End With
        Next X
        
   
    Else
        LblErrMsg.Caption = DisplayMsg("8012")
    End If
    
    LblErrMsg.Caption = DisplayMsg("9008")
    Me.MousePointer = vbDefault
    
End Sub

Private Sub Cmd_SubMenu_Click()
    DoEvents
    frmMainMenu.Show
    DoEvents
    Unload Me
End Sub

Private Sub cmdSearch_Click()
    Dim TglEnd1 As String, strSQL As String, ItemCode As String, sql As String
    Dim RsSearch As New ADODB.Recordset, RsInput As New ADODB.Recordset
    
    'With Grid
        LblErrMsg.Caption = ""
        LblRecord.Caption = 0
        
        Call up_Header
        
        Me.MousePointer = vbHourglass
        
        Db.BeginTrans
        
        TglEnd1 = DTEffective
        TglEnd1 = Format(TglEnd1, "mm/dd/yyyy")
        
        strSQL = "EXEC sp_BOMCostReport_Sel '" & txtItemCode & "', '" & TglEnd1 & "'"
        If RsSearch.State = adStateOpen Then RsSearch.Close
        
        RsSearch.CursorLocation = adUseClient
        If RsSearch.State <> adStateClosed Then RsSearch.Close
        RsSearch.Open strSQL, Db, adOpenKeyset, adLockOptimistic
        
'
'        strSQL = "  Declare @TglEnd as date=(SELECT DATEADD(s,-1,DATEADD(M, DATEDIFF(MONTH,0,'" & TglEnd1 & "')+1,0))) " & vbCrLf & _
'                 "  Declare @Total as numeric (18,2)  " & vbCrLf & _
'                              "             set @Total= ( " & vbCrLf & _
'                              "                         SELECT  " & vbCrLf & _
'                              "                         SUM (V.QTY*E.Price) Value " & vbCrLf & _
'                              "  " & vbCrLf & _
'                              "                         FROM [vBOMRecursif_GetDate] V  " & vbCrLf & _
'                              "                                  left join Item_Master B ON V.Item_Code=B.Item_Code " & vbCrLf & _
'                              "                                  left join ( " & vbCrLf & _
'                              "                                             SELECT D.Supplier_Code,C.Item_Code,C.Seq_No,BC_Type,BC40_No,BC40_Date,Currency_Code,Price FROM( " & vbCrLf & _
'                              "                                             Select item_code,Seq_No=max(Seq_No) From Part_Receipt  " & vbCrLf & _
'                              "                                                                     where  Receipt_Cls='R' AND Receipt_Date<=@TglEnd   " & vbCrLf & _
'                              "                                                                     --AND Receipt_Date<=@TglEnd AND Warehouse_Code='WH-001'  " & vbCrLf & _
'                              "                                                 group by Item_Code " & vbCrLf & _
'                              "                                                 )C LEFT JOIN Part_Receipt D ON C.Seq_No=D.Seq_No AND C.Item_Code=D.Item_Code " & vbCrLf & _
'                              "                                             )E ON V.Item_Code=E.Item_Code " & vbCrLf & _
'                              "                                  LEFT JOIN Trade_Master F ON E.Supplier_Code=F.Trade_Code " & vbCrLf & _
'                              "                                  LEFT JOIN Region_Cls G on G.Region_Cls= F.Region_Cls  " & vbCrLf & _
'                              "  " & vbCrLf & _
'                              "                         WHERE ROOT_ITEMCODE=RTRIM('" & txtItemCode & "') AND Coalesce(BC_Type,'')<>'' " & vbCrLf & _
'                              "                         ) "
'
'            strSQL = strSQL + " SELECT Distinct ('" & txtItemCode & "') Parent_ItemCode  " & vbCrLf & _
'                              "      ,@TglEnd Period, E.Supplier_Code,V.Item_Code,B.HS_Code as HS_Code,B.Item_Name as Item_Name,BC_Type,BC40_No,BC40_Date,  " & vbCrLf & _
'                              " CASE " & vbCrLf & _
'                              "  WHEN E.Currency_Code <> '02' then" & vbCrLf & _
'                              " (V.QTY*PRICE/(select TOP 1 Daily_ExchangeRate From Daily_ExchangeRate Where Currency_Code='02' ORDER BY ExchangeRate_Date DESC))" & vbCrLf & _
'                              " Else (V.QTY*PRICE) " & vbCrLf & _
'                              " End As Value, " & vbCrLf & _
'                              "    g.Description as Region, F.Country as Origin_Country, Round ((V.QTY*PRICE)/@Total, 2)  as Percentage " & vbCrLf & _
'                              "  " & vbCrLf & _
'                              " FROM [vBOMRecursif_GetDate] V   " & vbCrLf & _
'                              "          left join Item_Master B ON V.Item_Code=B.Item_Code " & vbCrLf & _
'                              "          left join ( " & vbCrLf & _
'                              "                     SELECT D.Receipt_Date, D.Supplier_Code,C.Item_Code,C.Seq_No,BC_Type,BC40_No,BC40_Date,Currency_Code,Price, D.Warehouse_Code FROM( " & vbCrLf & _
'                              "                     Select item_code,Seq_No=max(Seq_No) From Part_Receipt  " & vbCrLf & _
'                              "                                             where  Receipt_Cls='R' AND Receipt_Date<=@TglEnd AND Warehouse_Code='WH-001'   "
'
'            strSQL = strSQL + "                       group by Item_Code " & vbCrLf & _
'                              "                         )C LEFT JOIN Part_Receipt D ON C.Seq_No=D.Seq_No AND C.Item_Code=D.Item_Code " & vbCrLf & _
'                              "                      " & vbCrLf & _
'                              "                     )E ON V.Item_Code=E.Item_Code " & vbCrLf & _
'                              "          LEFT JOIN Trade_Master F ON E.Supplier_Code=F.Trade_Code  " & vbCrLf & _
'                              "          LEFT JOIN Region_Cls G on G.Region_Cls= F.Region_Cls  " & vbCrLf & _
'                              "  " & vbCrLf & _
'                              " WHERE ROOT_ITEMCODE=RTRIM('" & txtItemCode & "') AND Coalesce(BC_Type,'')<>'' AND Receipt_Date <= @TglEnd AND FinishGoodPart_Cls <> '01'" & vbCrLf & _
'                              " ORDER BY BC_Type "
'
'       If RsSearch.State <> adStateClosed Then RsSearch.Close
'        RsSearch.CursorLocation = adUseClient
'        RsSearch.Open strSQL, Db, adOpenForwardOnly, adLockOptimistic
        
        If RsSearch.EOF = False Then
        
            'Cek ParentItemCode
            strSQL = " Declare @TglEnd as date=(SELECT DATEADD(s,-1,DATEADD(M, DATEDIFF(MONTH,0,'" & TglEnd1 & "')+1,0))) " & vbCrLf & _
                    " Select * From Bom_Cost where Parent_ItemCode='" & txtItemCode & "' and Period = @TglEnd "
            
            If RsInput.State <> adStateClosed Then RsInput.Close
            RsInput.Open strSQL, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
            
            If RsInput.EOF Then
              
              Db.CommitTrans
              
              Do While Not RsSearch.EOF
                
                strSQL = " Insert into Bom_Cost (Parent_ItemCode, Period, Supplier_Code, Item_Code, HS_Code, Item_Name, BC_Type, BC40_No, " & vbCrLf & _
                         " BC40_Date, Value, Region, Origin_Country, Percentage, Last_User, Register_Date)" & vbCrLf & _
                         " values('" & Trim(RsSearch!parent_itemcode) & "','" & Trim(RsSearch!Period) & "','" & Trim(RsSearch!Supplier_Code) & "', " & vbCrLf & _
                         " '" & Trim(RsSearch!Item_Code) & "','" & Trim(RsSearch!HS_Code) & "', '" & Trim(RsSearch!item_name) & "','" & Trim(RsSearch!BC_Type) & "'," & vbCrLf & _
                         "'" & Trim(RsSearch!BC40_No) & "','" & Trim(RsSearch!BC40_Date) & "','" & Trim(RsSearch!Value) & "','" & Trim(RsSearch!Region) & "', '" & Trim(RsSearch!Origin_Country) & "', " & vbCrLf & _
                         "   '" & Trim(RsSearch!Percentage) & "','" & userLogin & "', getdate())"
                Db.Execute strSQL
                 RsSearch.MoveNext
                Loop
            Else
                Db.CommitTrans
                
               
                strSQL = " Declare @TglEnd as date=(SELECT DATEADD(s,-1,DATEADD(M, DATEDIFF(MONTH,0,'" & TglEnd1 & "')+1,0))) " & vbCrLf & _
                         " Delete From Bom_Cost Where Parent_ItemCode= '" & txtItemCode & "' and Period = @TglEnd   "
                
                Db.Execute strSQL
                
                Do While Not RsSearch.EOF
                    strSQL = " Insert into Bom_Cost (Parent_ItemCode, Period, Supplier_Code, Item_Code, HS_Code, Item_Name, BC_Type, BC40_No, " & vbCrLf & _
                         " BC40_Date, Value, Region, Origin_Country, Percentage, Last_User, Register_Date)" & vbCrLf & _
                         " values('" & Trim(RsSearch!parent_itemcode) & "','" & Trim(RsSearch!Period) & "','" & Trim(RsSearch!Supplier_Code) & "', " & vbCrLf & _
                         " '" & Trim(RsSearch!Item_Code) & "','" & Trim(RsSearch!HS_Code) & "', '" & Trim(RsSearch!item_name) & "','" & Trim(RsSearch!BC_Type) & "'," & vbCrLf & _
                         " '" & Trim(RsSearch!BC40_No) & "','" & Trim(RsSearch!BC40_Date) & "','" & Trim(RsSearch!Value) & "','" & Trim(RsSearch!Region) & "', '" & Trim(RsSearch!Origin_Country) & "', " & vbCrLf & _
                         "   '" & Trim(RsSearch!Percentage) & "','" & userLogin & "', getdate())"
                    Db.Execute strSQL
                    RsSearch.MoveNext
                Loop
            End If
             Set RsInput = Nothing
            End If
                  
        strSQL = " Declare @TglEnd as date=(SELECT DATEADD(s,-1,DATEADD(M, DATEDIFF(MONTH,0,'" & TglEnd1 & "')+1,0))) " & vbCrLf & _
                 " Select * From Bom_Cost where Parent_ItemCode='" & txtItemCode & "' and Period = @TglEnd order by Item_Code "
        
        If RsSearch.State <> adStateClosed Then RsSearch.Close
        RsSearch.CursorLocation = adUseClient
        RsSearch.Open strSQL, Db, adOpenForwardOnly, adLockOptimistic
        
        i = 0
        If RsSearch.EOF = False Then
        
        Do While Not RsSearch.EOF
        '" & Trim(RsSearch!Region) & "'
            i = i + 1
            Grid.AddItem i
            Grid.TextMatrix(i, ColHSNumber) = Trim(RsSearch!HS_Code)
            Grid.TextMatrix(i, ColItemCode) = Trim(RsSearch!Item_Code)
            Grid.TextMatrix(i, ColDescription) = Trim(RsSearch!item_name)
            Grid.TextMatrix(i, ColRegion) = IIf(IsNull(Trim(RsSearch("Region"))), "", Trim(RsSearch("Region")))
            Grid.TextMatrix(i, ColOriginCountry) = IIf(IsNull(Trim(RsSearch("Origin_Country"))), "", Trim(RsSearch("Origin_Country")))
            Grid.TextMatrix(i, ColInvoiceNumber) = Trim(RsSearch!BC40_No)
            Grid.TextMatrix(i, ColDate) = Format(RsSearch!BC40_Date, "dd MMM yyyy")
            Grid.TextMatrix(i, ColValue) = Trim(RsSearch!Value)
            Grid.TextMatrix(i, ColPercent) = Format(RsSearch!Percentage, gs_formatPercentage)
            
            RsSearch.MoveNext
            
        Loop
        
        If RsSearch.RecordCount >= 0 Then
            LblRecord = Format(RsSearch.RecordCount, "#,##0 Record")
        Else
            LblRecord = Format("0", "#,##0")
        End If
        
        Else
            LblErrMsg.Caption = DisplayMsg("0013")
        End If
                
        If RsSearch.State <> adStateClosed Then RsSearch.Close
        
        Me.MousePointer = vbDefault

End Sub

Private Sub cmd_clear_Click()
    Call up_Header
    Call up_Header2
    DTEffective.Value = Now
    LblRecord = "0 Record"
    LblErrMsg.Caption = ""
End Sub


Private Sub txtItemCode_Change()
Dim sql As String
Dim RS As New Recordset
Dim cmd As ADODB.Command

    
    sql = "Select * From item_master where FinishGoodPart_Cls='01' and Item_Code='" & txtItemCode.Text & "'"
    Set RS = Db.Execute(sql)
    
    If RS.EOF = False Then
        txtItemName(1).Caption = Trim(RS("Item_Name"))
    Else
        txtItemName(1).Caption = ""
        Exit Sub
    End If
    
    If txtItemCode.Text = "" Then txtItemName(1).Caption = ""
    
End Sub
