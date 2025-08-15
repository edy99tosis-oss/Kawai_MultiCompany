VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_pi_update 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Physical Inventory Update"
   ClientHeight    =   10980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15120
   Icon            =   "frm_pi_update.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10980
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtSearch 
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
      Left            =   3210
      MaxLength       =   25
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   8550
      Width           =   2430
   End
   Begin VB.CommandButton cmdSearch 
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
      Left            =   5715
      Style           =   1  'Graphical
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   8520
      Width           =   1125
   End
   Begin VB.TextBox txtitemcode 
      Height          =   240
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.CommandButton cmdScanBarcode 
      BackColor       =   &H0080FFFF&
      Caption         =   "Scan Barcode [F2]"
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
      Left            =   10305
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   9855
      Width           =   2130
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   420
      Left            =   12825
      TabIndex        =   19
      Top             =   360
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   741
   End
   Begin VB.CommandButton Cmd_Save 
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
      Index           =   1
      Left            =   12555
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9825
      Width           =   1035
   End
   Begin VB.CommandButton Cmd_Save 
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
      Index           =   0
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9825
      Width           =   1035
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   345
      TabIndex        =   17
      Top             =   9015
      Width           =   14355
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
         Left            =   135
         TabIndex        =   18
         Top             =   210
         Width           =   14055
      End
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
      Left            =   2610
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   9825
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
      Left            =   3930
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   9825
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
      Left            =   5220
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   9825
      Visible         =   0   'False
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
      Left            =   6540
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   9825
      Visible         =   0   'False
      Width           =   1125
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9825
      Width           =   1125
   End
   Begin VB.CommandButton Cmd_Save 
      BackColor       =   &H0080FFFF&
      Caption         =   "Searc&h"
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
      Left            =   3870
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1545
      Width           =   1035
   End
   Begin MSComCtl2.DTPicker DMonth 
      Height          =   315
      Left            =   2400
      TabIndex        =   1
      Top             =   1575
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
      Height          =   6285
      Left            =   345
      TabIndex        =   10
      Top             =   2070
      Width           =   14355
      _cx             =   25321
      _cy             =   11086
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
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search"
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
      Index           =   26
      Left            =   360
      TabIndex        =   25
      Top             =   8610
      Width           =   600
   End
   Begin MSForms.ComboBox cboSearch 
      Height          =   315
      Left            =   1050
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   8550
      Width           =   2085
      VariousPropertyBits=   612386843
      MaxLength       =   6
      DisplayStyle    =   7
      Size            =   "3678;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox CboLocationCD 
      Height          =   315
      Left            =   2385
      TabIndex        =   0
      Top             =   1125
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
      Left            =   675
      TabIndex        =   16
      Top             =   1155
      Width           =   1335
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
      Left            =   675
      TabIndex        =   15
      Top             =   1605
      Width           =   1335
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
      Left            =   5085
      TabIndex        =   14
      Top             =   1155
      Width           =   1605
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
      Left            =   6765
      TabIndex        =   13
      Top             =   1155
      Width           =   3015
   End
   Begin VB.Line Line1 
      X1              =   6765
      X2              =   9780
      Y1              =   1380
      Y2              =   1380
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Physical Inventory Update"
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
      Left            =   360
      TabIndex        =   12
      Top             =   360
      Width           =   14310
   End
   Begin VB.Label LblPesan 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "rere"
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
      Left            =   450
      TabIndex        =   11
      Top             =   9330
      Width           =   11940
   End
End
Attribute VB_Name = "frm_pi_update"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_upd As Double, dateUp As Date
Dim bteColProdCod As Byte
Dim bteColDesc As Byte
Dim bteColUnit As Byte
Dim bteColAddress As Byte
Dim bteColPreMonth As Byte
Dim bteColReceipt As Byte
Dim bteColSupply As Byte
Dim bteColLossReject As Byte
Dim bteColEnd As Byte
Dim bteColDiff As Byte
Dim bteColInventory As Byte
Dim bteColInv2 As Byte
Dim bteColReason As Byte
Dim bteColReason2 As Byte

Dim bytSort As Byte

Private Sub Header()
    bteColProdCod = 0
    bteColDesc = 1
    bteColUnit = 2
    bteColAddress = 3
    bteColPreMonth = 4
    bteColReceipt = 5
    bteColSupply = 6
    bteColLossReject = 7
    bteColEnd = 8
    bteColDiff = 9
    bteColInventory = 10
    bteColInv2 = 11
    bteColReason = 12
    bteColReason2 = 13
    
    grid.clear
    grid.Rows = 1
    grid.ColS = 14
    
    grid.TextMatrix(0, bteColProdCod) = "Product Code"
    grid.TextMatrix(0, bteColDesc) = "Description"
    grid.TextMatrix(0, bteColAddress) = "Address"
    grid.TextMatrix(0, bteColPreMonth) = "Pre Month Stock"
    grid.TextMatrix(0, bteColReceipt) = "Receipt Total"
    grid.TextMatrix(0, bteColSupply) = "Supply Total"
    grid.TextMatrix(0, bteColLossReject) = "Loss/Reject"
    grid.TextMatrix(0, bteColEnd) = "End of Month Stock"
    grid.TextMatrix(0, bteColUnit) = "Unit"
    grid.TextMatrix(0, bteColDiff) = "Differences"
    grid.TextMatrix(0, bteColInventory) = "Inventory"
    grid.TextMatrix(0, bteColInv2) = "Inv2"
    grid.TextMatrix(0, bteColReason) = "Reason"
    grid.TextMatrix(0, bteColReason2) = "Reason2"
    
    grid.ColWidth(bteColProdCod) = 2000
    grid.ColWidth(bteColDesc) = 2700
    grid.ColWidth(bteColAddress) = 800
    grid.ColWidth(bteColPreMonth) = 1500
    grid.ColWidth(bteColReceipt) = 1500
    grid.ColWidth(bteColSupply) = 1500
    grid.ColWidth(bteColLossReject) = 1400
    grid.ColWidth(bteColEnd) = 1900
    grid.ColWidth(bteColDiff) = 1300
    grid.ColWidth(bteColUnit) = 700
    grid.ColWidth(bteColInventory) = 1300
    grid.ColWidth(bteColInv2) = 1300
    grid.ColWidth(bteColReason) = 3500
    grid.ColWidth(bteColReason2) = 3500
    
    grid.ColAlignment(bteColProdCod) = flexAlignLeftCenter
    grid.ColAlignment(bteColDesc) = flexAlignLeftCenter
    grid.ColAlignment(bteColAddress) = flexAlignLeftCenter
    grid.ColAlignment(bteColPreMonth) = flexAlignRightCenter
    grid.ColAlignment(bteColReceipt) = flexAlignRightCenter
    grid.ColAlignment(bteColSupply) = flexAlignRightCenter
    grid.ColAlignment(bteColLossReject) = flexAlignRightCenter
    grid.ColAlignment(bteColEnd) = flexAlignRightCenter
    grid.ColAlignment(bteColDiff) = flexAlignRightCenter
    grid.ColAlignment(bteColUnit) = flexAlignLeftCenter
    grid.ColAlignment(bteColInventory) = flexAlignRightCenter
    grid.ColAlignment(bteColInv2) = flexAlignRightCenter
    grid.ColAlignment(bteColReason) = flexAlignLeftCenter
    grid.ColAlignment(bteColReason2) = flexAlignLeftCenter
    
    grid.Cell(flexcpAlignment, 0, 0, 0, bteColReason2) = flexAlignCenterCenter
    
    grid.ColHidden(bteColInv2) = True
    grid.ColHidden(bteColReason2) = True
    grid.ColFormat(bteColInventory) = gs_formatQty
    grid.ColFormat(bteColInv2) = gs_formatQty
    
    grid.FrozenCols = bteColPreMonth
End Sub

Private Sub CboLocationCD_Change()
Call clearGrid
If CboLocationCD.MatchFound Then
   LblLocationName = CboLocationCD.List(CboLocationCD.ListIndex, 1)
   LblErrMsg = ""
   'Call Browse
Else
   LblLocationName = ""
   LblErrMsg = DisplayMsg(4018) '"Invalid warehouse code !"
End If
End Sub

Private Sub CboLocationCD_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
Dim j  As Integer

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
    LblErrMsg = DisplayMsg(4018) '"Invalid warehouse code !": Exit Sub
Else
    LblErrMsg = ""
End If
End If
End Sub

Private Sub CboLocationCD_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Cmd_Save_Click(Index As Integer)
Dim i As Integer, j As Integer, Vdif As Double
Select Case Index
      Case 1:
            For i = 1 To grid.Rows - 1
                 If Trim(grid.TextMatrix(i, bteColInventory)) <> Trim(grid.TextMatrix(i, bteColInv2)) Then
                   grid.TextMatrix(i, bteColInventory) = Trim(grid.TextMatrix(i, bteColInv2))
                 ElseIf Trim(grid.TextMatrix(i, bteColReason)) <> Trim(grid.TextMatrix(i, bteColReason2)) Then
                   grid.TextMatrix(i, bteColReason) = Trim(grid.TextMatrix(i, bteColReason2))
                 End If
                 If grid.TextMatrix(i, bteColInventory) <> "" Then
                    Vdif = CDbl(grid.TextMatrix(i, bteColEnd)) - CDbl(IIf(grid.TextMatrix(i, bteColInventory) = "", 0, grid.TextMatrix(i, bteColInventory)))
                    grid.TextMatrix(i, bteColDiff) = Format(Vdif, gs_formatQty)
                Else
                    grid.TextMatrix(i, bteColDiff) = ""
                End If
            Next
            LblErrMsg = ""
       Case 8:
                frmMainMenu.Show
                Unload Me
       Case 9:
                If CboLocationCD.Text = "" Then
                   LblErrMsg = DisplayMsg(1042) '"Please choose warehouse !"
                Else
                    Me.MousePointer = vbHourglass
                        LblErrMsg = ""
                       Call Header
                       grid.Rows = 1
                       Call qBrowse
    
                    Me.MousePointer = vbDefault
                End If
        Case 0:
                If hakUpdate(Me.Name) = 0 Then _
                LblErrMsg = DisplayMsg(3008): Exit Sub
                
                Dim sqlControl As String
                
                j = 0
                For i = 0 To CboLocationCD.ListCount - 1
                    If UCase(Trim(CboLocationCD)) = UCase(Trim(CboLocationCD.List(i, 0))) Then
                        CboLocationCD = Trim(CboLocationCD.List(i, 0))
                        LblLocationName = Trim(CboLocationCD.List(i, 1))
                        j = 1: Exit For
                    End If
                Next
                
                If j = 0 Then LblErrMsg = DisplayMsg(4018) '"Invalid warehouse code !": Exit Sub
                        
                '#Validate Period
                Dim l_upd As Double
                LblErrMsg = up_ValidateDateRange(DMonth.Value, True)
                If Trim(LblErrMsg) <> "" Then Exit Sub
                LblErrMsg = ""
                
                Select Case up_GetDateRange(DMonth.Value)
                    Case 0:
                        If hakUpdate(Me.Name) = 0 Then
                           For i = 1 To grid.Rows - 1
                                 If Trim(grid.TextMatrix(i, bteColInventory)) <> Trim(grid.TextMatrix(i, bteColInv2)) Then
                                   grid.TextMatrix(i, bteColInventory) = Trim(grid.TextMatrix(i, bteColInv2))
                                 ElseIf Trim(grid.TextMatrix(i, bteColReason)) <> Trim(grid.TextMatrix(i, bteColReason2)) Then
                                   grid.TextMatrix(i, bteColReason) = Trim(grid.TextMatrix(i, bteColReason2))
                                 End If
                            Next
                            LblErrMsg = DisplayMsg(4046) '"Can't update fixed data !"
                        Else
                           Db.BeginTrans
                            For i = 1 To grid.Rows - 1
                                   If IsNumeric(grid.TextMatrix(i, bteColInventory)) = True And Trim(grid.TextMatrix(i, bteColInventory)) <> "" Then
                                        l_upd = CDbl(grid.TextMatrix(i, bteColInventory))
                                   Else
                                      
                                         If Trim(grid.TextMatrix(i, bteColInventory)) <> Trim(grid.TextMatrix(i, bteColInv2)) Or Trim(grid.TextMatrix(i, bteColReason)) <> Trim(grid.TextMatrix(i, bteColReason2)) Then
                                                Db.Execute "update stock_master with (updlock) " & _
                                                    "set lm_inventory = 0, " & _
                                                    "tm_premonth = 0, " & _
                                                    "tm_current = tm_receipt - tm_supply - tm_lossreject, " & _
                                                    "nm_premonth = tm_receipt -tm_supply - tm_lossreject, " & _
                                                    "nm_current = tm_receipt - tm_supply - tm_lossreject + nm_receipt - nm_supply - nm_lossreject, " & _
                                                    "lm_reason = '" & Trim(grid.TextMatrix(i, bteColReason)) & "', " & _
                                                    "Last_Update = getdate(), Last_User = '" & userLogin & "' " & _
                                                    "where warehouse_code = '" & Trim(CboLocationCD) & "' and item_code = '" & Trim(grid.TextMatrix(i, bteColProdCod)) & "'"
                                                grid.TextMatrix(i, bteColInv2) = Trim(grid.TextMatrix(i, bteColInventory))
                                                grid.TextMatrix(i, bteColReason2) = Trim(grid.TextMatrix(i, bteColReason))
                                        End If
                                        GoTo next0:
                                    End If
                                    
                                If Trim(grid.TextMatrix(i, bteColInventory)) <> Trim(grid.TextMatrix(i, bteColInv2)) Or Trim(grid.TextMatrix(i, bteColReason)) <> Trim(grid.TextMatrix(i, bteColReason2)) Then
                                
                                    Db.Execute "update stock_master with (updlock) " & _
                                        "set lm_inventory =" & l_upd & ", " & _
                                        "tm_premonth =  " & l_upd & ", " & _
                                        "tm_current =  tm_receipt - tm_supply - tm_lossreject  + " & l_upd & ", " & _
                                        "nm_premonth = tm_receipt - tm_supply - tm_lossreject  + " & l_upd & ", " & _
                                        "nm_current = tm_receipt - tm_supply - tm_lossreject + nm_receipt - nm_supply - nm_lossreject + " & l_upd & ", " & _
                                        "lm_reason = '" & Trim(grid.TextMatrix(i, bteColReason)) & "', " & _
                                        "Last_Update = getdate(), Last_User = '" & userLogin & "' " & _
                                        "where warehouse_code='" & Trim(CboLocationCD) & "' and item_code='" & Trim(grid.TextMatrix(i, bteColProdCod)) & "'"
                                                
                                   grid.TextMatrix(i, bteColInv2) = Trim(grid.TextMatrix(i, bteColInventory))
                                   grid.TextMatrix(i, bteColReason2) = Trim(grid.TextMatrix(i, bteColReason))
                                End If
next0:
                            Next
                              LblErrMsg = DisplayMsg(1101) '"Update data success !"
                             Db.CommitTrans
                        End If
                                              
                        Exit Sub
                        
                    Case 1:
                        Db.BeginTrans
                        
                            For i = 1 To grid.Rows - 1
                                   If IsNumeric(grid.TextMatrix(i, bteColInventory)) = True And Trim(grid.TextMatrix(i, bteColInventory)) <> "" Then
                                        l_upd = CDbl(grid.TextMatrix(i, bteColInventory))
                                    Else
                                      
                                         If Trim(grid.TextMatrix(i, bteColInventory)) <> Trim(grid.TextMatrix(i, bteColInv2)) Or Trim(grid.TextMatrix(i, bteColReason)) <> Trim(grid.TextMatrix(i, bteColReason2)) Then
                                            Db.Execute "update stock_master with (updlock) " & _
                                                "set tm_inventory = null, " & _
                                                "tm_reason = '" & Trim(grid.TextMatrix(i, bteColReason)) & "', " & _
                                                "nm_premonth = tm_premonth + tm_receipt -tm_supply - tm_lossreject, " & _
                                                "nm_current = tm_premonth + tm_receipt - tm_supply - tm_lossreject + nm_receipt - nm_supply - nm_lossreject, " & _
                                                "Last_Update = getdate(), Last_User = '" & userLogin & "' " & _
                                                "where warehouse_code='" & Trim(CboLocationCD) & "' and item_code='" & Trim(grid.TextMatrix(i, bteColProdCod)) & "'"
                                            grid.TextMatrix(i, bteColInv2) = Trim(grid.TextMatrix(i, bteColInventory))
                                            grid.TextMatrix(i, bteColReason2) = Trim(grid.TextMatrix(i, bteColReason))
                                        End If
                                        GoTo next1:
                                    End If
                                    
                                If Trim(grid.TextMatrix(i, bteColInventory)) <> Trim(grid.TextMatrix(i, bteColInv2)) Or Trim(grid.TextMatrix(i, bteColReason)) <> Trim(grid.TextMatrix(i, bteColReason2)) Then
                                    Db.Execute "update stock_master with (updlock) " & _
                                        "set tm_inventory =" & l_upd & ", " & _
                                        "tm_reason = '" & Trim(grid.TextMatrix(i, bteColReason)) & "', " & _
                                        "nm_premonth = " & l_upd & " , " & _
                                        "nm_current = " & l_upd & " + nm_receipt - nm_supply - nm_lossreject, " & _
                                        "Last_Update = getdate(), Last_User = '" & userLogin & "' " & _
                                        "where warehouse_code='" & Trim(CboLocationCD) & "' and item_code='" & Trim(grid.TextMatrix(i, bteColProdCod)) & "'"
                                    grid.TextMatrix(i, bteColInv2) = Trim(grid.TextMatrix(i, bteColInventory))
                                    grid.TextMatrix(i, bteColReason2) = Trim(grid.TextMatrix(i, bteColReason))
                                End If
next1:
                            Next
                         
                            LblErrMsg = DisplayMsg(1101) '"Update data success !"
                        
                        Db.CommitTrans
                            
                    Case 2:
                         Db.BeginTrans
                         
                            For i = 1 To grid.Rows - 1
                                    If IsNumeric(grid.TextMatrix(i, bteColInventory)) = True And Trim(grid.TextMatrix(i, bteColInventory)) <> "" Then
                                        l_upd = CDbl(grid.TextMatrix(i, bteColInventory))
                                    Else
                                   
                                        If Trim(grid.TextMatrix(i, bteColInventory)) <> Trim(grid.TextMatrix(i, bteColInv2)) Or Trim(grid.TextMatrix(i, bteColReason)) <> Trim(grid.TextMatrix(i, bteColReason2)) Then
                                            Db.Execute "update stock_master  with (updlock) " & _
                                                "set  nm_inventory = null, " & _
                                                "nm_reason = '" & Trim(grid.TextMatrix(i, bteColReason)) & "', " & _
                                                "Last_Update = getdate(), Last_User = '" & userLogin & "' " & _
                                                "where warehouse_code='" & Trim(CboLocationCD) & "' and item_code='" & Trim(grid.TextMatrix(i, bteColProdCod)) & "'"
                                            grid.TextMatrix(i, bteColInv2) = Trim(grid.TextMatrix(i, bteColInventory))
                                            grid.TextMatrix(i, bteColReason2) = Trim(grid.TextMatrix(i, bteColReason))
                                        End If
                                        GoTo next2:
                                    End If
                                    
                                If Trim(grid.TextMatrix(i, bteColInventory)) <> Trim(grid.TextMatrix(i, bteColInv2)) Or Trim(grid.TextMatrix(i, bteColReason)) <> Trim(grid.TextMatrix(i, bteColReason2)) Then
                                    Db.Execute "update stock_master with (updlock) " & _
                                        "set nm_inventory=" & l_upd & ", " & _
                                        "nm_reason = '" & Trim(grid.TextMatrix(i, bteColReason)) & "', " & _
                                        "Last_Update = getdate(), Last_User = '" & userLogin & "' " & _
                                        "where warehouse_code='" & Trim(CboLocationCD) & "' and item_code='" & Trim(grid.TextMatrix(i, bteColProdCod)) & "'"
                                    grid.TextMatrix(i, bteColInv2) = Trim(grid.TextMatrix(i, bteColInventory))
                                    grid.TextMatrix(i, bteColReason2) = Trim(grid.TextMatrix(i, bteColReason))
                                End If
next2:
                            Next
                           
                            LblErrMsg = DisplayMsg(1101) ' "Update data success !"
                            Db.CommitTrans

                 End Select
'Db.CommitTrans

End Select

End Sub

Private Sub cmdScanBarcode_Click()
    LblErrMsg.Caption = ""
    frm_pi_barcode.Show vbModal
End Sub

Private Sub cmdSearch_Click()
    Dim i As Double
    
    LblErrMsg = ""
    
    If txtSearch = "" Or grid.Rows = 2 Then txtSearch.SetFocus: Exit Sub
    If grid.Row = grid.Rows - 1 Then i = 2 Else i = grid.Row + 1
    
    Do
        Select Case cboSearch.ListIndex
        Case 0
            grid.Col = bteColProdCod
            If UCase(Mid(grid.TextMatrix(i, bteColProdCod), 1, Len(txtSearch))) = UCase(txtSearch) Then
                Exit Do
            End If
        Case 1
            grid.Col = bteColDesc
            If InStr(UCase(grid.TextMatrix(i, bteColDesc)), UCase(txtSearch)) <> 0 Then
                Exit Do
            End If
        End Select
        i = i + 1
        If i = grid.Rows - 1 Then
            txtSearch = ""
            i = 2
            LblErrMsg = DisplayMsg(8012)
            Exit Do
        End If
    Loop
    
    grid.Row = i
    grid.TopRow = i
    grid.Col = bteColInventory
    grid.SetFocus
    SendKeys "{left}"
    SendKeys "{right}"

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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then cmdScanBarcode_Click: KeyCode = 0
    If KeyCode = vbKeyF3 Then
        cmdSearch_Click
    End If
End Sub


Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
LblLocationName = ""
LblErrMsg = ""

DMonth = Format(Date, "MMM yyyy")
dateUp = DMonth.Value

CtrlMenu1.FormName = Me.Name
Me.Caption = "Physical Inventory Update"
Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"

With cboSearch
    .AddItem "Item Code"
    .AddItem "Description"
    .ListIndex = 0
End With

txtSearch = ""

Call StockLocation
DMonth = Format(Now, "mmmm yyyy")
Call Header
End Sub

Sub clearGrid()
grid.clear
grid.Rows = 1
Call Header
End Sub

Private Sub StockLocation()
Dim sql As String, ls_sql As String
Dim RsStock As New ADODB.Recordset
Dim i As Long

If RsStock.State <> adStateClosed Then RsStock.Close
ls_sql = " select * from (select wh_code, wh_name  from warehouse_master where stockcontrol_cls='01' union  " & _
      " select trade_code wh_code, trade_name wh_name from trade_master where trade_code in(select manufacture_code from manufacture_line))tbWarehouse order by wh_code "
RsStock.Open ls_sql, Db, adOpenDynamic, adLockOptimistic, adCmdText

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
Dim Vdif As Double
With grid
  If Col = bteColInventory Then
    If IsNumeric(.TextMatrix(Row, Col)) = True Then _
        If CDbl(.TextMatrix(Row, Col)) > gd_MaxQty Then .TextMatrix(Row, Col) = gd_MaxQty: LblErrMsg = DisplayMsg(4045) & " " & gd_MaxQty: Exit Sub
    If IsNumeric(.TextMatrix(Row, Col)) = False And Trim(.TextMatrix(Row, Col)) <> "" Then
        .TextMatrix(Row, Col) = .TextMatrix(Row, bteColInv2)
        LblErrMsg = DisplayMsg(4044) '"Please input valid quantity !"
    Else
        .TextMatrix(Row, Col) = Format(.TextMatrix(Row, Col), gs_formatQty)
        LblErrMsg = ""
    End If
    If grid.TextMatrix(Row, Col) <> "" Then
        'Vdif = CDbl(Grid.TextMatrix(Row, bteColEnd)) - CDbl(IIf(Grid.TextMatrix(Row, Col) = "", 0, Grid.TextMatrix(Row, Col)))
        'revision after BSRE meeting
        Vdif = CDbl(IIf(grid.TextMatrix(Row, Col) = "", 0, grid.TextMatrix(Row, Col))) - CDbl(grid.TextMatrix(Row, bteColEnd))
        grid.TextMatrix(Row, bteColDiff) = Format(Vdif, gs_formatQty)
    Else
        grid.TextMatrix(Row, bteColDiff) = ""
    End If
    
'    .TextMatrix(Row, bteColDiff) = Format(.TextMatrix(Row, bteColEnd) - .TextMatrix(Row, Col), gs_formatQty)
  ElseIf Col = bteColReason Then
    'Free Text
  End If
End With
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If grid.Col <> bteColInventory And grid.Col <> bteColReason Then
   Cancel = True
End If
If grid.Col = bteColReason Then
 grid.EditMaxLength = 255
Else
 grid.EditMaxLength = 0 'Unlimited
End If
End Sub

Sub qBrowse()

Dim sql As String
Dim RsStock As New ADODB.Recordset
Dim sqlControl As String, l_item_name As String, l_inv As String

Call Header

'#Validate Period
LblErrMsg = up_ValidateDateRange(DMonth.Value, False)
If Trim(LblErrMsg) <> "" Then Exit Sub

If RsStock.State <> adStateClosed Then RsStock.Close

sql = "select sm.*,im.makeritem_code, im.item_name,sheetcoil_cls, " & _
      vbLf & "width,length,thickness,address, " & _
      vbLf & "im.unit_cls, uc.description unitdesc " & _
      vbLf & "from stock_master sm " & _
      vbLf & "inner join item_master im " & _
      vbLf & "on sm.item_code = im.item_code " & _
      vbLf & "left join unit_cls uc " & _
      vbLf & "on uc.unit_cls = im.unit_cls " & _
      vbLf & "where warehouse_code='" & Trim(CboLocationCD) & "' "

RsStock.Open sql, Db, adOpenDynamic, adLockOptimistic
If RsStock.EOF = True And RsStock.BOF = True Then Exit Sub

 RsStock.MoveFirst

Dim Vdif As Double

Select Case up_GetDateRange(DMonth.Value)
Case 0:
    
    While RsStock.EOF = False
        l_item_name = uf_GetItemDescription(Trim(RsStock!Item_Code))
        l_inv = IIf(RsStock!lm_inventory = Null, "", Format(RsStock!lm_inventory, gs_formatQty))
        With grid
            .AddItem ""
            .TextMatrix(.Rows - 1, bteColProdCod) = Trim(RsStock!Item_Code)
            .TextMatrix(.Rows - 1, bteColDesc) = l_item_name
            .TextMatrix(.Rows - 1, bteColAddress) = Trim(RsStock!Address & "")
            .TextMatrix(.Rows - 1, bteColPreMonth) = Format(RsStock!lm_premonth, gs_formatQty)
            .TextMatrix(.Rows - 1, bteColReceipt) = Format(RsStock!lm_receipt, gs_formatQty)
            .TextMatrix(.Rows - 1, bteColSupply) = Format(RsStock!lm_supply, gs_formatQty)
            .TextMatrix(.Rows - 1, bteColLossReject) = Format(RsStock!lm_lossreject, gs_formatQty)
            .TextMatrix(.Rows - 1, bteColEnd) = Format(RsStock!lm_current, gs_formatQty)
            .TextMatrix(.Rows - 1, bteColUnit) = Trim(RsStock!unitdesc)
            .TextMatrix(.Rows - 1, bteColInventory) = l_inv
            If .TextMatrix(.Rows - 1, bteColInventory) <> "" Then
                Vdif = CDbl(IIf(.TextMatrix(.Rows - 1, bteColInventory) = "", 0, .TextMatrix(.Rows - 1, bteColInventory))) - CDbl(.TextMatrix(.Rows - 1, bteColEnd))
                .TextMatrix(.Rows - 1, bteColDiff) = Format(Vdif, gs_formatQty)
            Else
                .TextMatrix(.Rows - 1, bteColDiff) = ""
            End If
            .TextMatrix(.Rows - 1, bteColInv2) = l_inv
            .TextMatrix(.Rows - 1, bteColReason) = IIf(IsNull(RsStock!lm_reason), "", Trim(RsStock!lm_reason))
            .TextMatrix(.Rows - 1, bteColReason2) = IIf(IsNull(RsStock!lm_reason), "", Trim(RsStock!lm_reason))
        End With
        RsStock.MoveNext
    Wend

Case 1:

    While RsStock.EOF = False
        l_item_name = uf_GetItemDescription(Trim(RsStock!Item_Code))
        l_inv = IIf(RsStock!tm_inventory = Null, "", Format(RsStock!tm_inventory, gs_formatQty))
        With grid
            .AddItem ""
            .TextMatrix(.Rows - 1, bteColProdCod) = Trim(RsStock!Item_Code)
            .TextMatrix(.Rows - 1, bteColDesc) = l_item_name
            .TextMatrix(.Rows - 1, bteColAddress) = IIf(IsNull(Trim(RsStock!Address)), "", Trim(RsStock!Address))
            .TextMatrix(.Rows - 1, bteColPreMonth) = Format(RsStock!tm_premonth, gs_formatQty)
            .TextMatrix(.Rows - 1, bteColReceipt) = Format(RsStock!tm_receipt, gs_formatQty)
            .TextMatrix(.Rows - 1, bteColSupply) = Format(RsStock!tm_supply, gs_formatQty)
            .TextMatrix(.Rows - 1, bteColLossReject) = Format(RsStock!tm_lossreject, gs_formatQty)
            .TextMatrix(.Rows - 1, bteColEnd) = Format(RsStock!tm_current, gs_formatQty)
            .TextMatrix(.Rows - 1, bteColUnit) = Trim(RsStock!unitdesc)
            .TextMatrix(.Rows - 1, bteColInventory) = l_inv
            If .TextMatrix(.Rows - 1, bteColInventory) <> "" Then
                Vdif = CDbl(IIf(.TextMatrix(.Rows - 1, bteColInventory) = "", 0, .TextMatrix(.Rows - 1, bteColInventory))) - CDbl(IIf(.TextMatrix(.Rows - 1, bteColEnd) = "", 0, .TextMatrix(.Rows - 1, bteColEnd)))
                .TextMatrix(.Rows - 1, bteColDiff) = Format(Vdif, gs_formatQty)
            Else
                .TextMatrix(.Rows - 1, bteColDiff) = ""
            End If
            .TextMatrix(.Rows - 1, bteColInv2) = l_inv
            .TextMatrix(.Rows - 1, bteColReason) = IIf(IsNull(RsStock!tm_reason), "", Trim(RsStock!tm_reason))
            .TextMatrix(.Rows - 1, bteColReason2) = IIf(IsNull(RsStock!tm_reason), "", Trim(RsStock!tm_reason))
        End With
        RsStock.MoveNext
    Wend
    
Case 2:

    While RsStock.EOF = False
        l_inv = IIf(RsStock!nm_inventory = Null, "", Format(RsStock!nm_inventory, gs_formatQty))
        l_item_name = uf_GetItemDescription(Trim(RsStock!Item_Code))
        With grid
            .AddItem ""
            .TextMatrix(.Rows - 1, bteColProdCod) = Trim(RsStock!Item_Code)
            .TextMatrix(.Rows - 1, bteColDesc) = l_item_name
            .TextMatrix(.Rows - 1, bteColAddress) = IIf(IsNull(RsStock!Address), "", Trim(RsStock!Address))
            .TextMatrix(.Rows - 1, bteColPreMonth) = Format(RsStock!nm_premonth, gs_formatQty)
            .TextMatrix(.Rows - 1, bteColReceipt) = Format(RsStock!nm_receipt, gs_formatQty)
            .TextMatrix(.Rows - 1, bteColSupply) = Format(RsStock!nm_supply, gs_formatQty)
            .TextMatrix(.Rows - 1, bteColLossReject) = Format(RsStock!nm_lossreject, gs_formatQty)
            .TextMatrix(.Rows - 1, bteColEnd) = Format(RsStock!nm_current, gs_formatQty)
            .TextMatrix(.Rows - 1, bteColUnit) = Trim(RsStock!unitdesc)
            .TextMatrix(.Rows - 1, bteColInventory) = l_inv
            If .TextMatrix(.Rows - 1, bteColInventory) <> "" Then
                Vdif = CDbl(IIf(.TextMatrix(.Rows - 1, bteColInventory) = "", 0, .TextMatrix(.Rows - 1, bteColInventory))) - CDbl(.TextMatrix(.Rows - 1, bteColEnd))
                .TextMatrix(.Rows - 1, bteColDiff) = Format(Vdif, gs_formatQty)
            Else
                .TextMatrix(.Rows - 1, bteColDiff) = ""
            End If
            .TextMatrix(.Rows - 1, bteColInv2) = l_inv
            .TextMatrix(.Rows - 1, bteColReason) = IIf(IsNull(RsStock!nm_reason), "", Trim(RsStock!nm_reason))
            .TextMatrix(.Rows - 1, bteColReason2) = IIf(IsNull(RsStock!nm_reason), "", Trim(RsStock!nm_reason))
        End With
        RsStock.MoveNext
    Wend
               
End Select

    With grid
        For i = 1 To .Rows - 1
            .Cell(flexcpBackColor, i, bteColInventory) = vbWhite
            .Cell(flexcpBackColor, i, bteColReason) = vbWhite
        Next
    End With

'If txtitemcode = "" Then Exit Sub
'
'Dim pilrow As Integer
'For i = 1 To Grid.Rows - 1
'    If Grid.TextMatrix(i, bteColProdCod) = txtitemcode Then
'        pilrow = i: Exit For
'    End If
'Next i
'
'Grid.Col = bteColInventory
'Grid.Row = pilrow
'Grid.SetFocus

sql = "select"
End Sub

Private Sub grid_Click()
With grid
If .Row <> 0 Then
    If .Col = bteColInventory Or .Col = bteColReason Then .FocusRect = flexFocusInset Else .FocusRect = flexFocusNone
Else
End If
End With
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
 If KeyAscii = 13 And grid.Row + 1 <> grid.Rows Then grid.Row = grid.Row + 1
 'If Chr(KeyAscii) < 0 Or Chr(KeyAscii) > 9 And KeyAscii <> bteColInventory And KeyAscii <> Asc(".") Then
 '    KeyAscii = 0
 'End If
 'If KeyAscii = Asc(".") And InStr(1, Trim(Grid.TextMatrix(Grid.Row, bteColInventory)), ".") > 0 Then _
 '    KeyAscii = 0
 If grid.Col = bteColInventory Then
  If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 46 And KeyAscii <> 8 Then KeyAscii = 0
 ElseIf grid.Col = bteColReason Then
  'Free Text
 End If
End Sub

Private Sub Grid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
 If KeyAscii = 13 And grid.Row + 1 <> grid.Rows Then grid.Row = grid.Row + 1
 'If Chr(KeyAscii) < 0 Or Chr(KeyAscii) > 9 And KeyAscii <> 8 And KeyAscii <> Asc(".") Then
 '    KeyAscii = 0
 'End If
 'If KeyAscii = Asc(".") And InStr(1, Trim(Grid.TextMatrix(Grid.Row, bteColInventory)), ".") > 0 Then _
 '    KeyAscii = 0
 If Col = bteColInventory Then
  If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 46 And KeyAscii <> 8 Then KeyAscii = 0
 ElseIf Col = bteColReason Then
  'Free Text
 End If
End Sub

Sub SetPosisi()
Dim cr As Integer
Dim PosRow As Integer

PosRow = 1
For cr = 1 To grid.Rows - 1
    If Trim(grid.TextMatrix(cr, bteColProdCod)) = Trim(txtItemCode) Then
        PosRow = cr
    End If
Next

grid.Row = PosRow
grid.Col = bteColInventory
grid.SetFocus
'Grid.FocusRect = flexFocusInset
SendKeys "{right}"
SendKeys "{left}"

End Sub
