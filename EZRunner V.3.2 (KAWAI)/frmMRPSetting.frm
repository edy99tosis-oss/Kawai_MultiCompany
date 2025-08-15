VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMRPSetting 
   BackColor       =   &H00FDDFE3&
   Caption         =   "MRP Setting"
   ClientHeight    =   4635
   ClientLeft      =   360
   ClientTop       =   3720
   ClientWidth     =   8355
   Icon            =   "frmMRPSetting.frx":0000
   LinkMode        =   1  'Source
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   8355
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   660
      Left            =   420
      TabIndex        =   9
      Top             =   915
      Width           =   7515
      Begin MSComCtl2.DTPicker dtpYear 
         Height          =   315
         Left            =   735
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   225
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy"
         Format          =   334888963
         UpDown          =   -1  'True
         CurrentDate     =   39191
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
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
         Left            =   225
         TabIndex        =   10
         Top             =   285
         Width           =   390
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
      Left            =   420
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3765
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   420
      TabIndex        =   6
      Top             =   3105
      Width           =   7515
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
         TabIndex        =   7
         Top             =   195
         Width           =   7290
      End
   End
   Begin VB.CommandButton cmdSubmit 
      BackColor       =   &H0000FFFF&
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
      Left            =   6795
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3765
      Width           =   1140
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   6090
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   285
      Width           =   1845
      _extentx        =   3254
      _extenty        =   714
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   1395
      Left            =   420
      TabIndex        =   0
      Top             =   1680
      Width           =   7515
      _cx             =   13256
      _cy             =   2461
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   10932991
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483624
      BackColorAlternate=   -2147483643
      GridColor       =   12582912
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
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
      ScrollBars      =   0
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
   Begin VB.Label lblKet 
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
      Left            =   855
      TabIndex        =   8
      Top             =   2835
      Width           =   60
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MRP Setting"
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
      Left            =   420
      TabIndex        =   5
      Top             =   345
      Width           =   7515
   End
End
Attribute VB_Name = "frmMRPSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bteColProcess As Byte
Dim bteColJan As Byte
Dim bteColFeb As Byte
Dim bteColMar As Byte
Dim bteColApr As Byte
Dim bteColMay As Byte
Dim bteColJun As Byte
Dim bteColJul As Byte
Dim bteColAug As Byte
Dim bteColSep As Byte
Dim bteColOct As Byte
Dim bteColNov As Byte
Dim bteColDec As Byte

Private Sub SetGrid()
    
    bteColProcess = 0
    bteColJan = 1
    bteColFeb = 2
    bteColMar = 3
    bteColApr = 4
    bteColMay = 5
    bteColJun = 6
    bteColJul = 7
    bteColAug = 8
    bteColSep = 9
    bteColOct = 10
    bteColNov = 11
    bteColDec = 12
    
    With grid
        
        .clear
        .ColS = 13
        .FixedCols = 1
        .Rows = 4
        .FixedRows = 1
        
        .FormatString = "|^Jan|^Feb|^Mar|^Apr|^May|^Jun|^Jul|^Aug|^Sep|Oct|^Nov|^Dec"
        
        .TextMatrix(1, bteColProcess) = "Forecast"
        .TextMatrix(2, bteColProcess) = "Order"
        .TextMatrix(3, bteColProcess) = "Daily Production"
        
        .ColWidth(bteColProcess) = 1515
        .ColWidth(bteColJan) = 500
        .ColWidth(bteColFeb) = 500
        .ColWidth(bteColMar) = 500
        .ColWidth(bteColApr) = 500
        .ColWidth(bteColMay) = 500
        .ColWidth(bteColJun) = 500
        .ColWidth(bteColJul) = 500
        .ColWidth(bteColAug) = 500
        .ColWidth(bteColSep) = 500
        .ColWidth(bteColOct) = 500
        .ColWidth(bteColNov) = 500
        .ColWidth(bteColDec) = 500
        
        .Cell(flexcpChecked, 1, 1, .Rows - 1, .ColS - 1) = flexUnchecked
        
    End With
    
End Sub

Private Sub SetGridData()
    
    Dim adoRs As New ADODB.Recordset
    Dim intCount As Integer
    
    On Error GoTo errHandler
    
    grid.Cell(flexcpChecked, 1, 1, grid.Rows - 1, grid.ColS - 1) = flexUnchecked
    
    sql = "Select MRP_Year, MRP_Month, Calc_F, Calc_O, Calc_D " & _
        "From MRP_Setting " & _
        "Where MRP_Year = " & dtpYear.Year
    adoRs.Open sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
    While Not adoRs.EOF
        If adoRs.Fields("Calc_F") = 1 Then grid.Cell(flexcpChecked, 1, adoRs.Fields("MRP_Month")) = flexChecked
        If adoRs.Fields("Calc_O") = 1 Then grid.Cell(flexcpChecked, 2, adoRs.Fields("MRP_Month")) = flexChecked
        If adoRs.Fields("Calc_D") = 1 Then grid.Cell(flexcpChecked, 3, adoRs.Fields("MRP_Month")) = flexChecked
        adoRs.MoveNext
    Wend
    adoRs.Close
    
ErrExit:
    Set adoRs = Nothing
    Exit Sub
errHandler:
    LblErrMsg.Caption = err.Description
    err.clear
    Resume ErrExit
    
End Sub

Private Sub SaveData()
    
    Dim adoRs As New ADODB.Recordset
    Dim intCount As Integer
    
    On Error GoTo errHandler
    Me.MousePointer = vbHourglass
    
    With grid
        For intCount = 1 To .ColS - 1
            
            sql = "Select * From MRP_Setting Where MRP_Year = " & dtpYear.Year & " And MRP_Month = " & intCount
            adoRs.Open sql, Db, adOpenDynamic, adLockOptimistic, adCmdText
            If adoRs.EOF Then
                adoRs.AddNew
                adoRs.Fields("MRP_Year") = dtpYear.Year
                adoRs.Fields("MRP_Month") = Format(intCount, "00")
            End If
            If .Cell(flexcpChecked, 1, intCount) = flexChecked Then adoRs.Fields("Calc_F") = 1 Else adoRs.Fields("Calc_F") = 0
            If .Cell(flexcpChecked, 2, intCount) = flexChecked Then adoRs.Fields("Calc_O") = 1 Else adoRs.Fields("Calc_O") = 0
            If .Cell(flexcpChecked, 3, intCount) = flexChecked Then adoRs.Fields("Calc_D") = 1 Else adoRs.Fields("Calc_D") = 0
            adoRs.Fields("Last_Update") = Now
            adoRs.Fields("Last_User") = userLogin
            adoRs.update
            adoRs.Close
        Next
        
        sql = "Delete From MRP_Setting Where MRP_Year = " & dtpYear.Year & " And Calc_F = 0 And Calc_O = 0 And Calc_D = 0"
        Db.Execute sql
        
    End With
    
    LblErrMsg.Caption = DisplayMsg("1000")
    
ErrExit:
    Set adoRs = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
errHandler:
    LblErrMsg.Caption = err.Description
    err.clear
    Resume ErrExit
    
End Sub

Private Sub CmdSubMenu_Click()
    
    DoEvents
    frmMainMenu.Show
    DoEvents
    Unload Me

End Sub

Private Sub CmdSubmit_Click()
    
    SaveData
    
End Sub

Private Sub dtpYear_Change()
    
    LblErrMsg.Caption = ""
    SetGridData
    
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
    
    dtpYear.Value = Date
    SetGrid
    SetGridData
    
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    LblErrMsg.Caption = ""
    
End Sub

