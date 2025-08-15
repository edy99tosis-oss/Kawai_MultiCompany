VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmUploadAccountBalance 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Upload Account Balance"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmUploadAccountBalance.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cd1 
      Left            =   8280
      Top             =   10140
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdTemp 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Template"
      Height          =   375
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   10200
      Width           =   1000
   End
   Begin VB.PictureBox PicProg 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   1005
      Left            =   600
      ScaleHeight     =   975
      ScaleWidth      =   13905
      TabIndex        =   17
      Top             =   8220
      Width           =   13935
      Begin MSComctlLib.ProgressBar Prg1 
         Height          =   795
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   13680
         _ExtentX        =   24130
         _ExtentY        =   1402
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VB.Frame FrameFunction 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   600
      TabIndex        =   0
      Top             =   1080
      Width           =   13935
      Begin MSComDlg.CommonDialog cdg 
         Left            =   6120
         Top             =   600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdBrowse 
         BackColor       =   &H0080FFFF&
         Caption         =   "&Browse"
         Height          =   375
         Left            =   11040
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txtBrowse 
         Height          =   330
         Left            =   2040
         TabIndex        =   3
         Top             =   1320
         Width           =   8775
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Source File"
         Height          =   225
         Index           =   1
         Left            =   240
         TabIndex        =   16
         Top             =   1320
         Width           =   1155
      End
      Begin MSForms.ComboBox cboPeriod 
         Height          =   345
         Left            =   2040
         TabIndex        =   2
         Top             =   840
         Width           =   3180
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "5609;609"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Courier New"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Budget Period"
         Height          =   225
         Index           =   0
         Left            =   240
         TabIndex        =   15
         Top             =   840
         Width           =   1365
      End
      Begin MSForms.ComboBox cboYear 
         Height          =   345
         Left            =   2040
         TabIndex        =   1
         Top             =   360
         Width           =   1005
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "1764;609"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Courier New"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Budget Year"
         Height          =   225
         Index           =   2
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   1155
      End
   End
   Begin BSREBudget.CtrlMenu CtrlMenu1 
      Height          =   1005
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   49995
      _ExtentX        =   88186
      _ExtentY        =   1773
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   4620
      Left            =   600
      TabIndex        =   5
      Top             =   3195
      Width           =   13965
      _cx             =   24633
      _cy             =   8149
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483640
      BackColorFixed  =   12640511
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
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
   Begin VB.CommandButton cmdSubMenu 
      BackColor       =   &H0080FFFF&
      Caption         =   "Sub &Menu"
      Height          =   375
      Left            =   630
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   10200
      Width           =   1140
   End
   Begin VB.CommandButton cmdImport 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Import"
      Height          =   375
      Left            =   13590
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   10200
      Width           =   1000
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   630
      TabIndex        =   10
      Top             =   9300
      Width           =   13965
      Begin VB.Label LblErrMsg 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   345
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   13710
      End
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
      Left            =   600
      TabIndex        =   14
      Top             =   7920
      Width           =   2100
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Budget System"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   12720
      TabIndex        =   12
      Top             =   120
      Width           =   1785
   End
End
Attribute VB_Name = "FrmUploadAccountBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'##################################
'Created By     : Yakko
'Created Date   : 1 April 2009
'Last Update By :
'Last Update    :
'##################################

Option Explicit


Dim li_HakU As Integer, ls_Answer As String

Const col_No As Long = 0
Const col_Account As Long = 1
Const col_Amount As Long = 2
Const col_Des As Long = 3
Const col_Count As Long = 4

Dim ls_PathExcel As String


Private Sub cboPeriod_Change()
    LblErrMsg = ""
End Sub

Private Sub CboYear_Change()
    LblErrMsg = ""
End Sub

Private Sub CmdBrowse_Click()
Dim strFileName As String
    LblErrMsg = ""
    On Error GoTo errHandler
    cdg.CancelError = True
    cdg.filter = "Excel Files (*.xls)|*.xls"
    cdg.ShowOpen
    
    strFileName = cdg.FileName
    txtBrowse = strFileName
    If txtBrowse <> "" Then cmdImport.SetFocus
    Exit Sub
errHandler:
    
End Sub

Private Sub CmdImport_Click()
Dim Exl As Excel.Application
Dim IsOpen As Boolean
Dim iRow As Long
Dim ls_Sql As String
Dim ls_BudgetYear As String
Dim ls_BudgetPeriod As String
Dim ls_aCCount As String

Dim TotalUpdate As Long
Dim li_Upd As Long
Dim li_Error As Long
Dim ls_Error As String
Dim MaxRow As Long
    
    If uf_ValidateImport = False Then Exit Sub

    On Error GoTo errHandler
    LblErrMsg = ""
    ls_BudgetYear = CboYear
    ls_BudgetPeriod = Mid(cboPeriod, 1, 2)
    Me.MousePointer = 11
    cmdImport.Enabled = False
    Set Exl = New Excel.Application
    Exl.Workbooks.Open txtBrowse
    IsOpen = True
    
    '======= validasi header =================
    
    Debug.Print Mid(Exl.Range("A1").Text, 1, 12)
    Debug.Print Mid(Exl.Range("B1").Text, 1, 6)
    
    If Not (Mid(Exl.Range("A1").Text, 1, 12) = "Account Code" _
        And Mid(Exl.Range("B1").Text, 1, 6) = "Amount") Then
        
        Call DisplayMessageLabel("Invalid file !", LblErrMsg, ErrorMessage)
        txtBrowse.SetFocus
        GoTo errExit
    End If
    
    
    iRow = 1
    MaxRow = 0
    Prg1 = 0
    LblTotalRec = ""
    Do
        iRow = iRow + 1
        If Exl.Range("A" & iRow).Text = "" Then
            Exit Do
        Else
            MaxRow = MaxRow + 1
        End If
    Loop
    If MaxRow > 0 Then
        Prg1.Max = MaxRow
    Else
        Prg1.Max = 1
    End If
    
    
    iRow = 1
    Call up_GridHeader
    li_Error = 0
    On Error Resume Next
    Do
        iRow = iRow + 1
        If Exl.Range("A" & iRow).Text = "" Then
            Exit Do
        Else
            ls_aCCount = Exl.Range("A" & iRow).Value
            
            ls_Sql = "Update plbs_Account_Balance set " & vbCrLf & _
                "Amount01 = " & Val(Exl.Range("B" & iRow).Value) & vbCrLf & _
                ",Last_Update = GetDate() " & vbCrLf & _
                ",User_Update = '" & UserLogin & "' " & vbCrLf
                
            ls_Sql = ls_Sql & _
                "where Budget_Year = '" & ls_BudgetYear & "' " & vbCrLf & _
                "and Budget_Period = '" & ls_BudgetPeriod & "' " & vbCrLf & _
                "and Account_Code = '" & ls_aCCount & "' " & vbCrLf
                    
            Db.Execute ls_Sql, li_Upd
            If Err.Number <> 0 Then
                ls_Error = Err.Description
                GoTo AdaError
            End If
            
            If li_Upd = 0 Then
                ls_Sql = "Insert into plbs_account_Balance (" & vbCrLf & _
                    "Budget_year, budget_period, account_Code, amount01, " & vbCrLf & _
                    "last_update, user_update ) values (" & vbCrLf & _
                    "'" & ls_BudgetYear & "', '" & ls_BudgetPeriod & "', '" & ls_aCCount & "', "
                    
                ls_Sql = ls_Sql & _
                    Val(Exl.Range("B" & iRow).Value) & ", " & vbCrLf & _
                    "GetDate(), '" & UserLogin & "') "
                Db.Execute ls_Sql
            End If
            
            If Err.Number <> 0 Then
            
                If Err.Number = -2147217873 Then
                    ls_Error = Replace$(Right$(Err.Description, Len(Err.Description) - InStr(1, Err.Description, "column") - Len("column") - 1), "'.", "") & " is not found"
                Else
                    ls_Error = Err.Description
                End If
                
AdaError:
                li_Error = li_Error + 1
                Grid.Rows = Grid.Rows + 1
                Grid.Cell(flexcpAlignment, Grid.Rows - 1, 0, Grid.Rows - 1, Grid.Cols - 1) = flexAlignLeftCenter
                Grid.TextMatrix(Grid.Rows - 1, col_No) = iRow
                Grid.TextMatrix(Grid.Rows - 1, col_Account) = Exl.Range("A" & iRow).Value
                Grid.TextMatrix(Grid.Rows - 1, col_Amount) = Exl.Range("B" & iRow).Value
                Grid.TextMatrix(Grid.Rows - 1, col_Des) = ls_Error
                Grid.Cell(flexcpAlignment, Grid.Rows - 1, col_Amount) = flexAlignRightCenter
            Else
                TotalUpdate = TotalUpdate + 1
            End If
            Err.Clear
            ls_Error = ""
            Prg1 = Prg1 + 1
            
        End If
    Loop
    
    
    LblTotalRec = TotalUpdate & " record(s) updated"
    If li_Error = 0 Then
        If TotalUpdate = 0 Then
            Call DisplayMessageLabel("No data to import !", LblErrMsg, ErrorMessage)
        Else
            Call DisplayMessageLabel("Upload Account Balance completed successfully !", LblErrMsg, InformationMessage)
        End If
    Else
        Call DisplayMessageLabel("Upload Account Balance completed with error(s) !", LblErrMsg, InformationMessage)
    End If
    GoTo errExit
    Exit Sub
errHandler:
    Call DisplayMessageLabel(Err.Description, LblErrMsg, ErrorMessage)
    
errExit:
    If IsOpen Then Exl.Quit
    Set Exl = Nothing
    cmdImport.Enabled = True
    Me.MousePointer = 0
End Sub

Private Sub cmdTemp_Click()
    up_ExcelSave
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub CmdSubMenu_Click()
    ls_Answer = MsgBox("Are you sure want to close this menu ?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation")
    If ls_Answer = vbYes Then
        If cmdSubMenu.Caption = "Sub &Menu" Then
            DoEvents
            frmMainMenu.Show
            DoEvents
            Unload Me
        End If
    End If
End Sub

Private Sub Form_Load()
On Error GoTo errHandler
    Me.MousePointer = vbHourglass
    Me.Caption = "[" & frmcode(Me.Name) & "] " & Me.Caption
    li_HakU = AllowUpdate(Me.Name)
    up_SettingColor
    up_ClearScreen
    up_GridHeader
errExit:
    Me.MousePointer = vbDefault
    Exit Sub
errHandler:
    If InStr(1, Err.Description, "Connection") = 0 Then
        Call DisplayMessageLabel(Err.Description, LblErrMsg)
        Err.Clear
        Resume errExit
    Else
        Call DisplayMessageWindow("Connection to Server Error", Err.Description, "Error", OkOnly, DefaultButton1, Error)
        frmLogin.Show
        Unload Me
    End If
End Sub

Private Sub up_SettingColor()
    Me.BackColor = ColorSetting
    Dim frame As Control
    For Each frame In Me.Controls
        If TypeOf frame Is frame Then frame.BackColor = Me.BackColor
    Next
    Dim cmd As Control
    For Each cmd In Me.Controls
        If TypeOf cmd Is CommandButton And cmd.Name <> "CmdBrowse" Then cmd.BackColor = LightYellow
    Next
End Sub

Private Sub up_ClearScreen()
    up_FillCombo
    LblTotalRec = ""
    CboYear.ListIndex = -1
    cboPeriod.ListIndex = -1
End Sub

Private Sub up_FillCombo()
    G_ComboLoad CboYear, "Budget_Year", "exp_Budget_Master", 50, "50 pt"
    G_ComboLoad cboPeriod, "Budget_Period + ' - ' + Period_Name As Budget_Period", "BudgetPeriod_Master where Budget_Period in ('01','02', '04', '05')", 0, "100 pt;0 pt", "Period_ShortName"
End Sub

Private Sub up_GridHeader()
    With Grid
        .Cols = col_Count
        .Rows = 1
        .FixedRows = 1
        .FixedCols = 0
        .TextMatrix(0, col_No) = "Row No"
        .TextMatrix(0, col_Account) = "Account Code"
        .TextMatrix(0, col_Amount) = "Amount"
        .TextMatrix(0, col_Des) = "Error Description"
        .ColWidth(col_No) = 800
        .ColWidth(col_Account) = 2000
        .ColWidth(col_Amount) = 1800
        .ColWidth(col_Des) = 5000
        .Cell(flexcpAlignment, 0, 0, 0, col_Count - 1) = flexAlignCenterCenter
        
        .ColAlignment(col_Amount) = flexAlignRightCenter
        .ColFormat(col_Amount) = "#,0.##"
        
    End With
End Sub

Private Function uf_ValidateImport() As Boolean
    If CboYear.MatchFound = False Then
        Call DisplayMessageLabel("Please select Budget Year !", LblErrMsg, ErrorMessage)
        CboYear.SetFocus
        Exit Function
    ElseIf cboPeriod.MatchFound = False Then
        Call DisplayMessageLabel("Please select Budget Period !", LblErrMsg, ErrorMessage)
        cboPeriod.SetFocus
        Exit Function
    ElseIf Trim$(txtBrowse) = "" Then
        Call DisplayMessageLabel("Please select Excel File !", LblErrMsg, ErrorMessage)
        txtBrowse.SetFocus
        Exit Function
    End If
    
    uf_ValidateImport = True
End Function

Private Sub txtBrowse_Change()
    LblErrMsg = ""
End Sub

Private Sub up_ExcelSave()
    Dim li_idx As Double
    If G_CekExcelApp = False Then Call DisplayMessageLabel("Excel Application is not found !", LblErrMsg, ErrorMessage): Exit Sub

    
    LblErrMsg.Caption = ""
    
    CD1.CancelError = True
    CD1.filter = "Excel Files (*.xls)|*.xls"
    CD1.FileName = "Template Upload Account Balance"
    On Error GoTo errCancel
    CD1.ShowSave
    
    
    On Error GoTo errHandling
    If Len(CD1.FileName) = 0 Then Exit Sub
    If Dir(CD1.FileName) <> "" Then
        If MsgBox("Overwrite existing file?", vbExclamation + vbYesNo, "Overwrite") = vbNo Then Exit Sub
    End If
    ls_PathExcel = Mid(CD1.FileName, 1, Len(CD1.FileName) - Len(CD1.FileTitle))
    MousePointer = MousePointerConstants.vbHourglass
    Call DisplayMessageLabel("Please wait while generating excel file...", LblErrMsg, InformationMessage)
   
    up_ExcelOpen
    
    MousePointer = 0
    Exit Sub
errHandling:
    If Err.Number <> 0 Then
        MousePointer = 0
        If Err.Description = "Permission denied" Then
            Call DisplayMessageLabel("The file is still opened !", LblErrMsg, ErrorMessage)
        Else
            Call DisplayMessageLabel(Err.Description, LblErrMsg, ErrorMessage)
        End If
    End If
errCancel:
    
End Sub


Private Sub up_ExcelOpen()
    Dim Exl As Excel.Application
    Dim li_row As Double, ls_Temp As String
    Dim iRow As Long, iCol As Long
    Dim iMonth As Long
    
    On Error GoTo errHandler

    Me.MousePointer = 11
    Set Exl = New Excel.Application
    Exl.DisplayAlerts = False
    Exl.Workbooks.Add

    Exl.Range("A1:B1").Interior.ColorIndex = ClrOrange
    Exl.Range("A1:B1").Borders.Weight = xlThin
    Exl.Range("A1").Value = "Account Code"
    Exl.Range("B1").Value = "Amount"
    
    Exl.Range("A1").ColumnWidth = 15
    Exl.Range("B1").ColumnWidth = 20
    Exl.Range("A1").RowHeight = 25
    Exl.Range("A1:B1").WrapText = True
    Exl.Range("A1:B1").HorizontalAlignment = xlCenter
    Exl.Range("A1:B1").VerticalAlignment = xlCenter
        
    Exl.ActiveWorkbook.SaveAs FileName:=CD1.FileName, _
    FileFormat:=xlNormal, _
    Password:="", _
    WriteResPassword:="", _
    ReadOnlyRecommended:=False, _
    CreateBackup:=False
    
    Exl.DisplayAlerts = True
    Exl.Visible = True
    
    Me.MousePointer = 0
    LblErrMsg.Caption = ""
    
    Exit Sub

errHandler:
    Me.MousePointer = 0
    Call DisplayMessageLabel(Err.Description, LblErrMsg, ErrorMessage)
    On Error Resume Next
    Exl.Quit
    Set Exl = Nothing
End Sub


