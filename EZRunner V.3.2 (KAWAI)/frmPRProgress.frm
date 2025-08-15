VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPRProgress 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Report Purchase Request Progress"
   ClientHeight    =   4350
   ClientLeft      =   1935
   ClientTop       =   3585
   ClientWidth     =   9495
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPRProgress.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4350
   ScaleWidth      =   9495
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog Dlg 
      Left            =   4500
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Excel"
      Height          =   375
      Left            =   6510
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3750
      Width           =   1185
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   5685
      Left            =   300
      TabIndex        =   18
      Top             =   4320
      Visible         =   0   'False
      Width           =   14295
      _cx             =   25215
      _cy             =   10028
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
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
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
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   7380
      TabIndex        =   16
      Top             =   120
      Width           =   1845
      _extentx        =   3254
      _extenty        =   714
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   2115
      Left            =   270
      TabIndex        =   10
      Top             =   900
      Width           =   8925
      Begin MSComCtl2.DTPicker dtAwal 
         Height          =   315
         Left            =   1920
         TabIndex        =   0
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
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
         Format          =   294649859
         CurrentDate     =   37810
      End
      Begin MSComCtl2.DTPicker dtAkhir 
         Height          =   315
         Left            =   4080
         TabIndex        =   1
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
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
         Format          =   294715395
         CurrentDate     =   37810
      End
      Begin MSForms.ComboBox CboType 
         Height          =   315
         Left            =   1920
         TabIndex        =   2
         Top             =   675
         Width           =   1875
         VariousPropertyBits=   746604571
         MaxLength       =   10
         DisplayStyle    =   7
         Size            =   "3307;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Request Type"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   17
         Top             =   720
         Width           =   1170
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "to"
         Height          =   255
         Left            =   3630
         TabIndex        =   15
         Top             =   270
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Request Date From"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   14
         Top             =   315
         Width           =   1650
      End
      Begin MSForms.ComboBox CboItemCode 
         Height          =   315
         Left            =   1920
         TabIndex        =   4
         Top             =   1575
         Width           =   2910
         VariousPropertyBits=   746604571
         MaxLength       =   25
         DisplayStyle    =   3
         Size            =   "5133;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         Value           =   "AAAAAAAAAAAAAAB"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Request No"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   13
         Top             =   1170
         Width           =   975
      End
      Begin MSForms.ComboBox CboReqNo 
         Height          =   315
         Left            =   1920
         TabIndex        =   3
         Top             =   1110
         Width           =   1875
         VariousPropertyBits=   746604571
         MaxLength       =   10
         DisplayStyle    =   3
         Size            =   "3307;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         Value           =   "AAAAAAAAAB"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Code"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   12
         Top             =   1635
         Width           =   1155
      End
      Begin VB.Label LblItem 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   4890
         TabIndex        =   11
         Top             =   1635
         Width           =   3705
      End
      Begin VB.Line Line8 
         Index           =   0
         X1              =   4905
         X2              =   8595
         Y1              =   1875
         Y2              =   1875
      End
   End
   Begin VB.CommandButton cmdSubMenu 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Sub &Menu"
      Height          =   375
      Left            =   270
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3735
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   600
      Left            =   270
      TabIndex        =   8
      Top             =   3045
      Width           =   8925
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
         Left            =   90
         TabIndex        =   9
         Top             =   240
         Width           =   8775
      End
   End
   Begin VB.CommandButton cmdReport 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Preview"
      Height          =   375
      Left            =   8010
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3735
      Width           =   1185
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Report Purchase Request Progress"
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
      Left            =   2310
      TabIndex        =   7
      Top             =   210
      Width           =   3825
   End
End
Attribute VB_Name = "frmPRProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ClsProc As New ClsProc
Dim nilKosong As Boolean, i As Integer

Sub Kosong()
    dtAwal = Format(Year(Now) & "-" & Format(Month(Now), "#0") & "-01", "dd MMM yyyy")
    dtAkhir = Format(Now, "dd MMM yyyy")
    
    CboType.ListIndex = 0
    
    Call isiCboRequest
    Call isiCboItem
    
    CboReqNo.ListIndex = 0
    cboItemCode.ListIndex = 0
End Sub

Sub isiCboRequest() 'Filter Request No
Dim rscbo As New ADODB.Recordset 'Data Customer

With CboReqNo
    .clear
    .columnCount = 1
    .TextColumn = 1
    .Text = ""
    
    sql = "select PORequest_No From PORequest_Master " & _
        "Where PORequest_Date >= '" & Format(dtAwal, "yyyy-MM-dd") & _
            "' And PORequest_Date <= '" & Format(dtAkhir, "yyyy-MM-dd") & "'"
            
    If CboType.ListIndex = 1 Then
        sql = sql + " And Others_Cls = '0' And SheetCoil_Cls = '0' "
    ElseIf CboType.ListIndex = 2 Then
        sql = sql + " And Others_Cls = '0' And SheetCoil_Cls = '1' "
    ElseIf CboType.ListIndex = 3 Then
        sql = sql + " And Others_Cls = '1' "
    End If
    
    Set rscbo = Db.Execute(sql)
    
    .AddItem ""
    .List(0, 0) = strAll
        
    i = 1
    Do While Not (rscbo.EOF)
        .AddItem ""
        .List(i, 0) = Trim(rscbo(0))
        i = i + 1
        rscbo.MoveNext
    Loop
    
    .ListWidth = 250
    .ColumnWidths = "250 pt"
    
    Set rscbo = Nothing
End With
End Sub

Sub isiCboItem()
Dim rscbo As New ADODB.Recordset

With cboItemCode
    .clear
    .columnCount = 2
    .TextColumn = 1
    .Text = ""
    
    sql = "select PD.Item_Code, I.Item_Name From PORequest_Master PM, PORequest_Detail PD, Item_Master I " & _
        "Where PM.PORequest_No = PD.PORequest_No " & _
            "And PD.Item_Code = I.Item_Code "
                
    If CboReqNo.ListIndex = 0 Then
        sql = sql & "And PM.PORequest_Date >= '" & Format(dtAwal, "yyyy-MM-dd") & _
            "' And PM.PORequest_Date <= '" & Format(dtAkhir, "yyyy-MM-dd") & "' "
        If CboType.ListIndex = 1 Then
            sql = sql + " And PM.Others_Cls = '0' And PM.SheetCoil_Cls = '0' "
        ElseIf CboType.ListIndex = 2 Then
            sql = sql + " And PM.Others_Cls = '0' And PM.SheetCoil_Cls = '1' "
        ElseIf CboType.ListIndex = 3 Then
            sql = sql + " And PM.Others_Cls = '1' "
        End If
    Else
        sql = sql & "And PM.PORequest_No = '" & CboReqNo.Text & "' "
    End If
    sql = sql & "Order By PD.Item_Code"
    Set rscbo = Db.Execute(sql)
    
    .AddItem ""
    .List(0, 0) = strAll
    .List(0, 1) = strAll
    
    i = 1
    Do While Not (rscbo.EOF)
        .AddItem ""
        .List(i, 0) = Trim(rscbo("Item_Code"))
        .List(i, 1) = Trim(rscbo("Item_Name"))
        
        i = i + 1
        rscbo.MoveNext
    Loop
    
    .ListWidth = 400
    .ColumnWidths = "100 pt;300 pt"
    Set rscbo = Nothing
End With
End Sub

Private Sub CboItemCode_Change()
cboItemCode.Text = cboItemCode.Text
lblitem.Caption = ""
If cboItemCode.MatchFound = True Then
    lblitem.Caption = cboItemCode.List(cboItemCode.ListIndex, 1)
End If
End Sub

Private Sub CboReqNo_Change()
Call isiCboItem
End Sub

Private Sub cboType_Change()
Call isiCboRequest
End Sub

Private Sub CboType_Click()
Call isiCboRequest
End Sub

Private Sub Command1_Click()
Dim rsRpt As New ADODB.Recordset
Dim li_Idx As Integer

    
    
LblErrMsg = ""

If CDate(dtAwal) > CDate(dtAkhir) Then
    LblErrMsg.Caption = "Start Date must be lower than " & dtAkhir.Value & " !!!"
    Exit Sub
ElseIf CDate(dtAkhir) < CDate(dtAwal) Then
    LblErrMsg.Caption = "End Date must be higher than " & dtAwal.Value & " !!!"
    Exit Sub
End If

    CboReqNo.Text = CboReqNo.Text
    cboItemCode.Text = cboItemCode.Text
    If Trim(CboReqNo.Text) = "" Then
        LblErrMsg = DisplayMsg(1067) 'Please Input Req No
        CboReqNo.SetFocus
    ElseIf CboReqNo.MatchFound = False Then
        LblErrMsg = DisplayMsg(4081) 'Req No Not Found
        CboReqNo.SetFocus
    ElseIf Trim(cboItemCode.Text) = "" Then
        LblErrMsg = DisplayMsg(1009) 'Please Input Product Code
        cboItemCode.SetFocus
    ElseIf cboItemCode.MatchFound = False Then
        LblErrMsg = DisplayMsg(4003) 'Product Code Not Found
        cboItemCode.SetFocus
    Else
        Me.MousePointer = vbHourglass
        
        sql = SqlProgress
    
        Set rsRpt = Db.Execute(sql)
        If rsRpt.EOF Then
            LblErrMsg.Caption = DisplayMsg(4006)
        Else
            With grid
                .ColS = 15
                .Rows = 1
                .TextMatrix(0, 0) = "Request No"
                .TextMatrix(0, 1) = "Request Date"
                .TextMatrix(0, 2) = "Req Dlvy Date"
                .TextMatrix(0, 3) = "Product Code"
                .TextMatrix(0, 4) = "Item Name"
                .TextMatrix(0, 5) = "Class"
                .TextMatrix(0, 6) = "Dept."
                .TextMatrix(0, 7) = "Qty"
                .TextMatrix(0, 8) = "Unit"
                .TextMatrix(0, 9) = "Req Qty"
                .TextMatrix(0, 10) = "PO No"
                .TextMatrix(0, 11) = "PO Date"
                .TextMatrix(0, 12) = "PO Dlvy Date"
                .TextMatrix(0, 13) = "PO Qty"
                .TextMatrix(0, 14) = "Rem Qty"
            li_Idx = 1
            Do While Not rsRpt.EOF
                .Rows = .Rows + 1
                .TextMatrix(li_Idx, 0) = Trim(rsRpt!PORequest_No)
                .TextMatrix(li_Idx, 1) = Trim(rsRpt!PORequest_Date)
                .TextMatrix(li_Idx, 2) = Trim(rsRpt!ReqDelivery_Date)
                .TextMatrix(li_Idx, 3) = Trim(rsRpt!Item_Code)
                .TextMatrix(li_Idx, 4) = Trim(rsRpt!item_name)
                .TextMatrix(li_Idx, 5) = Trim(rsRpt!Class)
                .TextMatrix(li_Idx, 6) = Trim(rsRpt!Department_Cls)
                .TextMatrix(li_Idx, 7) = Trim(rsRpt!Number_Box)
                .TextMatrix(li_Idx, 8) = Trim(rsRpt!unit)
                .TextMatrix(li_Idx, 9) = Trim(rsRpt!Qty)
                If Trim(rsRpt!po_no) = "" Then
                    .TextMatrix(li_Idx, 10) = ""
                    .TextMatrix(li_Idx, 11) = ""
                    .TextMatrix(li_Idx, 12) = ""
                Else
                    .TextMatrix(li_Idx, 10) = Trim(rsRpt!po_no)
                    .TextMatrix(li_Idx, 11) = Trim(rsRpt!po_date)
                    .TextMatrix(li_Idx, 12) = Trim(rsRpt!delivery_Date)
                End If
                .TextMatrix(li_Idx, 13) = Trim(rsRpt!pQty)
                .TextMatrix(li_Idx, 14) = rsRpt!Qty - rsRpt!SUmQtyPO
                li_Idx = li_Idx + 1
                rsRpt.MoveNext
            Loop
            End With
            SaveExcel
        End If
        Set rsRpt = Nothing
    End If
    Me.MousePointer = vbDefault
End Sub


Private Sub SaveExcel()
Dim PathExcel As String
    On Error GoTo errHandling
    LblErrMsg = ""
    Dlg.filter = "Excel Files (*.xls)|*.xls"
    Dlg.filename = "PR Progress Report"
    Dlg.ShowSave
    If Dlg.FileTitle = "" Then Exit Sub
    If Len(Dlg.filename) = 0 Then Exit Sub
    If Dir(Dlg.filename) <> "" Then
        If MsgBox("Overwrite existing file?", vbExclamation + vbYesNo, "Overwrite") = vbNo Then Exit Sub
    End If
    PathExcel = Mid(Dlg.filename, 1, Len(Dlg.filename) - Len(Dlg.FileTitle))
    
    Me.MousePointer = vbHourglass
    
    With grid
        .FixedRows = 0
        .FixedCols = 0
        .SaveGrid PathExcel & "$tmp$" & Dlg.FileTitle, flexFileExcel
        .FixedRows = 1
    End With
    If Dir(Dlg.filename) <> "" Then Kill Dlg.filename
    Sleep 500
    Call OpenExcel(PathExcel)
    Exit Sub
errHandling:
    If err.number <> 0 Then
        Me.MousePointer = vbDefault
        If err.Description <> "Cancel was selected." Then
            If err.Description = "Permission denied" Then
                LblErrMsg = "File is still opened !"
            Else
                LblErrMsg = err.Description
            End If
        End If
        If Dir(PathExcel & "$tmp$" & Dlg.FileTitle) <> "" Then Kill PathExcel & "$tmp$" & Dlg.FileTitle
    End If
End Sub

Private Sub OpenExcel(PathExcel As String)
    
    On Error GoTo SheetClose
    Dim ExlFile As New Excel.application
    Dim rsCompany As New Recordset
    Dim AddRow As Byte


    ExlFile.Workbooks.Open PathExcel & "$tmp$" & Dlg.FileTitle

    'Copy whole cells to a New Excel
    ExlFile.Range(ExlFile.Cells(1, 1), ExlFile.Cells(grid.Rows, grid.ColS)).Copy
    ExlFile.Workbooks.Add
    ExlFile.ActiveSheet.Paste
    ExlFile.application.CutCopyMode = False
    With ExlFile.Selection.Font
        .Name = "Arial"
        .Size = 8
    End With
            
    With ExlFile

        .ActiveWorkbook.SaveAs filename:= _
            Dlg.filename, _
            FileFormat:=xlNormal, Password:="", WriteResPassword:="", _
            ReadOnlyRecommended:=False, CreateBackup:=False
    
        .Workbooks.Close
        Kill PathExcel & "$tmp$" & Dlg.FileTitle
        .Workbooks.Open Dlg.filename
    End With
    
    'Alignment
    With ExlFile.Range(ExlFile.Cells(1, 1), ExlFile.Cells(1, 10))
        .VerticalAlignment = xlVAlignCenter
        .HorizontalAlignment = xlHAlignCenter
    End With
    'Border
'    ExlFile.Range("A1", "J1").Borders(xlEdgeTop).LineStyle = xlContinuous
'    ExlFile.Range("A1", "J1").Borders(xlEdgeBottom).LineStyle = xlContinuous
'    ExlFile.Range("A" & (Grid.Rows - 1), "J" & (Grid.Rows - 1)).Borders(xlEdgeBottom).LineStyle = xlContinuous
        

        'Parameter
        With ExlFile
            .Rows(1).Insert xlDown
            .Rows(1).Insert xlDown
            .Range("A1") = "Product Code"
            .Range("B1") = ": " & cboItemCode.Text
            .Rows(1).Insert xlDown
            .Range("A1") = "Request No"
            .Range("B1") = ": " & CboReqNo.Text
            .Rows(1).Insert xlDown
            .Range("A1") = "Request Type"
            .Range("B1") = ": " & CboType.Text
            .Rows(1).Insert xlDown
            .Range("A1") = " Period "
            .Range("B1") = ": " & Format(dtAwal, "dd MMM YYYY") & " To " & Format(dtAkhir, "dd MMM YYYY")
        End With
        
    'Company Header
        ExlFile.Rows(1).Insert xlDown
        sql = "select rtrim(company_name) company_name, rtrim(address1) Address1, rtrim(Address2) Address2, rtrim(Province) Province, rtrim(city) City, Rtrim(Postal_Code) Postal_Code, Rtrim(phone1) Phone1, Rtrim(phone2) Phone2, rtrim(fax) Fax  From company_profile "
        If rsCompany.State <> adStateClosed Then rsCompany.Close
        rsCompany.Open sql, Db, adOpenDynamic, adLockOptimistic
        If rsCompany.EOF Then Me.MousePointer = vbDefault: Exit Sub
    
        ExlFile.Range("A1", "O1").Merge
        ExlFile.Range("A1") = "Phone: " & rsCompany!phone1 & " " & rsCompany!phone2 & " Fax: " & rsCompany!fax
        ExlFile.Range("A1").HorizontalAlignment = xlHAlignCenter
        ExlFile.Rows(1).Insert xlDown
        ExlFile.Range("A1", "O1").Merge
        ExlFile.Range("A1") = rsCompany!address1 & " " & rsCompany!address2 & " " & rsCompany!City & " " & rsCompany!Province & " " & rsCompany!postal_code
        ExlFile.Range("A1").HorizontalAlignment = xlHAlignCenter
        ExlFile.Rows(1).Insert xlDown
        ExlFile.Range("A1", "O1").Merge
        ExlFile.Range("A1") = rsCompany!company_name
        ExlFile.Range("A1").Font.Size = 10
        ExlFile.Range("A1").Font.Bold = True
        ExlFile.Range("A1").HorizontalAlignment = xlHAlignCenter
        ExlFile.Rows(1).Insert xlDown
    
    ExlFile.Range("A:J").Columns.AutoFit
    ExlFile.Range("A1").Select
    Set rsCompany = Nothing
      
    With ExlFile

    
        
        .Range("A10:O" & grid.Rows + 9).Select
       .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
       .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
       With .Selection.Borders(xlEdgeLeft)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
       End With
       With .Selection.Borders(xlEdgeTop)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
       End With
       With .Selection.Borders(xlEdgeBottom)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
       End With
       With .Selection.Borders(xlEdgeRight)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
       End With
       With .Selection.Borders(xlInsideVertical)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
       End With
       .Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
       .Range("A10:O10").Select
       .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
       .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
       With .Selection.Borders(xlEdgeLeft)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
       End With
       With .Selection.Borders(xlEdgeTop)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
       End With
       With .Selection.Borders(xlEdgeBottom)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
       End With
       With .Selection.Borders(xlEdgeRight)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
       End With
       With .Selection.Borders(xlInsideVertical)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
       End With
    
    
    
    .Range("H" & grid.Rows + 10) = "TOTAL"
    .Range("J" & grid.Rows + 10).Select
    .ActiveCell.Formula = "=SUM(J11:J" & grid.Rows + 9 & ")"
    .Range("J" & grid.Rows + 10).Select
    .Selection.Copy
    .Range("N" & grid.Rows + 10 & ":O" & grid.Rows + 10).Select
    .ActiveSheet.Paste
    .application.CutCopyMode = False
    
'    .Range("J10").Select
'    .Range(Selection, Selection.End(xlDown)).Select
'    .Selection.NumberFormat = "General"
'    .Selection.NumberFormat = "#,##0.00_);(#,##0.00)"
'
'    .Range("N10").Select
'    .Range(Selection, Selection.End(xlDown)).Select
'    .Selection.NumberFormat = "General"
'    .Selection.NumberFormat = "#,##0.00_);(#,##0.00)"
'
'    .Range("O10").Select
'    .Range(Selection, Selection.End(xlDown)).Select
'    .Selection.NumberFormat = "General"
'    .Selection.NumberFormat = "#,##0.00_);(#,##0.00)"
    
    
    .Range("A:O").Columns.AutoFit
    .Range("A5:O" & grid.Rows + 10).Select
    
    With ExlFile.Selection.Font
        .Name = "Arial"
        .Size = 8
    End With
    
    End With
    
    
    ExlFile.Visible = True
    ExlFile.ActiveWorkbook.save
    
    
    Exit Sub

SheetClose:
    If err.number <> 0 Then
        Me.MousePointer = vbDefault
        ExlFile.Workbooks.Close
        ExlFile.application.Quit
        If err.Description <> "Cancel was selected." Then LblErrMsg = err.Description
        If err.Description = "Permission denied" Then LblErrMsg.Caption = "The file is still opened!"
        If Dir(PathExcel & "$tmp$" & Dlg.FileTitle) <> "" Then Kill PathExcel & "$tmp$" & Dlg.FileTitle
    End If
    
End Sub

Private Sub Form_Load()
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"

    dtAkhir.Value = Now
    dtAwal.Value = Now
    
    CboType.AddItem strAll
    CboType.AddItem "Part/Material"
    CboType.AddItem "Sheet/Coil"
    CboType.AddItem "Other Item"
    
    Call Kosong
End Sub

Public Sub cmdReport_Click()
Dim application As New CRAXDDRT.application
Dim report As New CRAXDDRT.report
Dim rsRpt As New ADODB.Recordset
Dim Rpt As New FrmRpt3

    
    
LblErrMsg = ""

If CDate(dtAwal) > CDate(dtAkhir) Then
    LblErrMsg.Caption = "Start Date must be lower than " & dtAkhir.Value & " !!!"
    Exit Sub
ElseIf CDate(dtAkhir) < CDate(dtAwal) Then
    LblErrMsg.Caption = "End Date must be higher than " & dtAwal.Value & " !!!"
    Exit Sub
End If

    CboReqNo.Text = CboReqNo.Text
    cboItemCode.Text = cboItemCode.Text
    If Trim(CboReqNo.Text) = "" Then
        LblErrMsg = DisplayMsg(1067) 'Please Input Req No
        CboReqNo.SetFocus
    ElseIf CboReqNo.MatchFound = False Then
        LblErrMsg = DisplayMsg(4081) 'Req No Not Found
        CboReqNo.SetFocus
    ElseIf Trim(cboItemCode.Text) = "" Then
        LblErrMsg = DisplayMsg(1009) 'Please Input Product Code
        cboItemCode.SetFocus
    ElseIf cboItemCode.MatchFound = False Then
        LblErrMsg = DisplayMsg(4003) 'Product Code Not Found
        cboItemCode.SetFocus
    Else
        Me.MousePointer = vbHourglass
        
        sql = SqlProgress
    
        Set rsRpt = Db.Execute(sql)
        If rsRpt.EOF Then
            LblErrMsg.Caption = DisplayMsg(4006)
        Else
            'Dim kt1 As String
            sqlprint = sql
            tglAwalRptPrint = "'" & Format(dtAwal, "dd MMM yyyy") & " To " & Format(dtAkhir, "dd MMM yyyy") & "'"
            printorient = 2
            reportcode = "PORequestPrint"
            Set report = application.OpenReport(App.path & "\Reports\RptPOReqReport.rpt")
            report.Database.Tables(1).SetDataSource rsRpt
            report.FormulaFields(1).Text = tglAwalRptPrint
            Kt1 = "'" & CboType.Text & "'"
            report.FormulaFields(2).Text = "'" & CboType.Text & "'"
            Kt2 = "'" & CboReqNo.Text & "'"
            report.FormulaFields(3).Text = "'" & CboReqNo.Text & "'"
            Kt3 = "'" & cboItemCode.Text & " - " & lblitem.Caption & "'"
            report.FormulaFields(4).Text = "'" & cboItemCode.Text & " - " & lblitem.Caption & "'"
            report.PaperOrientation = crLandscape
            Rpt.CRViewer1.ReportSource = report
            Rpt.CRViewer1.ViewReport
            Rpt.CRViewer1.Zoom 1
            
            Rpt.WindowState = 2
            Rpt.Show 1
        End If
        Set rsRpt = Nothing
    End If
    Me.MousePointer = vbDefault
End Sub

Private Sub dtAwal_Change()
If CDate(dtAwal) > CDate(dtAkhir) Then
    LblErrMsg.Caption = "Start Date must be lower than " & dtAkhir.Value & " !!!"
    Exit Sub
ElseIf CDate(dtAkhir) < CDate(dtAwal) Then
    LblErrMsg.Caption = "End Date must be higher than " & dtAwal.Value & " !!!"
    Exit Sub
End If

    Call isiCboRequest
    If Format(dtAwal, "yyyy-MM-dd") > Format(dtAkhir, "yyyy-MM-dd") Then LblErrMsg = DisplayMsg(4068) & " " & Format(dtAkhir, "dd MMM yyyy") Else LblErrMsg = ""
End Sub

Private Sub dtAkhir_Change()
If CDate(dtAwal) > CDate(dtAkhir) Then
    LblErrMsg.Caption = "Start Date must be lower than " & dtAkhir.Value & " !!!"
    Exit Sub
ElseIf CDate(dtAkhir) < CDate(dtAwal) Then
    LblErrMsg.Caption = "End Date must be higher than " & dtAwal.Value & " !!!"
    Exit Sub
End If

    Call isiCboRequest
    If Format(dtAwal, "yyyy-MM-dd") > Format(dtAkhir, "yyyy-MM-dd") Then LblErrMsg = DisplayMsg(4066) & " " & Format(dtAwal, "dd MMM yyyy") Else LblErrMsg = ""
End Sub

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


Function SqlProgress() As String
sql = " SELECT     dbo.PORequest_Master.PORequest_No, dbo.PORequest_Detail.PoReq_SeqNo, dbo.PORequest_Master.PORequest_Date,  " & _
    "                       dbo.PORequest_Detail.ReqDelivery_Date, dbo.PORequest_Detail.Item_Code, ISNULL(dbo.PORequest_Detail.Item_Name,  " & _
    "                       dbo.Item_Master.Item_Name) AS Item_Name, Class = isnull(PORequest_Detail.Class, 'RM'), ISNULL(dbo.Item_Master.Number_Box, 0) AS Number_Box, dbo.PORequest_Detail.Qty,  " & _
    "                       (select description from unit_cls uc where uc.unit_cls= PORequest_Detail.unit_cls ) Unit, " & _
    "                       isnull(dbo.PurchaseOrder_Detail.PO_No,'') PO_NO, isnull(dbo.PurchaseOrder_Master.PO_Date,'') PO_Date, isnull(dbo.PurchaseOrder_Detail.Delivery_Date,'') Delivery_Date,  " & _
    "                       isnull(dbo.PurchaseOrder_Detail.Qty,0) AS PQty, isnull " & _
    "                           ((SELECT     SUM(qty) " & _
    "                               FROM         purchaseorder_detail " & _
    "                               GROUP BY porequest_no, poreq_seqno " & _
    "                               HAVING      purchaseorder_detail.PORequest_No = dbo.PORequest_Master.PORequest_No AND poreq_seqno = dbo.PORequest_detail.PoReq_seqno), 0)  "

    sql = sql + "                       AS SumQtyPO , dbo.PORequest_Master.Department_Cls " & _
                " FROM         dbo.PORequest_Master INNER JOIN " & _
                "                       dbo.PORequest_Detail ON dbo.PORequest_Master.PORequest_No = dbo.PORequest_Detail.PORequest_No LEFT OUTER JOIN " & _
                "                       dbo.PurchaseOrder_Master INNER JOIN " & _
                "                       dbo.PurchaseOrder_Detail ON dbo.PurchaseOrder_Master.PO_No = dbo.PurchaseOrder_Detail.PO_No ON  " & _
                "                       dbo.PORequest_Detail.PORequest_No = dbo.PurchaseOrder_Detail.PORequest_No AND  " & _
                "                       dbo.PORequest_Detail.PoReq_SeqNo = dbo.PurchaseOrder_Detail.POReq_SeqNo LEFT OUTER JOIN " & _
                "                       dbo.Item_Master ON dbo.PORequest_Detail.Item_Code = dbo.Item_Master.Item_Code "
                
     sql = sql & "WHERE dbo.poRequest_master.PORequest_Date >= '" & Format(dtAwal, "yyyy-MM-dd") & _
                    "' And dbo.poRequest_master.PORequest_Date <= '" & Format(dtAkhir, "yyyy-MM-dd") & "' "
        
    'type
        
    If CboType.ListIndex = 1 Then
        sql = sql & "AND dbo.poRequest_master.others_Cls = '0' AND dbo.poRequest_master.sheetcoil_Cls = '0' "
    ElseIf CboType.ListIndex = 2 Then
        sql = sql & "AND dbo.poRequest_master.others_Cls = '0' AND dbo.poRequest_master.sheetcoil_Cls = '1' "
    ElseIf CboType.ListIndex = 3 Then
        sql = sql & "AND dbo.poRequest_master.others_Cls = '1'"
    End If
    
    'reqno
    If CboReqNo.ListIndex > 0 Then sql = sql & "AND dbo.poRequest_master.PORequest_No = '" & CboReqNo.Text & "' "
    
     'Item
    If cboItemCode.ListIndex > 0 Then sql = sql & "And dbo.PORequest_Detail.Item_Code = '" & cboItemCode & "' "
     
     sql = sql + " ORDER BY dbo.PORequest_Master.PORequest_No, dbo.PORequest_Detail.PoReq_SeqNo "
SqlProgress = sql

End Function


