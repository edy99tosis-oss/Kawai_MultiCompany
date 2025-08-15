VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmBOMMasterlist 
   BackColor       =   &H00FDDFE3&
   Caption         =   "BOM Masterlist"
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8910
   Icon            =   "FrmBOMMasterlist.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   8910
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Print"
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
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3000
      Width           =   1140
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3000
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   240
      TabIndex        =   5
      Top             =   2040
      Width           =   8475
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
         TabIndex        =   6
         Top             =   195
         Width           =   8130
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   975
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   8475
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "..."
         Height          =   300
         Left            =   1920
         TabIndex        =   9
         Top             =   360
         Width           =   300
      End
      Begin VB.Label lblModel 
         BackColor       =   &H00FDDFE3&
         Caption         =   "lbl_Model"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2280
         TabIndex        =   7
         Top             =   405
         Width           =   2685
      End
      Begin VB.Line Line7 
         X1              =   2280
         X2              =   4920
         Y1              =   645
         Y2              =   645
      End
      Begin MSForms.ComboBox cboModel 
         Height          =   330
         Left            =   960
         TabIndex        =   0
         Top             =   360
         Width           =   885
         VariousPropertyBits=   746604571
         MaxLength       =   15
         DisplayStyle    =   3
         Size            =   "1561;582"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblSerialNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
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
         Left            =   240
         TabIndex        =   4
         Top             =   400
         Width           =   495
      End
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   6840
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   120
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   714
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BOM Masterlist"
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
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   8565
   End
End
Attribute VB_Name = "FrmBOMMasterlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS As New Recordset
Dim TempModelCls As String

Private Sub cboModel_Click()
lblModel.Caption = cboModel.List(cboModel.ListIndex, 1)
LblErrMsg.Caption = ""
End Sub

Private Sub cmdBrowse_Click()
LblErrMsg.Caption = ""

    Me.MousePointer = vbHourglass
    
    frm_BrowseItem_Model.Show 1
    TempModelCls = frm_BrowseItem_Model.txtTmpModel
    
    If Len(TempModelCls) > 2 Then
        cboModel.Text = "Multiselect"
        lblModel.Caption = "Multiselect"
    Else
        cboModel.Text = TempModelCls
        
        If cboModel.Text = "" Then
            lblModel.Caption = ""
        End If
        
    End If
    
     Me.MousePointer = vbDefault
End Sub

Private Sub cmdPrint_Click()

LblErrMsg = ""

MousePointer = vbHourglass

If lblModel.Caption = "Multiselect" Then
    sql = "SELECT RTRIM(im.Item_Code)Item_Code FROM dbo.Item_Master IM WHERE IM.Model_Cls IN (SELECT * FROM dbo.fnParseArray('" & TempModelCls & "', ','))"
Else
    sql = "SELECT RTRIM(im.Item_Code)Item_Code FROM dbo.Item_Master IM WHERE IM.Model_Cls = '" & Trim(cboModel.Text) & "'"
End If
 
    If RS.State = 1 Then RS.Close
    RS.Open sql, Db, adOpenKeyset, adLockOptimistic
    
    If RS.EOF Then
        LblErrMsg.Caption = DisplayMsg(8012)
        Exit Sub
    End If
        
    ExportXLS
    
MousePointer = vbDefault

End Sub



Private Sub Form_Load()
    If gb_Simulation = True Then Call up_InitSimulation(Me)
    
    up_Clear
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
    
End Sub

Private Sub up_Clear()
           
    '=====================Setting Combo Model=================
    cboModel.clear
    cboModel.columnCount = 2
    cboModel.TextColumn = 1
    i = 1
    Call up_FillComboModel
    
    cboModel.ColumnWidths = "20 pt; 70 pt"
    cboModel.ListWidth = 90
    cboModel.ListIndex = -1
    
    lblModel.Caption = ""
    '==============================================================
    
End Sub

Private Sub CmdSubMenu_Click()
    DoEvents
    frmMainMenu.Show
    DoEvents
    Unload Me
End Sub

Function ExportXLS()
    Dim oExcel          As Object
    Dim oExcelWrkBk     As Object
    Dim oExcelWrSht     As Object
    Dim bExcelOpened    As Boolean
    Dim iCols           As Integer
    Const xlCenter = -4108
    Dim cmd As ADODB.Command
    Dim sql As String
   
    'Start Excel
    On Error Resume Next
    Set oExcel = GetObject(, "Excel.Application")    'Bind to existing instance of Excel

    If err.number <> 0 Then    'Could not get instance of Excel, so create a new one
        err.clear
        Set oExcel = CreateObject("excel.application")
        bExcelOpened = False
    Else    'Excel was already running
        bExcelOpened = True
    End If
    
'    MousePointer = vbHourglass

    oExcel.ScreenUpdating = False
    oExcel.Visible = False   'Keep Excel hidden until we are done with our manipulation
    Set oExcelWrkBk = oExcel.Workbooks.Add()    'Start a new workbook
    Set oExcelWrSht = oExcelWrkBk.Sheets(1)
    
    'Execute Sotred Procedure

    Dim RS As ADODB.Recordset
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "SP_MasterlistModel_Select"
        
    cmd.Parameters.append cmd.CreateParameter("ModelCls", adVarChar, adParamInput, 11, Trim(cboModel.Text))
    cmd.Parameters.append cmd.CreateParameter("TempModelCls", adVarChar, adParamInput, 20, TempModelCls)
    
    Set RS = cmd.Execute
    
    If RS.EOF = False Then
    
        With RS
            If .RecordCount <> 0 Then
                'Build our Header
                For iCols = 0 To RS.Fields.Count - 1
                    oExcelWrSht.Cells(4, iCols + 1).Value = RS.Fields(iCols).Name
                Next
                
                oExcelWrSht.Range("a1") = "PT Kawai Indonesia Plant 3"
                oExcelWrSht.Range("a1").Columns.Font.Name = "Consolas"
                oExcelWrSht.Range("a1").Columns.Font.Size = "11"
                oExcelWrSht.Range("a1").Columns.Font.Bold = True
                
                oExcelWrSht.Range("a3") = "MASTERLIST MODEL : " & lblModel.Caption
                oExcelWrSht.Range("a3").Columns.Font.Name = "Consolas"
                oExcelWrSht.Range("a3").Columns.Font.Size = "11"
                oExcelWrSht.Range("a3").Columns.Font.Bold = True
                
                With oExcelWrSht.Range(oExcelWrSht.Cells(4, 1), _
                                       oExcelWrSht.Cells(4, RS.Fields.Count))
                    
                    .Font.ColorIndex = 2
                    .Interior.ColorIndex = 1
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .WrapText = True
                End With
                
                oExcelWrSht.Range("a4").Columns.ColumnWidth = 15
                oExcelWrSht.Range("b4").Columns.ColumnWidth = 50
                oExcelWrSht.Range("c4").Columns.ColumnWidth = 12
                oExcelWrSht.Range("d4").Columns.ColumnWidth = 45
                oExcelWrkBk.ActiveSheet.Cells(4, RS.Fields.Count).ColumnWidth = 11
                
                oExcelWrSht.Range(oExcelWrSht.Cells(4, 5), _
                                  oExcelWrSht.Cells(4, (RS.Fields.Count - 1))).ColumnWidth = 13 'Resize our Columns based on the headings
                
               oExcelWrkBk.ActiveSheet.Rows(4).EntireRow.RowHeight = 30
                                  
                'Copy the data from our query into Excel
                oExcelWrSht.Range("A5").CopyFromRecordset RS
                oExcelWrSht.Range("A4").Select  'Return to the top of the page
                            
                oExcelWrSht.Range(oExcelWrSht.Cells(4, (RS.Fields.Count - 1)), _
                                  oExcelWrSht.Cells((RS.RecordCount + 4), (RS.Fields.Count - 1))).HorizontalAlignment = xlCenter
                
'                MousePointer = vbDefault
                
                oExcel.Visible = True
                Set oExcelWrSht = Nothing
                Set oExcelWrkBk = Nothing
                oExcel.ScreenUpdating = True
                
                LblErrMsg.Caption = DisplayMsg(9008)
             
            End If
        End With
    
    Else
        LblErrMsg.Caption = DisplayMsg(8012)
    End If
    
    
End Function

Private Sub up_FillComboModel()
Dim sql As String
Dim RS As New Recordset
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_Model_Sel"
    
    Set RS = cmd.Execute

    With cboModel
        .clear
        .columnCount = 2
        .ColumnWidths = "50pt;180pt"
        .ListWidth = 230
        .ListRows = 15
    
        i = 0
        
        Do While Not RS.EOF
            .AddItem
            .List(i, 0) = Trim(RS("Model_Cls") & "")
            .List(i, 1) = Trim(RS("Description") & "")
            
            RS.MoveNext
            i = i + 1
        Loop
        
        .ListIndex = 0
    End With
End Sub
