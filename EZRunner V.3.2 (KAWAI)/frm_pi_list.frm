VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_pi_list 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Physical Inventory List"
   ClientHeight    =   3870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8595
   Icon            =   "frm_pi_list.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3870
   ScaleWidth      =   8595
   StartUpPosition =   2  'CenterScreen
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   420
      Left            =   6427
      TabIndex        =   10
      Top             =   420
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   741
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
      Left            =   307
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3075
      Width           =   1125
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   307
      TabIndex        =   4
      Top             =   2340
      Width           =   7980
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
         Left            =   105
         TabIndex        =   5
         Top             =   195
         Width           =   7755
      End
   End
   Begin VB.CommandButton Cmd_Save 
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
      Index           =   0
      Left            =   7252
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3090
      Width           =   1035
   End
   Begin MSComCtl2.DTPicker DMonth 
      Height          =   315
      Left            =   2025
      TabIndex        =   1
      Top             =   1740
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
      Format          =   150863875
      UpDown          =   -1  'True
      CurrentDate     =   37798
   End
   Begin MSForms.ComboBox CboLocationCD 
      Height          =   315
      Left            =   2025
      TabIndex        =   0
      Top             =   1320
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
      Left            =   375
      TabIndex        =   9
      Top             =   1350
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
      Left            =   375
      TabIndex        =   8
      Top             =   1800
      Width           =   1335
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
      Left            =   4545
      TabIndex        =   7
      Top             =   1350
      Width           =   3015
   End
   Begin VB.Line Line1 
      X1              =   4545
      X2              =   7560
      Y1              =   1575
      Y2              =   1575
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Physical Inventory List"
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
      Left            =   307
      TabIndex        =   6
      Top             =   450
      Width           =   6630
   End
End
Attribute VB_Name = "frm_pi_list"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dateUp As Date

Private Sub CboLocationCD_Change()
If CboLocationCD.MatchFound Then
   LblLocationName = CboLocationCD.List(CboLocationCD.ListIndex, 1)
   LblErrMsg = ""
Else
   LblLocationName = ""
   LblErrMsg = DisplayMsg(4018) ' "Invalid warehouse code !"
End If
End Sub

Private Sub CboLocationCD_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
Dim j As Integer
If KeyCode = 13 Then
  j = 0
For i = 0 To CboLocationCD.ListCount - 1
    If UCase(Trim(CboLocationCD)) = UCase(Trim(CboLocationCD.List(i, 0))) Then
        CboLocationCD = Trim(CboLocationCD.List(i, 0))
        LblLocationName = Trim(CboLocationCD.List(i, 1))
        j = 1: Exit For
    End If
Next

If j = 0 Then
    LblErrMsg = DisplayMsg(4018) '"Invalid warehouse code !"
    Exit Sub
Else
    LblErrMsg = ""
End If
End If
End Sub

Private Sub CboLocationCD_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Cmd_Save_Click(Index As Integer)
Dim j As Integer
Select Case Index
       Case 8:
                
                frmMainMenu.Show
                
                Unload Me

        Case 0:
                
        
              Dim application As New CRAXDDRT.application
              Dim report As New CRAXDDRT.report
              Dim rsRpt As New ADODB.Recordset
              Dim Rpt As New FrmRpt3
              Dim sqlControl As String, RsInvControl As New ADODB.Recordset

                 
                 j = 0
                For i = 0 To CboLocationCD.ListCount - 1
                    If UCase(Trim(CboLocationCD)) = UCase(Trim(CboLocationCD.List(i, 0))) Then
                        CboLocationCD = Trim(CboLocationCD.List(i, 0))
                        LblLocationName = Trim(CboLocationCD.List(i, 1))
                        j = 1: Exit For
                    End If
                Next
                
                If j = 0 Then
                    LblErrMsg = DisplayMsg(4018) '"Invalid warehouse code !"
                    Exit Sub
                End If
        
LblErrMsg = up_ValidateDateRange(DMonth.Value, False)
If Trim(LblErrMsg) <> "" Then Exit Sub

             
              LblErrMsg = ""
              Me.MousePointer = vbHourglass
        
              sql = "select rtrim(im.makeritem_code) makeritem_code, descriptions=case isnull(im.sheetcoil_cls,0) when 0 then " & _
                    vbLf & "rtrim(im.item_name) else rtrim(im.item_name) + ' (' + rtrim(sh.description) + ', T' + cast(im.thickness as varchar(15)) + ' x W' + cast(im.width as varchar(15)) + ' x L' + cast(im.length as varchar(15)) + ')'  end, " & _
                    vbLf & "sm.*, rtrim(address) address, isnull(wh_name,trade_name) wh_name from stock_master sm left join warehouse_master wm on " & _
                    vbLf & "sm.warehouse_code = wm.wh_code " & _
                    vbLf & "left join trade_master tm on tm.Trade_code=sm.warehouse_code " & _
                    vbLf & "left join item_master im on sm.item_code=im.item_code " & _
                    vbLf & "left join sheetcoil_cls sh on im.sheetcoil_cls = sh.sheetcoil_cls " & _
                    vbLf & "where warehouse_code='" & Trim(CboLocationCD) & "' "

              sql = sql & vbLf & "order by  warehouse_code,sm.item_Code "
            
              If rsRpt.State <> adStateClosed Then rsRpt.Close
              'rsRpt.Open Sql, Db, adOpenDynamic, adLockOptimistic
              Set rsRpt = Db.Execute(sql)
            
              If rsRpt.EOF Then LblErrMsg.Caption = DisplayMsg(4006): Me.MousePointer = vbDefault: Exit Sub
            
              sqlprint = sql
              reportcode = "pilist"
              printorient = 2
              Set report = application.OpenReport(App.path & "\Reports\rpt_pi_list.rpt")
              report.Database.Tables(1).SetDataSource rsRpt
                       
''#####################################################################
''# Qty Digit and decimal
report.FormulaFields(2).Text = "" & gi_decimalDigitQty & ""
report.FormulaFields(3).Text = "" & gi_decimalDigitQty & ""
''#####################################################################
              
               Select Case up_GetDateRange(DMonth.Value)

                Case 0:
                        report.Sections(4).Suppress = False
                        report.Sections(5).Suppress = True
                        report.Sections(6).Suppress = True

                 Case 1:

                        report.Sections(4).Suppress = True
                        report.Sections(5).Suppress = False
                        report.Sections(6).Suppress = True
                                                                       
                 Case 2:

                        report.Sections(4).Suppress = True
                        report.Sections(5).Suppress = True
                        report.Sections(6).Suppress = False
                                                                
              End Select
              Dim dates As String
            
             dates = Format(DMonth.Value, "MMM yyyy")
             dtMPList = DMonth.Value
             datePiList = Format(DMonth.Value, "MMM yyyy")
             report.FormulaFields(1).Text = "'" & dates & "'"
             report.ReportTitle = "Physical Inventory List ( Summary )"
            
              Rpt.CRViewer1.ReportSource = report
              Rpt.CRViewer1.ViewReport
              Rpt.CRViewer1.Zoom 1
            
              Rpt.WindowState = 2
              Rpt.Show 1
            
              Me.MousePointer = vbDefault
                
End Select
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    LblErrMsg.Caption = ErrMsg
End If
End Sub



Private Sub DMonth_Change()
If Format(DMonth.Value, "MM") < Format(dateUp, "MM") And Val(Format(DMonth.Value, "MM")) = 1 And Val(Format(dateUp, "MM")) = 12 Then _
            DMonth.Year = DMonth.Year + 1: GoTo pass
    If Format(DMonth.Value, "MM") > Format(dateUp, "MM") And Val(Format(DMonth.Value, "MM")) = 12 And Val(Format(dateUp, "MM")) = 1 Then _
            DMonth.Year = DMonth.Year - 1
pass:
    dateUp = Format(DMonth.Value, "dd MMM yyyy")
    
'DMonth.Value = Format(DMonth.Value, "MMM yyyy")
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
LblLocationName = ""
LblErrMsg = ""
DMonth.Value = Format(Date, "MMM yyyy")
dateUp = DMonth.Value
CtrlMenu1.FormName = Me.Name
Me.Caption = "Physical Inventory List"
Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"

Call StockLocation
DMonth = Format(Now, "mmmm yyyy")
End Sub


Private Sub StockLocation()
Dim sql As String, ls_sql As String, RsStock As New ADODB.Recordset
Dim i As Integer

If RsStock.State <> adStateClosed Then RsStock.Close

ls_sql = " select * from (select wh_code, wh_name  from warehouse_master where stockcontrol_cls='01' union  " & _
      " select trade_code wh_code, trade_name wh_name from trade_master where trade_code in(select manufacture_code from manufacture_line))tbWarehouse order by wh_code "
      
RsStock.Open ls_sql, Db, adOpenDynamic, adLockOptimistic

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


