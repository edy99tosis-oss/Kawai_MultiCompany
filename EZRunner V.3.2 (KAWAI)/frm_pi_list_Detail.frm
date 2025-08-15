VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_pi_list_Detail 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Physical Inventory List (Detail)"
   ClientHeight    =   4260
   ClientLeft      =   3495
   ClientTop       =   6720
   ClientWidth     =   8595
   Icon            =   "frm_pi_list_Detail.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4260
   ScaleWidth      =   8595
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   315
      Left            =   4459
      TabIndex        =   11
      Top             =   2145
      Width           =   330
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
      Left            =   311
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3435
      Width           =   1125
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   311
      TabIndex        =   4
      Top             =   2700
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
      Left            =   7256
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3450
      Width           =   1035
   End
   Begin MSComCtl2.DTPicker DMonth 
      Height          =   315
      Left            =   2029
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
      Format          =   149684227
      UpDown          =   -1  'True
      CurrentDate     =   37798
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   420
      Left            =   6424
      TabIndex        =   12
      Top             =   375
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   741
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Physical Inventory List (Detail)"
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
      Left            =   304
      TabIndex        =   13
      Top             =   405
      Width           =   7050
   End
   Begin VB.Label Label19 
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
      Index           =   3
      Left            =   409
      TabIndex        =   10
      Top             =   2220
      Width           =   915
   End
   Begin MSForms.ComboBox CboItem 
      Height          =   315
      Left            =   2029
      TabIndex        =   9
      Top             =   2160
      Width           =   2370
      VariousPropertyBits=   612386843
      MaxLength       =   15
      DisplayStyle    =   3
      Size            =   "4180;556"
      ListRows        =   15
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cbolocationcd 
      Height          =   315
      Left            =   2029
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
      AutoSize        =   -1  'True
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
      Height          =   195
      Index           =   0
      Left            =   409
      TabIndex        =   8
      Top             =   1350
      Width           =   1305
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      Height          =   195
      Left            =   409
      TabIndex        =   7
      Top             =   1800
      Width           =   1125
   End
   Begin VB.Label LblLocationName 
      AutoSize        =   -1  'True
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
      Height          =   195
      Left            =   4549
      TabIndex        =   6
      Top             =   1350
      Width           =   1440
   End
   Begin VB.Line Line1 
      X1              =   4549
      X2              =   7564
      Y1              =   1575
      Y2              =   1575
   End
End
Attribute VB_Name = "frm_pi_list_Detail"
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
'On Error Resume Next
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
                
                Dim PilField As String
               Select Case up_GetDateRange(DMonth.Value)

                Case 0:
                        PilField = "LM_Current"
                 Case 1:
                        PilField = "TM_Current"
                 Case 2:
                        PilField = "NM_Current"
              End Select


              LblErrMsg = ""
              Me.MousePointer = vbHourglass

              sql = "select rtrim(im.makeritem_code) makeritem_code, descriptions=case isnull(im.sheetcoil_cls,0) when 0 then " & _
                    vbLf & "rtrim(im.item_name) else rtrim(im.item_name) + ' (' + rtrim(sh.description) + ', T' + cast(im.thickness as varchar(15)) + ' x W' + cast(im.width as varchar(15)) + ' x L' + cast(im.length as varchar(15)) + ')'  end, " & _
                    vbLf & "sm.item_code,sm.warehouse_code," & PilField & ",rtrim(address) address, isnull(wh_name,trade_name) wh_name,PackingStyle = isnull(PackingStyle,''),NumberCase,sm.warehouse_code+sm.item_Code as prod_barcode  " & _
                    vbLf & "from stock_master sm left join warehouse_master wm on " & _
                    vbLf & "sm.warehouse_code = wm.wh_code " & _
                    vbLf & "left join trade_master tm on tm.Trade_code=sm.warehouse_code " & _
                    vbLf & "left join item_master im on sm.item_code=im.item_code " & _
                    vbLf & "left join sheetcoil_cls sh on im.sheetcoil_cls = sh.sheetcoil_cls "
                    
            sql = sql & "Inner Join " & _
                    vbLf & "(Select Item_Code,PackingStyle=description,NumberCase from " & _
                    vbLf & "(Select Item_Code,finishgoodpart_cls,packingstyle_cls,packingstylematerial_cls, " & _
                    vbLf & "case finishgoodpart_cls when '01' then packingstyle_cls else packingstylematerial_cls end as PS_Cls, " & _
                    vbLf & "case finishgoodpart_cls when '01' then number_entering else Number_Box end as NumberCase " & _
                    vbLf & "from item_master) a left join packingStyle_Cls b on a.ps_cls=b.packingstyle_cls) pc " & _
                    vbLf & "on im.item_code=pc.item_code "
                    
                sql = sql & vbLf & "where warehouse_code='" & Trim(CboLocationCD) & "' "
                    
                    If CboItem.ListIndex <> 0 Then sql = sql & vbLf & "And makeritem_code='" & Trim(CboItem) & "' "

              sql = sql & vbLf & "order by  warehouse_code,sm.item_Code "

              If rsRpt.State <> adStateClosed Then rsRpt.Close
              
              'rsRpt.Open Sql, Db, adOpenDynamic, adLockOptimistic
              Set rsRpt = Db.Execute(sql)

              If rsRpt.EOF Then LblErrMsg.Caption = DisplayMsg(4006): Me.MousePointer = vbDefault: Exit Sub
              
              sqlprint = sql
              reportcode = "pilistdet"
              printorient = 1
              Set report = application.OpenReport(App.path & "\Reports\rpt_pi_list_Details.rpt")
              report.Database.Tables(1).SetDataSource rsRpt

'#####################################################################
'# Qty Digit and decimal
report.FormulaFields(2).Text = "" & gi_decimalDigitQty & ""
report.FormulaFields(3).Text = "" & gi_decimalDigitQty & ""
'#####################################################################

              Dim dates As String
             
             report.DiscardSavedData
             
             dates = Format(DMonth.Value, "MMM yyyy")
             dtMPList = DMonth.Value
             datePiList = Format(DMonth.Value, "MMM yyyy")
             report.FormulaFields(1).Text = "'" & dates & "'"
             report.ReportTitle = "Physical Inventory List (Detail)"

              Rpt.CRViewer1.ReportSource = report
              Rpt.CRViewer1.ViewReport
              Rpt.CRViewer1.Zoom 1

              Rpt.WindowState = 2
              Rpt.Show 1

              Me.MousePointer = vbDefault

End Select
End Sub

Private Sub Command1_Click()
 Me.MousePointer = vbHourglass
 
   frm_BrowseItem.getItemCode = CboItem.Text
   frm_BrowseItem.Show 1
   CboItem.Text = frm_BrowseItem.getItemCode
 
 Me.MousePointer = vbDefault

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

DMonth.Value = Format(DMonth.Value, "MMM yyyy")
End Sub

Private Sub Form_Load()

If gb_Simulation = True Then Call up_InitSimulation(Me)
LblLocationName = ""
LblErrMsg = ""
DMonth.Value = Format(Date, "MMM yyyy")
dateUp = DMonth.Value
CtrlMenu1.FormName = Me.Name
Me.Caption = "Physical Inventory List Detail"

Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"

Call StockLocation
DMonth = Format(Now, "mmmm yyyy")
End Sub


Private Sub StockLocation()
Dim sql As String, ls_sql As String, RsStock As New ADODB.Recordset
Dim lrs_ss As New ADODB.Recordset
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


If lrs_ss.State <> adStateClosed Then lrs_ss.Close
ls_sql = " select * from item_master"


lrs_ss.Open ls_sql, Db, adOpenKeyset, adLockOptimistic

CboItem.columnCount = 2
CboItem.clear

CboItem.AddItem ""
CboItem.List(0, 0) = "ALL"
CboItem.List(0, 1) = ""

i = 1
    While lrs_ss.EOF = False
            CboItem.AddItem ""
            CboItem.List(i, 0) = Trim(lrs_ss("item_code"))
            CboItem.List(i, 1) = Trim(lrs_ss("item_name"))

            i = i + 1
            lrs_ss.MoveNext
    Wend

CboItem.ColumnWidths = "100 pt; 0 pt"
CboItem.ListWidth = 150
CboItem.ListRows = 15
    
If lrs_ss.State <> adStateClosed Then lrs_ss.Close

CboItem.ListIndex = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then Cancel = 1
End Sub


