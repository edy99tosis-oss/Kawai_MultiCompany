VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEffMaterialConsumptionRpt 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Material Consumption Report"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9330
   Icon            =   "frmEffMaterialConsumptionRpt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   9330
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBrowser 
      Caption         =   "..."
      Height          =   300
      Left            =   4575
      TabIndex        =   21
      Top             =   2347
      Width           =   300
   End
   Begin VB.CommandButton CmdPreview 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Preview"
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
      Left            =   7890
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4380
      Width           =   1140
   End
   Begin VB.CommandButton cmdExcel 
      BackColor       =   &H0080FFFF&
      Caption         =   "To E&xcel"
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
      Left            =   6660
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4380
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4380
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   585
      Left            =   360
      TabIndex        =   10
      Top             =   3630
      Width           =   8655
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
         Height          =   315
         Left            =   105
         TabIndex        =   11
         Top             =   180
         Width           =   7815
         WordWrap        =   -1  'True
      End
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   7140
      TabIndex        =   8
      Top             =   420
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   714
   End
   Begin MSComCtl2.DTPicker DTDate 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "MMM YYYY"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   0
      EndProperty
      Height          =   315
      Left            =   1935
      TabIndex        =   0
      Top             =   1170
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
      CurrentDate     =   38734
   End
   Begin MSForms.Label Label1 
      Height          =   315
      Index           =   0
      Left            =   360
      TabIndex        =   20
      Top             =   1620
      Width           =   1515
      BackColor       =   16637923
      VariousPropertyBits=   8388627
      Caption         =   "Factory Code"
      Size            =   "2672;556"
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label1 
      Height          =   240
      Index           =   1
      Left            =   360
      TabIndex        =   19
      Top             =   1995
      Width           =   1515
      BackColor       =   16637923
      VariousPropertyBits=   8388627
      Caption         =   "Machine No."
      Size            =   "2672;423"
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label1 
      Height          =   315
      Index           =   2
      Left            =   360
      TabIndex        =   18
      Top             =   2370
      Width           =   1515
      BackColor       =   16637923
      VariousPropertyBits=   8388627
      Caption         =   "Material Code"
      Size            =   "2672;556"
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label1 
      Height          =   315
      Index           =   3
      Left            =   360
      TabIndex        =   17
      Top             =   2745
      Width           =   1515
      BackColor       =   16637923
      VariousPropertyBits=   8388627
      Caption         =   "Consumption Cls"
      Size            =   "2672;556"
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cboAction 
      Height          =   315
      Index           =   1
      Left            =   1935
      TabIndex        =   2
      Top             =   1965
      Width           =   1665
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "2937;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cboAction 
      Height          =   315
      Index           =   2
      Left            =   1935
      TabIndex        =   3
      Top             =   2340
      Width           =   2610
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "4604;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cboAction 
      Height          =   315
      Index           =   3
      Left            =   1935
      TabIndex        =   4
      Top             =   2715
      Width           =   1665
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "2937;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cboAction 
      Height          =   315
      Index           =   0
      Left            =   1935
      TabIndex        =   1
      Top             =   1590
      Width           =   1665
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "2937;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   3930
      X2              =   8010
      Y1              =   1890
      Y2              =   1890
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   3660
      X2              =   8010
      Y1              =   3015
      Y2              =   3015
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   4965
      X2              =   8000
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   3930
      X2              =   8010
      Y1              =   2265
      Y2              =   2265
   End
   Begin VB.Label lblDesc 
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
      Height          =   225
      Index           =   0
      Left            =   3930
      TabIndex        =   16
      Top             =   1665
      Width           =   4020
   End
   Begin VB.Label lblDesc 
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
      Height          =   225
      Index           =   1
      Left            =   3930
      TabIndex        =   15
      Top             =   2040
      Width           =   4020
   End
   Begin VB.Label lblDesc 
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
      Height          =   225
      Index           =   2
      Left            =   4965
      TabIndex        =   14
      Top             =   2415
      Width           =   3030
   End
   Begin VB.Label lblDesc 
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
      Height          =   225
      Index           =   3
      Left            =   3690
      TabIndex        =   13
      Top             =   2790
      Width           =   4290
   End
   Begin MSForms.Label Label2 
      Height          =   315
      Left            =   360
      TabIndex        =   12
      Top             =   1200
      Width           =   1515
      BackColor       =   16637923
      VariousPropertyBits=   8388627
      Caption         =   "Period"
      Size            =   "2672;556"
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Material Consumption Report"
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
      Left            =   510
      TabIndex        =   9
      Top             =   390
      Width           =   8355
   End
End
Attribute VB_Name = "frmEffMaterialConsumptionRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dateUp As Date

Private Sub cboAction_Click(Index As Integer)
  If cboAction(Index).ListIndex = -1 Then Exit Sub
  With cboAction(Index)
    lbldesc(Index).Caption = IIf(IsNull(.List(.ListIndex, 1)), "", .List(.ListIndex, 1))
  End With
End Sub

Private Sub cmdBrowser_Click()
 Me.MousePointer = vbHourglass
 frm_BrowseItem.getItemCode = cboAction(2).Text
 frm_BrowseItem.Show 1
 cboAction(2).Text = frm_BrowseItem.getItemCode
 Me.MousePointer = vbDefault
End Sub

Private Sub CmdExcel_Click()
  Dim i As Integer
  Dim s_sql As String
  LblErrMsg = ""
  For i = 0 To 3
    With cboAction(i)
      If .MatchFound = False And .Text <> "" Then
        Select Case i
          Case 0: LblErrMsg = DisplayMsg(4060): cboAction(i).SetFocus: Exit Sub
          Case 1: LblErrMsg = DisplayMsg(8015): cboAction(i).SetFocus: Exit Sub
          Case 2: LblErrMsg = DisplayMsg(8016): cboAction(i).SetFocus: Exit Sub
          Case 3: LblErrMsg = DisplayMsg(8017): cboAction(i).SetFocus: Exit Sub
        End Select
      ElseIf cboAction(i).ListIndex = -1 Then
        Select Case i
          Case 0: LblErrMsg = DisplayMsg(1060): cboAction(i).SetFocus: Exit Sub
          Case 1: LblErrMsg = DisplayMsg(8018): cboAction(i).SetFocus: Exit Sub
          Case 2: LblErrMsg = DisplayMsg(8019): cboAction(i).SetFocus: Exit Sub
          Case 3: LblErrMsg = DisplayMsg(8020): cboAction(i).SetFocus: Exit Sub
        End Select
      End If
    End With
  Next
  LblErrMsg = ""
  
  toExcel
End Sub

Private Sub CmdPreview_Click()

  Dim i As Integer
  Dim s_sql As String
  LblErrMsg = ""
  For i = 0 To 3
    With cboAction(i)
      If .MatchFound = False And .Text <> "" Then
        Select Case i
          Case 0: LblErrMsg = DisplayMsg(4060): cboAction(i).SetFocus: Exit Sub
          Case 1: LblErrMsg = DisplayMsg(8015): cboAction(i).SetFocus: Exit Sub
          Case 2: LblErrMsg = DisplayMsg(8016): cboAction(i).SetFocus: Exit Sub
          Case 3: LblErrMsg = DisplayMsg(8017): cboAction(i).SetFocus: Exit Sub
        End Select
      ElseIf cboAction(i).ListIndex = -1 Then
        Select Case i
          Case 0: LblErrMsg = DisplayMsg(1060): cboAction(i).SetFocus: Exit Sub
          Case 1: LblErrMsg = DisplayMsg(8018): cboAction(i).SetFocus: Exit Sub
          Case 2: LblErrMsg = DisplayMsg(8019): cboAction(i).SetFocus: Exit Sub
          Case 3: LblErrMsg = DisplayMsg(8020): cboAction(i).SetFocus: Exit Sub
        End Select
      End If
    End With
  Next
  LblErrMsg = ""
  
              Dim application As New CRAXDDRT.application
              Dim report As New CRAXDDRT.report
              Dim rsRpt As New ADODB.Recordset
              Dim Rpt As New FrmRpt3
              Dim sqlControl As String, RsInvControl As New ADODB.Recordset
              
              
              Me.MousePointer = vbHourglass
            
             Dim ls_where1 As String
             Dim ls_where2 As String
             Dim ls_where3 As String
             Dim ls_where4 As String
             
            For i = 0 To 3
              If cboAction(i).ListIndex <> -1 Then
                Select Case i:
                   Case 0:
                         If cboAction(i).Text <> "ALL" Then
                                ls_where1 = "And R.Supplier_Code = '" & Trim(cboAction(i).Text) & "'"
                        End If
                   Case 1:
                           If cboAction(i).Text <> "ALL" Then
                              ls_where2 = " And R.PO_No = '" & Trim(cboAction(i).Text) & "'"
                           End If
                   Case 2:
                           If cboAction(i).Text <> "ALL" Then
                              ls_where3 = " And ps.childitem_code = '" & Trim(cboAction(i).Text) & "'"
                           End If
                   Case 3:
                           If cboAction(i).Text <> "ALL" Then
                              ls_where4 = " And mc.materialconsump_cls = '" & Trim(cboAction(i).Text) & "'"
                           End If
                End Select
              End If
            Next
            
            sql = " select rtrim(manufacture_code)manufacture_code,          " & _
            " rtrim(manufacture_desc) manufacture_desc,          " & _
            " rtrim(line_code) line_code,           " & _
            " rtrim(line_name)line_name,           " & _
            " childitem_code,            " & _
            " item_name,                  " & _
            " materialconsump_cls,            " & _
            " Description , " & _
            " sum(qty01)qty01, " & _
            " sum(qty02)qty02, " & _
            " sum(qty03)qty03, "

sql = sql + " sum(qty04)qty04, " & _
            " sum(qty05)qty05, " & _
            " sum(qty06)qty06, " & _
            " sum(qty07)qty07, " & _
            " sum(qty08)qty08, " & _
            " sum(qty09)qty09, " & _
            " sum(qty10)qty10, " & _
            " sum(qty11)qty11, " & _
            " sum(qty12)qty12, " & _
            " sum(qty13)qty13, " & _
            " sum(qty14)qty14, "

sql = sql + " sum(qty15)qty15, " & _
            " sum(qty16)qty16, " & _
            " sum(qty17)qty17, " & _
            " sum(qty18)qty18, " & _
            " sum(qty19)qty19, " & _
            " sum(qty20)qty20, " & _
            " sum(qty21)qty21, " & _
            " sum(qty22)qty22, " & _
            " sum(qty23)qty23, " & _
            " sum(qty24)qty24, " & _
            " sum(qty25)qty25, "

sql = sql + " sum(qty26)qty26, " & _
            " sum(qty27)qty27, " & _
            " sum(qty28)qty28, " & _
            " sum(qty29)qty29, " & _
            " sum(qty30)qty30, " & _
            " sum(qty31)qty31 " & _
            " from " & _
            " ( " & _
            " select  " & _
            " manufacture_code,          " & _
            " manufacture_desc, "

sql = sql + " line_code,           " & _
            " line_name,           " & _
            " childitem_code,            " & _
            " item_name,            " & _
            " childsupply_date,            " & _
            " materialconsump_cls,            " & _
            " Description , " & _
            " Qty01=case when day(childsupply_date)=1 then qty else 0 end, " & _
            " Qty02=case when day(childsupply_date)=2 then qty else 0 end, " & _
            " Qty03=case when day(childsupply_date)=3 then qty else 0 end, " & _
            " Qty04=case when day(childsupply_date)=4 then qty else 0 end, "

sql = sql + " Qty05=case when day(childsupply_date)=5 then qty else 0 end, " & _
            " Qty06=case when day(childsupply_date)=6 then qty else 0 end, " & _
            " Qty07=case when day(childsupply_date)=7 then qty else 0 end, " & _
            " Qty08=case when day(childsupply_date)=8 then qty else 0 end, " & _
            " Qty09=case when day(childsupply_date)=9 then qty else 0 end, " & _
            " Qty10=case when day(childsupply_date)=10 then qty else 0 end, " & _
            " Qty11=case when day(childsupply_date)=11 then qty else 0 end, " & _
            " Qty12=case when day(childsupply_date)=12 then qty else 0 end, " & _
            " Qty13=case when day(childsupply_date)=13 then qty else 0 end, " & _
            " Qty14=case when day(childsupply_date)=14 then qty else 0 end, " & _
            " Qty15=case when day(childsupply_date)=15 then qty else 0 end, "

sql = sql + " Qty16=case when day(childsupply_date)=16 then qty else 0 end, " & _
            " Qty17=case when day(childsupply_date)=17 then qty else 0 end, " & _
            " Qty18=case when day(childsupply_date)=18 then qty else 0 end, " & _
            " Qty19=case when day(childsupply_date)=19 then qty else 0 end, " & _
            " Qty20=case when day(childsupply_date)=20 then qty else 0 end, " & _
            " Qty21=case when day(childsupply_date)=21 then qty else 0 end, " & _
            " Qty22=case when day(childsupply_date)=22 then qty else 0 end, " & _
            " Qty23=case when day(childsupply_date)=23 then qty else 0 end, " & _
            " Qty24=case when day(childsupply_date)=24 then qty else 0 end, " & _
            " Qty25=case when day(childsupply_date)=25 then qty else 0 end, " & _
            " Qty26=case when day(childsupply_date)=26 then qty else 0 end, "

sql = sql + " Qty27=case when day(childsupply_date)=27 then qty else 0 end, " & _
            " Qty28=case when day(childsupply_date)=28 then qty else 0 end, " & _
            " Qty29=case when day(childsupply_date)=29 then qty else 0 end, " & _
            " Qty30=case when day(childsupply_date)=30 then qty else 0 end, " & _
            " Qty31=case when day(childsupply_date)=31 then qty else 0 end " & _
            " from  " & _
            " ( " & _
            "    select manufacture_code = R.Supplier_Code,      " & _
            "   tm.trade_name manufacture_desc, " & _
            "    line_code = R.PO_No,      ps.childitem_code,      im.item_name,      " & _
            "    ps.childsupply_date,      qty = sum(isnull(ps.consumption_qty,0)),      mc.materialconsump_cls,       "

sql = sql + "   mc.Description,      ml.line_name  " & _
            " from part_supply ps " & _
                "left join Part_Receipt R on Convert(Char, R.Seq_No) = ps.DO_No And R.ProductionResult_Cls = 1 " & _
                "left join item_master im on im.item_code = ps.childitem_code " & _
                "left join manufacture_line ml on ml.Manufacture_Code = R.Supplier_Code And ml.Line_Code = R.PO_No " & _
                "left join  materialconsump_cls mc on   ps.materialconsump_cls = mc.materialconsump_cls " & _
                "left join trade_master tm  on  tm.trade_code=ml.manufacture_code " & _
            "Where ( ps.do_no is not null and ps.do_no <> '')  " & _
                "  And datepart(month,ps.childsupply_date) = '" & Format(dtDate, "MM") & "'  " & _
                "  And datepart(year,ps.childsupply_date) = '" & Format(dtDate, "yyyy") & "'  " & _
            ls_where1 & " " & _
            ls_where2 & " " & _
            ls_where3 & " " & _
            ls_where4 & " " & _
            "    " & _
            "   group by R.Supplier_Code,  trade_name,         R.PO_No,           ml.line_name,           ps.childitem_code,           im.item_name,           ps.childsupply_date,           mc.materialconsump_cls,           mc.Description  " & _
            "    " & _
            " )tbA " & _
            " )tbB " & _
            " group by  " & _
            " manufacture_code,         " & _
            " manufacture_desc,  "

sql = sql + " line_code,           " & _
            " line_name,           " & _
            " childitem_code,            " & _
            " item_name,              " & _
            " materialconsump_cls,            " & _
            " Description  "

            
              If rsRpt.State <> adStateClosed Then rsRpt.Close
              rsRpt.Open sql, Db, adOpenDynamic, adLockOptimistic
            
              If rsRpt.EOF Then LblErrMsg.Caption = DisplayMsg(4006): Me.MousePointer = vbDefault: Exit Sub
            
                        
              Set report = application.OpenReport(App.path & "\Reports\rptMaterialConsumption.rpt")
              report.Database.Tables(1).SetDataSource rsRpt
                       
            
             reportcode = "materialConsumption"
             printorient = "2"
             sqlprint = sql
             sqlprint2 = Format(dtDate.Value, "yyyy-MM-01")
             report.FormulaFields(34).Text = "'" & CDate(sqlprint2) & "'"
             F_Factory = cboAction(0).Text & " / " & lbldesc(0).Caption
             report.FormulaFields(35).Text = "'" & F_Factory & "'"
             report.ReportTitle = "Material Consumption Report"
            
              Rpt.CRViewer1.ReportSource = report
              Rpt.CRViewer1.ViewReport
              Rpt.CRViewer1.Zoom 1
            
              Rpt.WindowState = 2
              Rpt.Show 1
            
              Me.MousePointer = vbDefault
  
End Sub

Private Sub CmdSubMenu_Click()
  frmMainMenu.Show
  Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub DTDate_Change()
If Format(dtDate.Value, "MM") < Format(dateUp, "MM") And Val(Format(dtDate.Value, "MM")) = 1 And Val(Format(dateUp, "MM")) = 12 Then _
            dtDate.Year = dtDate.Year + 1: GoTo pass
    If Format(dtDate.Value, "MM") > Format(dateUp, "MM") And Val(Format(dtDate.Value, "MM")) = 12 And Val(Format(dateUp, "MM")) = 1 Then _
            dtDate.Year = dtDate.Year - 1
pass:
    dateUp = Format(dtDate.Value, "dd MMM yyyy")
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
  Me.Caption = Me.Caption & " (Menu ID : " & frmcode("frmMaterialConsumptionRpt") & ")"
  IsiCombo
  dtDate.Value = Format(Now, "MMM yyyy")
  dateUp = dtDate.Value
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = 0 Then Cancel = 1
End Sub

Sub toExcel()
 Dim xlapp As New Excel.application
 Dim Idx As Long
 Dim rsJoin As New ADODB.Recordset
 Dim ls_sql As String
 Dim ls_where As String
 Dim dayOfMonth As Long
 Dim i As Integer
 Dim j As Integer
 Dim maxCol As String
 Dim ls_group As String
 Dim ls_orderby As String
 Dim temp_line As String
 Dim temp_item As String
 Dim temp_consumption As String
 Dim subTotalMaterial(1 To 31) As Double
 Dim subTotalMachine(1 To 31) As Double
 Dim GrandTotalMachine(1 To 31) As Double
 Dim totPerRow(1 To 4) As Double
 Dim A_Idx As Integer
 Dim idx2 As Integer


 
  
 Screen.MousePointer = vbHourglass
          
 ' jumlah hari dalam 1 bulan
 With dtDate
   Select Case Format(.Value, "MM")
      Case 1, 3, 5, 7, 8, 10, 12: dayOfMonth = 31
                 
      Case 2: ' kabisat
              If Format(.Value, "MM") Mod 4 = 0 Then
                 dayOfMonth = 29
              Else
                 dayOfMonth = 28
              End If
      Case 4, 6, 9, 11: dayOfMonth = 30
   End Select
 End With
 
 
 ls_sql = " select manufacture_code = R.Supplier_Code,  tm.Trade_name Manufacture_Name," & _
                    "     line_code = R.PO_No, " & _
                    "     ps.childitem_code, " & _
                    "     im.item_name, " & _
                    "     ps.childsupply_date, " & _
                    "     qty = sum(isnull(ps.consumption_qty,0)), " & _
                    "     mc.materialconsump_cls, " & _
                    "     mc.Description, " & _
                    "     ml.line_name " & _
                    " from part_supply ps " & _
                    " left join Part_Receipt R on convert(char, R.Seq_No) = ps.DO_No And R.ProductionResult_Cls = 1 " & _
                    " left join item_master im on im.Item_Code = ps.ChildItem_Code " & _
                    " left join materialconsump_cls mc on mc.materialconsump_cls = ps.materialconsump_cls " & _
                    " left join manufacture_line ml on ml.Line_Code = R.PO_No " & _
                    " left join trade_master tm on R.Supplier_Code = tm.trade_code " & _
                    " Where  ( ps.do_no is not null and ps.do_no <> '') " & _
                        " And datepart(month,ps.childsupply_date) = '" & Format(dtDate, "MM") & "' " & _
                        " And datepart(year,ps.childsupply_date) = '" & Format(dtDate, "YYYY") & "' " '& _

 ls_group = " group by R.Supplier_Code, tm.Trade_name, " & _
                    "          R.PO_No, " & _
                    "          ml.line_name, " & _
                    "          ps.childitem_code, " & _
                    "          im.item_name, " & _
                    "          ps.childsupply_date, " & _
                    "          mc.materialconsump_cls, " & _
                    "          mc.Description "
 ls_orderby = " order by R.Supplier_Code, " & _
                    "          R.PO_No, " & _
                    "          ps.childitem_code, " & _
                    "          mc.materialconsump_cls, " & _
                    "          ml.line_name, " & _
                    "          im.item_name, " & _
                    "          mc.Description, " & _
                    "          ps.childsupply_date "
 ' filter
 For i = 0 To 3
   If cboAction(i).ListIndex <> -1 Then
     Select Case i:
        Case 0:
                If cboAction(i).Text <> "ALL" Then
                    ls_where = "And R.Supplier_Code = '" & Trim(cboAction(i).Text) & "'"
                End If
        Case 1:
                If cboAction(i).Text <> "ALL" Then
                   ls_where = ls_where & " And R.PO_No = '" & Trim(cboAction(i).Text) & "'"
                End If
        Case 2:
                If cboAction(i).Text <> "ALL" Then
                   ls_where = ls_where & " And ps.childitem_code = '" & Trim(cboAction(i).Text) & "'"
                End If
        Case 3:
                If cboAction(i).Text <> "ALL" Then
                   ls_where = ls_where & " And mc.materialconsump_cls = '" & Trim(cboAction(i).Text) & "'"
                End If
     End Select
   End If
 Next
 
 ls_sql = ls_sql & ls_where & ls_group & ls_orderby
  
 If rsJoin.State = 1 Then rsJoin.Close
 rsJoin.CursorLocation = adUseClient
 rsJoin.Open ls_sql, Db, 1, 3

If Not rsJoin.EOF Then
    Screen.MousePointer = vbHourglass
    With xlapp
        .Workbooks.Add
        
        'Header
        
        .Range("b1") = "Factory Code"
        .Range("c1", "h1").Merge
        .Range("c1") = " : " & cboAction(0).Text & " / " & lbldesc(0).Caption
        
        .Range("b2") = "Period "
        .Range("c2", "f2").Merge
        .Range("c2") = " : " & Format(dtDate.Value, "MMM/YYYY")
        .Range("b1", "h2").Font.Bold = True
        
        .Range("b4", "d4").Merge
        .Range("b4") = "Machine No."
        .Range("b5", "d5").Merge
        .Range("b5") = "Material Code"
        
        .Range("e4", "e5").Merge
        .Range("e4") = "Consumption Classification"
        .Range("e4").horizontalAlignment = xlCenter
        .Range("e4").verticalAlignment = xlCenter
        
        .Range("f5") = "Total"
        .Range("f5").horizontalAlignment = xlCenter
        .Range("f5").verticalAlignment = xlCenter

        ' set tanggal days of month header
        For i = 1 To dayOfMonth
          If Asc("f") + i > Asc("z") Then
             .Range("a" & Chr((Asc("a") + j)) & "5") = i
             maxCol = "a" & Chr((Asc("a") + j))
             j = j + 1
          ElseIf Asc("f") + i <= Asc("z") Then
             .Range(Chr((Asc("f") + i)) & "5") = i
          End If
        Next i
        

        .Range("f4", maxCol & "4").Merge
        .Range("f4", maxCol & "4") = "Days of Month"


        ' Allign Center
        .Range("b4", maxCol & "5").horizontalAlignment = xlCenter
    'End Header
    
    'Content
         Idx = 5
         rsJoin.MoveFirst
         While Not rsJoin.EOF
'         Idx = Idx + 1
'         .Range("a" & Idx) = rsJoin("childitem_code")
         
           A_Idx = Format(rsJoin("childsupply_date"), "DD")
            
            If temp_line <> rsJoin("line_code") Then
               'seting content1 (isi machine no.)
                 Idx = Idx + 1
                 .Range("b" & Idx, "c" & Idx).Merge
                 '.Range("b" & Idx) = "Machine No."
                 .Range("b" & Idx) = "Factory Code / Line Code"
                 .Range("b" & Idx).horizontalAlignment = xlCenter
                 .Range("d" & Idx, maxCol & Idx).Merge
                 .Range("d" & Idx) = ": " & Trim(rsJoin("Manufacture_Code")) & " (" & Trim(rsJoin("Manufacture_name")) & ") / " & Trim(rsJoin("line_code")) & " (" & Trim(rsJoin("line_name")) & ")"
                  temp_line = rsJoin("line_code")
                  
                  totPerRow(3) = rsJoin("qty")
            Else
                  totPerRow(3) = totPerRow(3) + rsJoin("qty") ' total per row per machine
            End If

           If temp_consumption <> rsJoin("materialconsump_cls") Then
                If temp_item = rsJoin("childitem_code") Then
                   temp_consumption = rsJoin("materialconsump_cls")
                   Idx = Idx + 1
                   totPerRow(2) = totPerRow(2) + rsJoin("qty") ' total per row per item code
                Else
                   totPerRow(2) = rsJoin("qty")
                   temp_consumption = rsJoin("materialconsump_cls")
                End If
                totPerRow(1) = rsJoin("qty") ' total per row per consumption cls
           Else
               If temp_item = rsJoin("childitem_code") Then
                  totPerRow(1) = totPerRow(1) + rsJoin("qty")
                  totPerRow(2) = totPerRow(2) + rsJoin("qty")
               Else
                  totPerRow(1) = rsJoin("qty")
                  totPerRow(2) = rsJoin("qty")
               End If
           End If
           
           totPerRow(4) = totPerRow(4) + rsJoin("qty") ' total per row grand total
           
           If temp_item <> rsJoin("childitem_code") _
                  And temp_line = rsJoin("line_code") Then
                'setting content2 (isi material code)
                 Idx = Idx + 1
                .Range("b" & Idx, "c" & Idx).Merge
                .Range("b" & Idx) = "Material Code"
                .Range("b" & Idx).horizontalAlignment = xlCenter
                .Range("d" & Idx, maxCol & Idx).Merge
                .Range("d" & Idx) = ": " & Trim(rsJoin("childitem_code")) & " / " & Trim(rsJoin("item_name"))
                temp_item = rsJoin("childitem_code")
                Idx = Idx + 1
                idx2 = Idx
           End If
           
           '1 . Cetak Total Per row per Comsumption Cls
           .Range("f" & Idx) = totPerRow(1)
           .Range("f" & Idx).NumberFormat = gs_formatQty
           .Range("f" & Idx).horizontalAlignment = xlRight
           
           'Content detail (isi Detail) -> masih kecil dari colum "Z"
           If Asc("f") + A_Idx > Asc("z") Then
              .Range("a" & Chr((Asc("f") + A_Idx - 26)) & Idx) = Format(rsJoin("qty"), gs_formatQty)
              .Range("a" & Chr((Asc("f") + A_Idx - 26)) & Idx).NumberFormat = gs_formatQty
              .Range("e" & Idx) = Trim(rsJoin("materialconsump_cls")) & " / " & Trim(rsJoin("description"))
              .Range("a" & Chr((Asc("f") + A_Idx - 26)) & Idx).horizontalAlignment = xlRight
           Else 'Content detail (isi Detail) -> lebih besar dari colum "Z"
              .Range(Chr(Asc("f") + A_Idx) & Idx) = Format(rsJoin("qty"), gs_formatQty)
              .Range(Chr(Asc("f") + A_Idx) & Idx).NumberFormat = gs_formatQty
              .Range(Chr(Asc("f") + A_Idx) & Idx).horizontalAlignment = xlRight
              .Range("e" & Idx) = Trim(rsJoin("materialconsump_cls")) & " / " & Trim(rsJoin("description"))
           End If
           
             ' hitung subtotal material, sub total machine, grand total
               subTotalMaterial(A_Idx) = _
                 subTotalMaterial(A_Idx) + rsJoin("qty")
               subTotalMachine(A_Idx) = _
                 subTotalMachine(A_Idx) + rsJoin("qty")
               GrandTotalMachine(A_Idx) = _
                 GrandTotalMachine(A_Idx) + rsJoin("qty")
                           
           rsJoin.MoveNext
                     
           ' if empty then set "-"
            If Not rsJoin.EOF Then
                If Trim(temp_consumption) <> IIf(rsJoin.EOF, "", Trim(rsJoin("materialconsump_cls"))) _
                      Or Trim(temp_item) <> IIf(rsJoin.EOF, "", Trim(rsJoin("childitem_code"))) Then
                     i = 1
                     j = 0
                     For i = 1 To dayOfMonth
                      If Asc("f") + i > Asc("z") Then
                          If .Range("a" & Chr((Asc("a") + j)) & Idx) = "" Then .Range("a" & Chr((Asc("a") + j)) & Idx) = "-": .Range("a" & Chr((Asc("a") + j)) & Idx).horizontalAlignment = xlCenter
                          j = j + 1
                       ElseIf Asc("f") + i <= Asc("z") Then
                          If .Range(Chr((Asc("f") + i)) & Idx) = "" Then .Range(Chr((Asc("f") + i)) & Idx) = "-": .Range(Chr((Asc("f") + i)) & Idx).horizontalAlignment = xlCenter
                       End If
                     Next i
                End If
             Else
                i = 1
                j = 0
                For i = 1 To dayOfMonth
                  If Asc("f") + i > Asc("z") Then
                     If .Range("a" & Chr((Asc("a") + j)) & Idx) = "" Then .Range("a" & Chr((Asc("a") + j)) & Idx) = "-": .Range("a" & Chr((Asc("a") + j)) & Idx).horizontalAlignment = xlCenter
                     j = j + 1
                  ElseIf Asc("f") + i <= Asc("z") Then
                     If .Range(Chr((Asc("f") + i)) & Idx) = "" Then .Range(Chr((Asc("f") + i)) & Idx) = "-": .Range(Chr((Asc("f") + i)) & Idx).horizontalAlignment = xlCenter
                  End If
                Next i
             End If
             
           If Not rsJoin.EOF Then
               
               
               ' end of group material (item code)
               If temp_item <> rsJoin("childitem_code") _
                      And temp_line = rsJoin("line_code") Then
                    Idx = Idx + 1
                    .Range("e" & Idx) = "Sub Total Material"
                    .Range("e" & Idx).Font.Bold = True
                    .Range("e" & Idx).horizontalAlignment = xlCenter
                      
                    '2 . Cetak total per row per item code
                   .Range("f" & Idx) = totPerRow(2)
                   .Range("f" & Idx).NumberFormat = gs_formatQty
                   .Range("f" & Idx).horizontalAlignment = xlRight
                
                    'Cetak subtotal material
                    i = 1
                    j = 0
                    For i = 1 To dayOfMonth
                     If Asc("f") + i > Asc("z") Then
                         If subTotalMaterial(i) <> 0 Then
                           .Range("a" & Chr((Asc("a") + j)) & Idx) = Format(subTotalMaterial(i), gs_formatQty)
                           .Range("a" & Chr((Asc("a") + j)) & Idx).NumberFormat = gs_formatQty
                           .Range("a" & Chr((Asc("a") + j)) & Idx).horizontalAlignment = xlRight
                         Else
                           .Range("a" & Chr((Asc("a") + j)) & Idx) = "-"
                           .Range("a" & Chr((Asc("a") + j)) & Idx).horizontalAlignment = xlCenter
                         End If
                         
                         subTotalMaterial(i) = 0
                         j = j + 1
                      ElseIf Asc("f") + i <= Asc("z") Then
                          If subTotalMaterial(i) <> 0 Then
                            .Range(Chr((Asc("f") + i)) & Idx) = Format(subTotalMaterial(i), gs_formatQty)
                            .Range(Chr((Asc("f") + i)) & Idx).horizontalAlignment = xlRight
                            .Range(Chr((Asc("f") + i)) & Idx).NumberFormat = gs_formatQty
                          Else ' jika nol cetak "-"
                            .Range(Chr((Asc("f") + i)) & Idx) = "-"
                            .Range(Chr((Asc("f") + i)) & Idx).horizontalAlignment = xlCenter
                          End If
                         subTotalMaterial(i) = 0
                      End If
                    Next i
                   .Range("b" & idx2, "d" & Idx).Merge
               End If

               ' end of group line code (machine no.)
               If temp_line <> rsJoin("line_code") And temp_line <> "" Then

                     Idx = Idx + 1
                    .Range("e" & Idx) = "Sub Total Material"
                    .Range("e" & Idx).Font.Bold = True
                    .Range("e" & Idx).horizontalAlignment = xlCenter
                    
                    '2 . Cetak total per row per consumption cls
                   .Range("f" & Idx) = totPerRow(2)
                   .Range("f" & Idx).NumberFormat = gs_formatQty
                   .Range("f" & Idx).horizontalAlignment = xlRight
                   
                    ' cetak subtotalmaeterial & subtotalmachine
                    i = 1
                    j = 0
                    For i = 1 To dayOfMonth
                      If Asc("f") + i > Asc("z") Then
                         If subTotalMaterial(i) <> 0 Then
                           .Range("a" & Chr((Asc("a") + j)) & Idx) = Format(subTotalMaterial(i), gs_formatQty)
                           .Range("a" & Chr((Asc("a") + j)) & Idx).NumberFormat = gs_formatQty
                           .Range("a" & Chr((Asc("a") + j)) & Idx).horizontalAlignment = xlRight
                           
                           .Range("a" & Chr((Asc("a") + j)) & Idx + 1) = Format(subTotalMachine(i), gs_formatQty)
                           .Range("a" & Chr((Asc("a") + j)) & Idx + 1).NumberFormat = gs_formatQty
                           .Range("a" & Chr((Asc("a") + j)) & Idx + 1).horizontalAlignment = xlRight
                           
                         Else
                           .Range("a" & Chr((Asc("a") + j)) & Idx) = "-"
                           .Range("a" & Chr((Asc("a") + j)) & Idx).horizontalAlignment = xlCenter
                           .Range("a" & Chr((Asc("a") + j)) & Idx + 1) = "-"
                            .Range("a" & Chr((Asc("a") + j)) & Idx + 1).horizontalAlignment = xlCenter
                         End If
                          
                         subTotalMaterial(i) = 0
                         subTotalMachine(i) = 0
                         j = j + 1
                      ElseIf Asc("f") + i <= Asc("z") Then
                         If subTotalMaterial(i) <> 0 Then
                           .Range(Chr((Asc("f") + i)) & Idx) = Format(subTotalMaterial(i), gs_formatQty)
                           .Range(Chr((Asc("f") + i)) & Idx).NumberFormat = gs_formatQty
                           .Range(Chr((Asc("f") + i)) & Idx).horizontalAlignment = xlRight
                           
                           .Range(Chr((Asc("f") + i)) & Idx + 1) = Format(subTotalMachine(i), gs_formatQty)
                           .Range(Chr((Asc("f") + i)) & Idx + 1).NumberFormat = gs_formatQty
                           .Range(Chr((Asc("f") + i)) & Idx + 1).horizontalAlignment = xlRight
                         Else
                           .Range(Chr((Asc("f") + i)) & Idx) = "-"
                           .Range(Chr((Asc("f") + i)) & Idx).horizontalAlignment = xlCenter
                           .Range(Chr((Asc("f") + i)) & Idx + 1) = "-"
                           .Range(Chr((Asc("f") + i)) & Idx + 1).horizontalAlignment = xlCenter
                         End If
                         
                         subTotalMaterial(i) = 0
                         subTotalMachine(i) = 0
                      End If
                    Next i

                    On Error Resume Next
                    .Range("b" & idx2, "d" & Idx).Merge
                   If err.Description = "" Then
                        Idx = Idx + 1
                        .Range("b" & Idx, "e" & Idx).Merge
                        .Range("b" & Idx) = "Sub Total Machine No."
                        .Range("b" & Idx).Font.Bold = True
                        .Range("b" & Idx).horizontalAlignment = xlRight
                        
                        '3 . Cetak total per row per machine
                        .Range("f" & Idx) = totPerRow(3)
                        .Range("f" & Idx).NumberFormat = gs_formatQty
                        .Range("f" & Idx).horizontalAlignment = xlRight
                    Else
                        Idx = Idx + 1
                    End If
               End If
            Else  'jika sudah end of file, maka cetak subtotalmaterial, subtotalmachine, gradtotalamchine
                  'setting contentfooter all

                     Idx = Idx + 1
                    .Range("e" & Idx) = "Sub Total Material"
                    .Range("e" & Idx).Font.Bold = True
                    .Range("e" & Idx).horizontalAlignment = xlCenter
                    
                    '2 . Cetak total per row per item code
                    .Range("f" & Idx) = Format(totPerRow(2), gs_formatQty)
                    .Range("f" & Idx).NumberFormat = gs_formatQty
                    .Range("f" & Idx).horizontalAlignment = xlRight
                    
                    'cetak subTotalMaterial, subTotalMachine, gradtotalamchine
                    i = 1
                    j = 0
                    For i = 1 To dayOfMonth
                      If Asc("f") + i > Asc("z") Then
                         If subTotalMaterial(i) <> 0 Then
                           ' subtotal material
                           .Range("a" & Chr((Asc("a") + j)) & Idx) = Format(subTotalMaterial(i), gs_formatQty)
                           .Range("a" & Chr((Asc("a") + j)) & Idx).NumberFormat = gs_formatQty
                           .Range("a" & Chr((Asc("a") + j)) & Idx).horizontalAlignment = xlRight
                           'subtotal machine
                           .Range("a" & Chr((Asc("a") + j)) & Idx + 1) = Format(subTotalMachine(i), gs_formatQty)
                           .Range("a" & Chr((Asc("a") + j)) & Idx + 1).NumberFormat = gs_formatQty
                           .Range("a" & Chr((Asc("a") + j)) & Idx + 1).horizontalAlignment = xlRight
                           'grand total
                           .Range("a" & Chr((Asc("a") + j)) & Idx + 2).NumberFormat = gs_formatQty
                           .Range("a" & Chr((Asc("a") + j)) & Idx + 2) = Format(GrandTotalMachine(i), gs_formatQty)
                           .Range("a" & Chr((Asc("a") + j)) & Idx + 2).horizontalAlignment = xlRight
                         Else
                           .Range("a" & Chr((Asc("a") + j)) & Idx) = "-"
                           .Range("a" & Chr((Asc("a") + j)) & Idx).horizontalAlignment = xlCenter
                           .Range("a" & Chr((Asc("a") + j)) & Idx + 1) = "-"
                           .Range("a" & Chr((Asc("a") + j)) & Idx + 1).horizontalAlignment = xlCenter
                           .Range("a" & Chr((Asc("a") + j)) & Idx + 2) = "-"
                           .Range("a" & Chr((Asc("a") + j)) & Idx + 2).horizontalAlignment = xlCenter
                         End If
                         
                         subTotalMaterial(i) = 0
                         subTotalMachine(i) = 0
                         GrandTotalMachine(i) = 0
                         j = j + 1
                      ElseIf Asc("f") + i <= Asc("z") Then
                         If subTotalMaterial(i) <> 0 Then
                            .Range(Chr((Asc("f") + i)) & Idx) = Format(subTotalMaterial(i), gs_formatQty)
                            .Range(Chr((Asc("f") + i)) & Idx).NumberFormat = gs_formatQty
                            .Range(Chr((Asc("f") + i)) & Idx).horizontalAlignment = xlRight
                            
                            .Range(Chr((Asc("f") + i)) & Idx + 1).NumberFormat = gs_formatQty
                            .Range(Chr((Asc("f") + i)) & Idx + 1).horizontalAlignment = xlRight
                            .Range(Chr((Asc("f") + i)) & Idx + 1) = Format(subTotalMachine(i), gs_formatQty)
                            
                            .Range(Chr((Asc("f") + i)) & Idx + 2) = Format(GrandTotalMachine(i), gs_formatQty)
                            .Range(Chr((Asc("f") + i)) & Idx + 2).NumberFormat = gs_formatQty
                            .Range(Chr((Asc("f") + i)) & Idx + 2).horizontalAlignment = xlRight
                         Else
                           .Range(Chr((Asc("f") + i)) & Idx) = "-"
                           .Range(Chr((Asc("f") + i)) & Idx).horizontalAlignment = xlCenter
                           
                           .Range(Chr((Asc("f") + i)) & Idx + 1).horizontalAlignment = xlCenter
                           .Range(Chr((Asc("f") + i)) & Idx + 1) = "-"
                         
                           .Range(Chr((Asc("f") + i)) & Idx + 2) = "-"
                           .Range(Chr((Asc("f") + i)) & Idx + 2).horizontalAlignment = xlCenter
                         End If
                         subTotalMaterial(i) = 0
                         subTotalMachine(i) = 0
                         GrandTotalMachine(i) = 0
                      End If
                    Next
                    On Error Resume Next
                    .Range("b" & idx2, "d" & Idx).Merge
                    
                   ' If Err.Description = "" Then
                        Idx = Idx + 1
                        .Range("b" & Idx, "e" & Idx).Merge
                        .Range("b" & Idx) = "Sub Total Machine No."
                        .Range("b" & Idx).horizontalAlignment = xlRight
                        .Range("b" & Idx).Font.Bold = True
                        
                        '3 . Cetak total per row per machine
                        .Range("f" & Idx) = Format(totPerRow(3), gs_formatQty)
                        .Range("f" & Idx).NumberFormat = gs_formatQty
                        .Range("f" & Idx).horizontalAlignment = xlRight
                        
                        Idx = Idx + 1
                        .Range("b" & Idx, "e" & Idx).Merge
                        .Range("b" & Idx) = "Grand Total"
                        .Range("b" & Idx).Font.Bold = True
                        .Range("b" & Idx).horizontalAlignment = xlRight
                        
                         '4 . Cetak total per row tuk grand total
                        .Range("f" & Idx) = Format(totPerRow(4), gs_formatQty)
                        .Range("f" & Idx).NumberFormat = gs_formatQty
                        .Range("f" & Idx).horizontalAlignment = xlRight
'                    Else
                        Idx = Idx + 1
'                    End If
            End If

         Wend
         
        .Range("f5") = "Total"
        .Range("f5").horizontalAlignment = xlCenter
        .Range("f5").verticalAlignment = xlCenter
        
        ' set tanggal days of month header
        j = 0
        For i = 1 To dayOfMonth
          If Asc("f") + i > Asc("z") Then
             .Range("a" & Chr((Asc("a") + j)) & "5") = i
             maxCol = "a" & Chr((Asc("a") + j))
             j = j + 1
          ElseIf Asc("f") + i <= Asc("z") Then
             .Range(Chr((Asc("f") + i)) & "5") = i
          End If
          .Range(Chr((Asc("f") + i)) & "5").NumberFormat = "#0"
          .Range(Chr((Asc("f") + i)) & "5").horizontalAlignment = xlCenter
          .Range(Chr((Asc("f") + i)) & "5").verticalAlignment = xlCenter
        Next i
    

        .Range("b1", maxCol & Idx).Font.Name = "Arial"
        .Range("b1", maxCol & Idx).Font.Size = 8
        .Range("b4", maxCol & Idx).Borders.LineStyle = xlContinuous
        
        .Visible = True
        .ActiveSheet.PageSetup.PaperSize = xlPaperA4
        .ActiveSheet.PageSetup.Orientation = 2
        .ActiveSheet.PageSetup.LeftMargin = application.InchesToPoints(0.4)
        .ActiveSheet.PageSetup.RightMargin = application.InchesToPoints(0.4)
        .Range("A:E").Columns.AutoFit
        .WindowState = xlMaximized
   End With
Else
    rsJoin.Close
    Set rsJoin = Nothing
    LblErrMsg = DisplayMsg(4006)
End If
Screen.MousePointer = vbDefault
If rsJoin.State = 1 Then rsJoin.Close: Set rsJoin = Nothing
End Sub


Sub IsiCombo()

 Dim SSql As String
 Dim i As Long
 
 'Isi ComboBox Factory
 Call up_FillCombo(cboAction(0), "manufacture_line ml", "distinct ml.manufacture_code, tm.trade_name", ", Trade_master tm where  ml.manufacture_code = tm.trade_code order by ml.manufacture_code", True)
cboAction(0).ColumnWidths = "80 pt; 200 pt"
cboAction(0).ListWidth = 280
cboAction(0).ListRows = 15

' Dim rsFactory As Recordset
' Set rsFactory = New Recordset
' SSql = " select distinct ml.manufacture_code, tm.trade_name from manufacture_line ml, Trade_master tm" & _
'        " where  ml.manufacture_code = tm.trade_code " & _
'        " order by ml.manufacture_code"
' Set rsFactory = Db.Execute(SSql)
'
' With cboAction(0)
'      .ColumnCount = 2
'      i = 0
'      Do While Not rsFactory.EOF
'           .AddItem ""
'           .List(i, 0) = Trim(rsFactory("manufacture_code"))
'           .List(i, 1) = Trim(rsFactory("trade_name"))
'           i = i + 1
'           rsFactory.MoveNext
'      Loop
'      .ColumnWidths = "50 pt; 200 pt"
'      .ListWidth = 250
'      .ListRows = 15
' End With

 'Isi ComboBox Manufacture Line
 Dim rsMachine As Recordset
 Set rsMachine = New Recordset
 SSql = "select * from Manufacture_Line order by manufacture_code, Line_code"
    Set rsMachine = Db.Execute(SSql)
    With cboAction(1)
        .clear
        .columnCount = 2
        .ColumnWidths = "80pt;170pt"
        .ListWidth = 250
        .ListRows = 15
        .TextColumn = 1
        
        .AddItem ""
        .List(0, 0) = "ALL"
        .List(0, 1) = "ALL"
        i = 1
        Do While Not rsMachine.EOF
            .AddItem
            .List(i, 0) = Trim(rsMachine("Line_code"))
            .List(i, 1) = Trim(rsMachine("Line_Name"))
            rsMachine.MoveNext
            i = i + 1
        Loop
        
    End With

 
 'Isi Combobox Product Code
 Dim rsItemMaster As Recordset
 Set rsItemMaster = New Recordset
 SSql = "  select item_code, Item_name from item_master where use_endday >= convert(char(8), getdate(), 112) " & _
          " order by item_code"
 Set rsItemMaster = Db.Execute(SSql)
 
 With cboAction(2)
        .clear
        .columnCount = 2
        .ColumnWidths = "150pt;250pt"
        .ListWidth = 400
        .ListRows = 15
        
        .TextColumn = 1
        .AddItem ""
        .List(0, 0) = "ALL"
        .List(0, 1) = "ALL"
        i = 1
        Do While Not rsItemMaster.EOF
            .AddItem
            .List(i, 0) = Trim(rsItemMaster("item_code"))
            .List(i, 1) = Trim(rsItemMaster("Item_name"))
            rsItemMaster.MoveNext
            i = i + 1
        Loop
  End With

 'Isi ComboBox comsumption CLS
 Dim rsConsumption As Recordset
 Set rsConsumption = New Recordset
 
 SSql = "select * from MaterialConsump_Cls"
 Set rsConsumption = Db.Execute(SSql)
 
 With cboAction(3)
        .clear
        .columnCount = 2
        .ColumnWidths = "50pt;100pt;"
        .ListWidth = 150
        .ListRows = 15
        .TextColumn = 1
        
        .AddItem ""
        .List(0, 0) = "ALL"
        .List(0, 1) = "ALL"
        
        i = 1
        Do While Not rsConsumption.EOF
            .AddItem
            .List(i, 0) = Trim(rsConsumption("MaterialConsump_Cls"))
            .List(i, 1) = Trim(rsConsumption("Description"))
            rsConsumption.MoveNext
            i = i + 1
        Loop
  End With

For i = 0 To 3
  cboAction(i).ListIndex = 0
Next
End Sub

