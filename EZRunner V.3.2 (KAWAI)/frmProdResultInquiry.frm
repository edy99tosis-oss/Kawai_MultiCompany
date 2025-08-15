VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmProdResultInquiry 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Production Schedule/Result Inquiry"
   ClientHeight    =   10560
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15210
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProdResultInquiry.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10560
   ScaleWidth      =   15210
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExcel 
      BackColor       =   &H0080FFFF&
      Caption         =   "E&xcel"
      Height          =   375
      Left            =   9345
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9915
      Width           =   1410
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Worksheet"
      Height          =   375
      Index           =   1
      Left            =   12270
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   9915
      Width           =   1260
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H0080FFFF&
      Caption         =   "Sea&rch"
      Height          =   405
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2580
      Width           =   1065
   End
   Begin VB.CommandButton cmdReport 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Print"
      Height          =   375
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9915
      Width           =   1410
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "&To Result Input"
      Height          =   375
      Index           =   0
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9915
      Width           =   1530
   End
   Begin VB.CommandButton cmdSubMenu 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Sub &Menu"
      Height          =   375
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9915
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   150
      TabIndex        =   22
      Top             =   9210
      Width           =   14925
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
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   195
         Width           =   14685
      End
   End
   Begin VB.ComboBox cboRemaining 
      Height          =   315
      ItemData        =   "frmProdResultInquiry.frx":0E42
      Left            =   7320
      List            =   "frmProdResultInquiry.frx":0E4C
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2640
      Width           =   885
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1215
      Left            =   150
      TabIndex        =   14
      Top             =   1200
      Width           =   14925
      Begin VB.Label lblNm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   0
         Left            =   2940
         TabIndex        =   18
         Top             =   330
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Machine No :"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Top             =   780
         Width           =   1110
      End
      Begin MSForms.ComboBox cbo 
         Height          =   315
         Index           =   0
         Left            =   1500
         TabIndex        =   0
         Top             =   270
         Width           =   1335
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2355;556"
         ListRows        =   15
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         Value           =   "AAAAAA"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cbo 
         Height          =   315
         Index           =   1
         Left            =   1500
         TabIndex        =   1
         Top             =   720
         Width           =   1335
         VariousPropertyBits=   746604571
         MaxLength       =   3
         DisplayStyle    =   3
         Size            =   "2355;556"
         ListRows        =   15
         ShowDropButtonWhen=   2
         Value           =   "AAA"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblNm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   1
         Left            =   2940
         TabIndex        =   16
         Top             =   750
         Width           =   960
      End
      Begin VB.Line Line8 
         Index           =   1
         X1              =   2940
         X2              =   4440
         Y1              =   1020
         Y2              =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Factory CD :"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   15
         Top             =   330
         Width           =   1095
      End
      Begin VB.Line Line8 
         Index           =   0
         X1              =   2940
         X2              =   8010
         Y1              =   540
         Y2              =   540
      End
   End
   Begin MSComCtl2.DTPicker dtAwal 
      Height          =   330
      Left            =   1620
      TabIndex        =   2
      Top             =   2625
      Width           =   1785
      _ExtentX        =   3149
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
      CustomFormat    =   "dd MMM yyyy"
      Format          =   141230083
      CurrentDate     =   37860
   End
   Begin MSComCtl2.DTPicker dtAkhir 
      Height          =   330
      Left            =   3810
      TabIndex        =   3
      Top             =   2625
      Width           =   1785
      _ExtentX        =   3149
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
      CustomFormat    =   "dd MMM yyyy"
      Format          =   141230083
      CurrentDate     =   37891
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   5985
      Left            =   150
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3165
      Width           =   14925
      _cx             =   26326
      _cy             =   10557
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
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   1
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
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   13230
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   405
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   714
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remaining Cls"
      Height          =   195
      Index           =   4
      Left            =   5940
      TabIndex        =   21
      Top             =   2700
      Width           =   1230
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "to"
      Height          =   195
      Index           =   3
      Left            =   3510
      TabIndex        =   20
      Top             =   2700
      Width           =   165
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Schedule Date :"
      Height          =   195
      Index           =   2
      Left            =   150
      TabIndex        =   19
      Top             =   2700
      Width           =   1380
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Production Schedule/Result Inquiry"
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
      Left            =   5595
      TabIndex        =   13
      Top             =   405
      Width           =   4035
   End
End
Attribute VB_Name = "frmProdResultInquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i As Long
Dim nilKosong As Boolean

Public fromProd As Boolean

Dim bteColSelect As Byte
Dim bteColDate As Byte
Dim bteColPart As Byte
Dim bteColProdCode As Byte
Dim bteColDesc As Byte
Dim bteColLotNo As Byte
Dim bteColPackCode As Byte
Dim bteColPackDesc As Byte
Dim bteColPackSize As Byte
Dim BteColSerialFrom As Byte
Dim BteColSerialTo As Byte
Dim bteColPlan As Byte
Dim bteColResult As Byte
Dim bteColRemain As Byte
Dim bteColComplete As Byte
Dim bteColWHCode As Byte
Dim bteColSeqNo As Byte
Dim bteColUnitCls As Byte
Dim bteColAuto As Byte
Dim bteColCustCode As Byte
Dim bteColCustName As Byte
Dim bteColPONo As Byte

Private Sub headerGrid()
    
    bteColSelect = 0
    bteColDate = 1
    bteColPart = 2
    bteColProdCode = 3
    bteColDesc = 4
    bteColLotNo = 5
    bteColPackCode = 6
    bteColPackDesc = 7
    bteColPackSize = 8
    BteColSerialFrom = 9
    BteColSerialTo = 10
    bteColPlan = 9 + 2
    bteColResult = 10 + 2
    bteColRemain = 11 + 2
    bteColComplete = 12 + 2
    bteColWHCode = 13 + 2
    bteColSeqNo = 14 + 2
    bteColUnitCls = 15 + 2
    bteColAuto = 16 + 2
    bteColCustCode = 17 + 2
    bteColCustName = 18 + 2
    bteColPONo = 19 + 2
    
    With grid
        .clear
        .ColS = 20 + 2
        .Rows = 1
        
        .TextMatrix(0, bteColSelect) = ""
        .TextMatrix(0, bteColDate) = "Schedule Date"
        .TextMatrix(0, bteColPart) = "Part Number"
        .TextMatrix(0, bteColProdCode) = "Product Code"
        .TextMatrix(0, bteColProdCode) = "Product Code"
        .TextMatrix(0, bteColDesc) = "Description"
        .TextMatrix(0, bteColLotNo) = "Lot No"
        .TextMatrix(0, bteColPackCode) = "Packing Code"
        .TextMatrix(0, bteColPackDesc) = "Description"
        .TextMatrix(0, bteColPackSize) = "Packing Size"
        .TextMatrix(0, BteColSerialFrom) = "Serial From"    ' Add 20090210
        .TextMatrix(0, BteColSerialTo) = "Serial To"    ' Add 20090210
        .TextMatrix(0, bteColPlan) = "Plan"
        .TextMatrix(0, bteColResult) = "Result"
        .TextMatrix(0, bteColRemain) = "Remaining"
        .TextMatrix(0, bteColComplete) = "Complete"
        .TextMatrix(0, bteColWHCode) = "WHCode"
        .TextMatrix(0, bteColSeqNo) = "SeqNo"
        .TextMatrix(0, bteColUnitCls) = "UnitCls"
        .TextMatrix(0, bteColAuto) = "Auto"
        .TextMatrix(0, bteColCustCode) = "Cust. Code"
        .TextMatrix(0, bteColCustName) = "Cust. Name"
        .TextMatrix(0, bteColPONo) = "PO. No."
        
        .ColWidth(bteColSelect) = 275
        .ColWidth(bteColDate) = 1400
        .ColWidth(bteColPart) = 1700
        .ColWidth(bteColProdCode) = 1700
        .ColWidth(bteColDesc) = 3000
        .ColWidth(bteColLotNo) = 1000
        .ColWidth(BteColSerialFrom) = 1100      ' Add 20090210
        .ColWidth(BteColSerialTo) = 1100         ' Add 20090210
        .ColWidth(bteColPlan) = 1500
        .ColWidth(bteColResult) = 1500
        .ColWidth(bteColRemain) = 1500
        .ColWidth(bteColComplete) = 1000
        .ColWidth(bteColAuto) = 1000
        .ColWidth(bteColCustCode) = 1200
        .ColWidth(bteColCustName) = 2800
        .ColWidth(bteColPONo) = 2000
        
        .ColHidden(bteColPackCode) = True
        .ColHidden(bteColPackDesc) = True
        .ColHidden(bteColPackSize) = True
        .ColHidden(bteColWHCode) = True
        .ColHidden(bteColSeqNo) = True
        .ColHidden(bteColUnitCls) = True
        .ColHidden(bteColCustCode) = True
        .ColHidden(bteColCustName) = True
        .ColHidden(bteColPONo) = True
        
        .ColAlignment(bteColSelect) = flexAlignCenterCenter
        .ColAlignment(bteColDate) = flexAlignLeftCenter
        .ColAlignment(bteColPart) = flexAlignLeftCenter
        .ColAlignment(bteColProdCode) = flexAlignLeftCenter
        .ColAlignment(bteColDesc) = flexAlignLeftCenter
        .ColAlignment(bteColLotNo) = flexAlignLeftCenter
        .ColAlignment(bteColPackCode) = flexAlignLeftCenter
        .ColAlignment(bteColPackDesc) = flexAlignLeftCenter
        .ColAlignment(bteColPackSize) = flexAlignRightCenter
        
        .ColAlignment(BteColSerialFrom) = flexAlignCenterCenter
        .ColAlignment(BteColSerialTo) = flexAlignCenterCenter
        
        .ColAlignment(bteColPlan) = flexAlignRightCenter
        .ColAlignment(bteColResult) = flexAlignRightCenter
        .ColAlignment(bteColRemain) = flexAlignRightCenter
        .ColAlignment(bteColComplete) = flexAlignCenterCenter
        .ColAlignment(bteColAuto) = flexAlignCenterCenter
        .ColAlignment(bteColCustCode) = flexAlignLeftCenter
        .ColAlignment(bteColCustName) = flexAlignLeftCenter
        .ColAlignment(bteColPONo) = flexAlignLeftCenter
                
        .EditMaxLength = 1
    End With
End Sub


Sub viewWorksheet()
Dim application As New CRAXDDRT.application
Dim report As New CRAXDDRT.report
Dim rsRpt As New ADODB.Recordset
Dim Rpt As New FrmRpt3
Dim sqlResult As String

Dim lngCount As Long
Dim strCheck As String
Dim booCheck As Boolean

    Me.MousePointer = vbHourglass
    
    If cbo(0) = "" Then
        LblErrMsg = DisplayMsg(1040)
        cbo(0).SetFocus
    Else
        cbo(0) = cbo(0)
        If cbo(0).MatchFound = False Then
            LblErrMsg = DisplayMsg(4016)
            cbo(0).SetFocus
        Else
        
            strCheck = "": booCheck = False
            For lngCount = 1 To grid.Rows - 1
                If grid.Cell(flexcpChecked, lngCount, 0) = flexChecked Then
                    If Not booCheck Then booCheck = True
                    If strCheck <> "" Then
                        strCheck = strCheck & ", "
                    End If
                    strCheck = strCheck & grid.TextMatrix(lngCount, bteColSeqNo)
                End If
            Next
            
            If Not booCheck Then
                LblErrMsg = DisplayMsg(8011)
                Me.MousePointer = vbDefault
                Exit Sub
            End If
        
            LblErrMsg = ""
            '**** Utk Report
            
            sql = " select * from " & _
                    "   (select    rtrim(a.factory_code) as factory_code, rtrim(a.line_code) as line_code, a.schedule_date,  " & _
                    " a.SerialNoFrom PSerialFrom,a.SerialNoTo PSerialTo, " & _
                    "       rtrim(a.item_code) as item_code, rtrim(a.lot_no) as lot_no, a.seq_no, a.qty, " & _
                    "       QtyResult = ISNULL((Select Sum(Qty) from Part_Receipt where DailySeq_No = a.Seq_No And ProductionResult_Cls = 1 and receipt_cls='p1'),0), " & _
                    "       rtrim(a.unit_cls) as unit_cls,   (select description from unit_cls uc where uc.unit_cls=a.unit_cls) unit_desc,a.complete_cls, " & _
                    "       a.remark, b.item_name, (select trade_name from trade_master where trade_code = '" & cbo(0) & "') factory_name,  " & _
                    "       (select line_name from manufacture_line where manufacture_code = '" & cbo(0) & "' and line_code = '" & cbo(1) & "') line_name, " & _
                    "       b.number_entering, rtrim(a.prod_barcode) prod_barcode, " & _
                    "       (select rtrim(production_person) from company_profile) Prod_Person, " & _
                    "       (select rtrim(production_position) from company_profile) production_Position,  " & _
                    "       (select rtrim(QC_person) from company_profile) QC_Person, " & _
                    "       (select rtrim(QC_position) from company_profile) QC_Position,   " & _
                    "       (select rtrim(PPC_person) from company_profile) PPC_Person,  " & _
                    "       (select rtrim(PPC_position) from company_profile) PPC_Position  " & _
                    "       from daily_production a  " & _
                    "       left join item_master b on b.item_code=a.item_code  " & _
                    "   where a.factory_code='" & cbo(0) & "' and a.line_code='" & cbo(1) & "' and a.schedule_date>='" & Format(dtAwal, "yyyy-MM-dd") & "' and a.schedule_date<='" & Format(dtAkhir, "yyyy-MM-dd") & "'  " & _
                    "   and a.seq_no in (" & strCheck & ")) x where "
                    
            If cboRemaining = "Yes" Then
                sql = sql & "Qty - QtyResult > 0 And (Complete_Cls is Null Or Complete_Cls = 0) "
            Else
                sql = sql & "(Qty - QtyResult <= 0 Or Complete_Cls =1) "
            End If
                    
            sql = sql & " order by schedule_date, item_code, lot_no, seq_no "
            Set rsRpt = Db.Execute(sql)
            
            If rsRpt.EOF Then
                LblErrMsg.Caption = DisplayMsg(4006)
            Else
                sqlprint = sql
                reportcode = "rptWorksheet"
                printorient = 1
                tglAwalRptPrint = Format(dtAwal, "dd MMM yyyy")
                tglAkhirRptPrint = Format(dtAkhir, "dd MMM yyyy")
                
                Set report = application.OpenReport(App.path & "\Reports\rptWorksheet.rpt")
                report.Database.Tables(1).SetDataSource rsRpt
                
                '#####################################################################
                '# Qty Digit and decimal
                report.FormulaFields(2).Text = "" & gi_decimalDigitQty & ""
                report.FormulaFields(3).Text = "" & gi_decimalDigitQty & ""
                report.FormulaFields(4).Text = "" & gi_decimalDigitBox & ""
                report.FormulaFields(5).Text = "" & gi_decimalDigitBox & ""
                '#####################################################################
                            
                Rpt.CRViewer1.ReportSource = report
                Rpt.CRViewer1.ViewReport
                Rpt.CRViewer1.Zoom 1
                
                Rpt.WindowState = 2
                Rpt.Show 1
            End If
            Set rsRpt = Nothing
        End If
    End If
    
    Me.MousePointer = vbDefault
    
End Sub

Sub previewReport()
Dim application As New CRAXDDRT.application
Dim report As New CRAXDDRT.report
Dim rsRpt As New ADODB.Recordset
Dim Rpt As New FrmRpt3
Dim sqlResult As String

    Me.MousePointer = vbHourglass
    
    If cbo(0) = "" Then
        LblErrMsg = DisplayMsg(1040)
        cbo(0).SetFocus
    Else
        cbo(0) = cbo(0)
        If cbo(0).MatchFound = False Then
            LblErrMsg = DisplayMsg(4016)
            cbo(0).SetFocus
        Else
            LblErrMsg = ""
            '**** Utk Report
    
            
            sql = "Select z.* from("
            
            
            sql = sql & _
                  vbLf & " select rtrim(a.factory_code) as factory_code, rtrim(a.line_code) as line_code, a.schedule_date,  " & _
                  vbLf & " a.SerialNoFrom PSerialFrom,a.SerialNoTo PSerialTo, " & _
                  vbLf & "   rtrim(b.makeritem_code) makeritem_code, rtrim(a.item_code) as item_code, rtrim(a.lot_no) as lot_no, a.seq_no, a.qty, rtrim(a.unit_cls) as unit_cls, (select description from unit_cls uc where uc.unit_cls=a.unit_cls)  unit_desc, " & _
                  vbLf & "   a.remark, b.item_name, (select trade_name from trade_master where trade_code = '" & cbo(0) & "') factory_name,  " & _
                  vbLf & "   (select line_name from manufacture_line where manufacture_code = '" & cbo(0) & "' and line_code = '" & cbo(1) & "') line_name,  " & _
                  vbLf & "   c.receipt_date, b.number_entering,  a.complete_cls, c.SerialNoFrom RSerialFrom,C.SerialNoTo RSerialTo,isnull(c.Qty,0) QtyResult," & _
                  vbLf & "   (select rtrim(production_person) from company_profile) Prod_Person, " & _
                  vbLf & "   (select rtrim(production_position) from company_profile) production_Position,  " & _
                  vbLf & "   (select rtrim(QC_person) from company_profile) QC_Person, " & _
                  vbLf & "   (select rtrim(QC_position) from company_profile) QC_Position,   " & _
                  vbLf & "   (select rtrim(PPC_person) from company_profile) PPC_Person,  " & _
                  vbLf & "   (select rtrim(PPC_position) from company_profile) PPC_Position,  " & _
                  vbLf & "   (select isnull(sum(qty),0) from part_receipt where dailyseq_no = a.seq_no and receipt_cls = 'P1' and receipt_date= c.receipt_date) QtyResult1 " & _
                  vbLf & " from daily_production a  " & _
                  vbLf & "   left join item_master b on b.item_code=a.item_code  " & _
                  vbLf & "   left join part_receipt c on (c.dailyseq_no = a.seq_no and c.receipt_cls = 'P1') " & _
                  vbLf & " where a.factory_code='" & cbo(0) & "' and a.line_code='" & cbo(1) & "' and a.schedule_date>='" & Format(dtAwal, "yyyy-MM-dd") & "' and a.schedule_date<='" & Format(dtAkhir, "yyyy-MM-dd") & "' "
  
            sql = sql & _
                ") z "
            
            If cboRemaining = "Yes" Then
                sql = sql & vbLf & " where z.Qty - z.QtyResult > 0 And (z.Complete_Cls is Null Or z.Complete_Cls = 0) "
            Else
                sql = sql & vbLf & " where (z.Qty - z.QtyResult <= 0 Or z.Complete_Cls =1) "
            End If
            
            sql = sql & _
                  vbLf & " order by z.schedule_date, z.item_code, z.lot_no, z.seq_no "
                  
            Set rsRpt = Db.Execute(sql)
            
            If rsRpt.EOF Then
                LblErrMsg.Caption = DisplayMsg(4006)
            Else
                sqlprint = sql
                printorient = 2
                reportcode = "rptProdResultInquiry"
                tglAwalRptPrint = Format(dtAwal, "dd MMM yyyy")
                tglAkhirRptPrint = Format(dtAkhir, "dd MMM yyyy")
                
                Set report = application.OpenReport(App.path & "\Reports\rptProdResultInquiry.rpt")
                report.Database.Tables(1).SetDataSource rsRpt
                report.FormulaFields(1).Text = "'" & Format(dtAwal, "dd MMM yyyy") & "'"
                report.FormulaFields(2).Text = "'" & Format(dtAkhir, "dd MMM yyyy") & "'"
                            
                '#####################################################################
                '# Qty Digit and decimal
                report.FormulaFields(4).Text = "" & gi_decimalDigitQty & ""
                report.FormulaFields(5).Text = "" & gi_decimalDigitQty & ""
                report.FormulaFields(6).Text = "" & gi_decimalDigitBox & ""
                report.FormulaFields(7).Text = "" & gi_decimalDigitBox & ""
                '#####################################################################
                
                Rpt.CRViewer1.ReportSource = report
                Rpt.CRViewer1.ViewReport
                Rpt.CRViewer1.Zoom 1
                
                Rpt.WindowState = 2
                Rpt.Show 1
            End If
            Set rsRpt = Nothing
        End If
    End If
    
    Me.MousePointer = vbDefault
    
End Sub

Sub Kosong()
    nilKosong = True
    cbo(0) = ""
    lblNm(0) = ""
    cbo(1) = ""
    lblNm(1) = ""
    dtAwal = Format(Year(Now) & "-" & Format(Month(Now), "#0") & "-01", "dd MMM yyyy")
    dtAkhir = Format(Now, "dd MMM yyyy")
    cboRemaining.ListIndex = 0
    Call headerGrid
    nilKosong = False
End Sub

'******** Combo Factory Code **********
Sub isiCboCust() 'Isi Combo Dealer CD dr Customer Master
Dim RsCust As New ADODB.Recordset 'Data Customer

With cbo(0)
    .clear
    .columnCount = 2
    .TextColumn = 1
    
    '******** Ambil dr Customer Master utk Combo Dealer CD
    sql = "select Trade_Code,Trade_Name from Trade_Master " & _
        "where trade_code in (select distinct manufacture_code from manufacture_line) " & _
        "order by Trade_Code"
    Set RsCust = Db.Execute(sql)
    
    i = 0
    Do While Not (RsCust.EOF)
        .AddItem ""
        .List(i, 0) = Trim(RsCust(0))
        .List(i, 1) = Trim(RsCust(1))
        i = i + 1
        RsCust.MoveNext
    Loop
    
    .Text = ""
    .ListWidth = 350
    .ColumnWidths = "50 pt;300 pt"
    
    Set RsCust = Nothing
End With
End Sub

'******** Filter Combo Line Code **********
Sub isiCboLine(factoryCD As String)
Dim rscbo As New ADODB.Recordset

With cbo(1)
    .clear
    .columnCount = 2
    .TextColumn = 1
    
    sql = "select Line_Code,Line_Name from Manufacture_line " & _
        "where Manufacture_Code = '" & factoryCD & _
        "' order by Line_Code"
    Set rscbo = Db.Execute(sql)
     
    i = 0
    Do While Not (rscbo.EOF)
        .AddItem ""
        .List(i, 0) = Trim(rscbo(0))
        .List(i, 1) = Trim(rscbo(1))
        i = i + 1
        rscbo.MoveNext
    Loop
    
    .Text = ""
    .ListWidth = 200
    .ColumnWidths = "50 pt;150 pt"
    
    Set rscbo = Nothing
End With
End Sub

Private Sub CmdExcel_Click()
    Dim xlapp As New Excel.application
    Dim rsCek As New Recordset, Idx As Long, tempi As String, tempcls As String
    Dim bolcls As Boolean, bolcur As Boolean
    Dim tempSeq As String, tempResult As Double
    Dim rsCompany As New Recordset
    
    MousePointer = vbHourglass
    
            sql = "Select z.* from("
            
            sql = sql & _
                  vbLf & " select rtrim(a.factory_code) as factory_code, rtrim(a.line_code) as line_code, a.schedule_date,  " & _
                  vbLf & " a.SerialNoFrom PSerialFrom,a.SerialNoTo PSerialTo, " & _
                  vbLf & "   rtrim(b.makeritem_code) makeritem_code, rtrim(a.item_code) as item_code, rtrim(a.lot_no) as lot_no, a.seq_no, a.qty, rtrim(a.unit_cls) as unit_cls, (select description from unit_cls uc where uc.unit_cls=a.unit_cls)  unit_desc, " & _
                  vbLf & "   a.remark, b.item_name, (select trade_name from trade_master where trade_code = '" & cbo(0) & "') factory_name,  " & _
                  vbLf & "   (select line_name from manufacture_line where manufacture_code = '" & cbo(0) & "' and line_code = '" & cbo(1) & "') line_name,  " & _
                  vbLf & "   c.receipt_date, b.number_entering,  a.complete_cls, c.SerialNoFrom RSerialFrom,C.SerialNoTo RSerialTo, isnull(c.Qty,0) Qtyresult," & _
                  vbLf & "   (select rtrim(production_person) from company_profile) Prod_Person, " & _
                  vbLf & "   (select rtrim(production_position) from company_profile) production_Position,  " & _
                  vbLf & "   (select rtrim(QC_person) from company_profile) QC_Person, " & _
                  vbLf & "   (select rtrim(QC_position) from company_profile) QC_Position,   " & _
                  vbLf & "   (select rtrim(PPC_person) from company_profile) PPC_Person,  " & _
                  vbLf & "   (select rtrim(PPC_position) from company_profile) PPC_Position,  " & _
                  vbLf & "   (select isnull(sum(qty),0) from part_receipt where dailyseq_no = a.seq_no and receipt_cls = 'P1' and receipt_date= c.receipt_date) QtyResult1 " & _
                  vbLf & " from daily_production a  " & _
                  vbLf & "   left join item_master b on b.item_code=a.item_code  " & _
                  vbLf & "   left join part_receipt c on (c.dailyseq_no = a.seq_no and c.receipt_cls = 'P1') " & _
                  vbLf & " where a.factory_code='" & cbo(0) & "' and a.line_code='" & cbo(1) & "' and a.schedule_date>='" & Format(dtAwal, "yyyy-MM-dd") & "' and a.schedule_date<='" & Format(dtAkhir, "yyyy-MM-dd") & "' "
                
            sql = sql & _
                ") z "
            
            If cboRemaining = "Yes" Then
                sql = sql & vbLf & "where z.Qty - z.QtyResult > 0 And (z.Complete_Cls is Null Or z.Complete_Cls = 0) "
            Else
                sql = sql & vbLf & "where (z.Qty - z.QtyResult <= 0 Or z.Complete_Cls =1) "
            End If
            
            sql = sql & _
                  vbLf & " order by z.schedule_date, z.item_code, z.lot_no, z.seq_no "
    
    
    If rsCek.State <> adStateClosed Then rsCek.Close
    
    Set rsCek = Db.Execute(sql)
    
    If rsCek.EOF Then
        LblErrMsg.Caption = DisplayMsg(4006)
    Else
            
        With xlapp
            
            .Workbooks.Add

            .Range("a2", "l2").Merge
            .Range("a2") = "Production Schedule / Result"

            .Range("a4") = "Factory Code"
            .Range("b4") = ": " & Trim(rsCek!Factory_Code)
            .Range("c4", "j4").Merge
            .Range("c4") = "Factory Name :  " & rsCek!Factory_Name
            .Range("a5") = "Machine No."
            .Range("b5") = ": " & Trim(rsCek!line_code)
            .Range("c5", "j5").Merge
            .Range("c5") = "Machine Name :  " & rsCek!line_name
            .Range("b6", "d6").Merge
            .Range("a6") = "Period"
            .Range("b6") = ": " & Format(dtAwal, "dd MMMM YYYY") & " to " & Format(dtAkhir, "dd MMMM YYYY")

            Idx = 8
            Do While Not rsCek.EOF
               
                If Idx = 8 Then
                    .Range("a" & Idx).horizontalAlignment = xlCenter
                    .Range("a" & Idx) = "Schedule Date"
                    .Range("b" & Idx) = "Product Code"
                    .Range("c" & Idx) = "Part Number"
                    .Range("d" & Idx) = "Description"
                    .Range("e" & Idx) = "Lot No."
                    ' Add 20090210
                    .Range("f" & Idx) = "Schedule"
                    .Range("f" & Idx, "h" & Idx).Merge
                    .Range("f" & Idx + 1) = "Serial From"
                    .Range("g" & Idx + 1) = "Serial To"
                    .Range("h" & Idx + 1) = "Qty"
                    '---
                    .Range("i" & Idx) = "Pcs / Case"
                    .Range("j" & Idx) = "Unit"
                    .Range("k" & Idx) = "Result Date"
                    ' Add 20090210
                    .Range("l" & Idx) = "Result"
                    .Range("l" & Idx, "n" & Idx).Merge
                    .Range("l" & Idx + 1) = "Serial From"
                    .Range("m" & Idx + 1) = "Serial To"
                    .Range("n" & Idx + 1) = "Qty"
                    '--------
                    .Range("o" & Idx) = "Total"
                    .Range("p" & Idx) = "Remark"
                    
                    .Range("a" & Idx, "p" & Idx).horizontalAlignment = xlCenter
                    .Range("a" & Idx + 1, "p" & Idx + 1).horizontalAlignment = xlCenter
                    
                    .Range("i" & Idx).horizontalAlignment = xlCenter
                    .Range("a" & Idx, "p" & Idx).Borders(xlEdgeTop).LineStyle = xlContinuous
                    .Range("a" & Idx + 1, "p" & Idx).Borders(xlEdgeBottom).LineStyle = xlContinuous
                    Idx = Idx + 2
                End If
                
                'Content
                
                .Range("a" & Idx).horizontalAlignment = xlCenter
                .Range("i" & Idx).horizontalAlignment = xlCenter
                
                If tempSeq <> rsCek!Seq_no Then
                    tempResult = 0
                    .Range("a" & Idx) = Format(rsCek!schedule_date, "DD-MMM-YYYY")
                    .Range("b" & Idx) = Trim(rsCek!Item_Code)
                    .Range("c" & Idx) = Trim(rsCek!MakerItem_Code)
                    .Range("d" & Idx) = Trim(rsCek!item_name)
                    .Range("e" & Idx) = "'" & Trim(rsCek!Lot_no)
                    ' Add 20090210
                    .Range("f" & Idx) = "'" & Trim(rsCek!PSerialFrom)
                    .Range("g" & Idx) = "'" & Trim(rsCek!PSerialTo)
                    '---
                    .Range("h" & Idx) = Format(rsCek!Qty, gs_formatQty)
                    .Range("i" & Idx) = Format(rsCek!number_entering, gs_formatBox)
                    .Range("j" & Idx) = Trim(rsCek!Unit_Desc)
                    If Not IsNull(rsCek!Receipt_Date) Then
                        .Range("k" & Idx) = Format(rsCek!Receipt_Date, "DD-MMM-YYYY")
                        ' Add 20090210
                        .Range("l" & Idx) = "'" & Trim(rsCek!RSerialFrom)
                        .Range("m" & Idx) = "'" & Trim(rsCek!RSerialTo)
                        '---
                        'Qty Result Not From Summary of Qty
                        .Range("n" & Idx) = Format(rsCek!QtyResult, gs_formatQty)
                        .Range("o" & Idx) = Format(tempResult + rsCek!QtyResult, gs_formatQty)
'                        .Range("n" & Idx) = Format(rsCek!QtyResult, gs_formatQty)
'                        .Range("o" & Idx) = Format(tempResult + rsCek!QtyResult, gs_formatQty)
                        
                        '---
                        .Range("p" & Idx) = Trim(rsCek!Remark)
                    End If
                ElseIf Not IsNull(rsCek!Receipt_Date) Then
                        .Range("k" & Idx) = Format(rsCek!Receipt_Date, "DD-MMM-YYYY")
                        ' Add 20090210
                        .Range("l" & Idx) = "'" & Trim(rsCek!RSerialFrom)
                        .Range("m" & Idx) = "'" & Trim(rsCek!RSerialTo)
                        '---
                        'Qty Result Not From Summary of Qty
                        .Range("n" & Idx) = Format(rsCek!QtyResult, gs_formatQty)
                        .Range("o" & Idx) = Format(tempResult + rsCek!QtyResult, gs_formatQty)
'                        .Range("n" & Idx) = Format(rsCek!QtyResult, gs_formatQty)
'                        .Range("o" & Idx) = Format(tempResult + rsCek!QtyResult, gs_formatQty)
                        
                        '---
                        .Range("p" & Idx) = Trim(rsCek!Remark)
                End If
                tempSeq = rsCek!Seq_no
                'tempResult = tempResult + rsCek!QtyResult
                tempResult = tempResult + rsCek!QtyResult
                
                Idx = Idx + 1
                rsCek.MoveNext
            Loop
            
            '################################################
            '#Format Qty
            .Range("h9:h" & Idx - 1).NumberFormat = gs_formatQty
            .Range("n9:o" & Idx - 1).NumberFormat = gs_formatQty
            '################################################
            
            .Range("a" & Idx, "p" & Idx).Borders(xlEdgeTop).LineStyle = xlContinuous
            Idx = Idx + 3
            
            .Range("a1", "p" & Idx + 3).Columns.Font.Name = "Arial"
            .Range("a1", "p" & Idx + 3).Columns.Font.Size = 8
            .Range("a2", "p2").Columns.Font.Name = "Arial"
            .Range("a2", "p2").Columns.Font.Size = "10"
            .Range("a2", "p2").Columns.Font.Bold = True
            .Range("a2", "p2").horizontalAlignment = xlCenter

            .ActiveSheet.PageSetup.Orientation = 2
            .WindowState = xlMaximized
            .Range("a1", "p" & Idx + 3).Columns.AutoFit
            .Visible = True
        End With
    End If
    MousePointer = vbDefault

End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
    fromProd = True
    nilKosong = True
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
    'frm_part_supply.ls_dataStatus = "invalid"
    Call isiCboCust
    Call Kosong
    nilKosong = False
End Sub

Private Sub cbo_KeyDown(Index As Integer, KeyCode As MSForms.ReturnInteger, Shift As Integer)
If nilKosong = True Then Exit Sub
    If KeyCode = 13 Then Call cbo_Click(Index)
End Sub

Private Sub cbo_Change(Index As Integer)
If nilKosong = True Then Exit Sub
    lblNm(Index) = ""
    'Hapus Manufacture Line * Desc
    If Index = 0 Then cbo(1).clear: lblNm(1) = "": Call headerGrid
End Sub

Private Sub cbo_LostFocus(Index As Integer)
If nilKosong = True Then Exit Sub
    If lblNm(Index) = "" Then Call cbo_Click(Index)
End Sub

'*********** Tampilkan Data *********
Private Sub cbo_Click(Index As Integer)
If nilKosong = True Then Exit Sub

If cbo(Index) <> "" Then
    cbo(Index) = cbo(Index)
    If cbo(Index).MatchFound = True Then
        lblNm(Index) = cbo(Index).Column(1)
        If Index = 0 Then 'panggil Manufacture Line
            Call isiCboLine(cbo(0)): lblNm(1) = ""
        End If
        LblErrMsg = ""
    Else
        lblNm(Index) = ""
        If Index = 0 Then 'Hapus Manufacture Line & Desc Line
            cbo(1).clear: lblNm(1) = ""
        End If
        LblErrMsg = DisplayMsg(4016 + Index) 'Err Msg en Panggil Grid
    End If
Else
    lblNm(Index) = ""
    If Index = 0 Then 'Hapus Manufacture Line * Desc
        cbo(1).clear: lblNm(1) = ""
    End If
    LblErrMsg = ""
End If
End Sub

Public Sub cmdSearch_Click()
    cbo(0) = cbo(0)
    cbo(1) = cbo(1)
    If cbo(0) = "" Then
        LblErrMsg = DisplayMsg(1040)
        cbo(0).SetFocus
    ElseIf cbo(0).MatchFound = False Then
        LblErrMsg = DisplayMsg(4016)
        cbo(0).SetFocus
    ElseIf cbo(1).MatchFound = False And cbo(1) <> "" Then
        LblErrMsg = DisplayMsg(4017)
        cbo(1).SetFocus
    Else
        LblErrMsg = ""
        Call IsiGrid
    End If
    
End Sub

Sub IsiGrid()
Dim rsProd As New ADODB.Recordset
Dim sqlResult As String

Me.MousePointer = vbHourglass

If nilKosong = True Then Exit Sub
With grid
    Call headerGrid
    
    sql = "select a.Seq_No,Schedule_Date,a.Item_Code,a.SerialNoFrom,a.SerialNoTo, " & _
            " rtrim(b.makeritem_code) makeritem_code, rtrim(b.Item_Name) Descr, Lot_No," & _
            " qty "
    
    sqlResult = "(select Isnull(Sum(Qty),0) from " & _
            "Part_Receipt where DailySeq_No = a.Seq_No " & _
            "And ProductionResult_Cls = 1 and receipt_cls='p1') "
    
    sql = sql & "," & sqlResult & " as Result " & _
        ",(Qty - " & sqlResult & ") as Sisa, Isnull(a.Complete_Cls,0) Complete_Cls," & _
        "WH_Code,b.Unit_Cls, a.auto_cls, pcd.cust_code, pcd.po_no, tm.trade_name " & _
        "from Daily_Production a " & _
        "left join Item_Master b on a.Item_Code = b.Item_Code " & _
        "left join productioncalculate_detail pcd on a.plancust_code = pcd.cust_code and a.planpo_seqno = pcd.seq_no and a.plan_seqno = pcd.plan_seqno and a.item_code = pcd.planitem_code and a.planpo_no = pcd.po_no  " & _
        "left join trade_master tm on pcd.cust_code = tm.trade_code " & _
        "Where Schedule_Date >= '" & Format(dtAwal, "yyyy-MM-dd") & _
        "' and Schedule_Date <= '" & Format(dtAkhir, "yyyy-MM-dd") & _
        "' and Factory_Code = '" & cbo(0) & _
        "' And a.line_Code = '" & cbo(1) & "'"

    If cboRemaining = "Yes" Then
        sql = sql & " And Qty - " & sqlResult & " > 0 " & _
            "And (Complete_Cls is Null Or Complete_Cls = 0) "
    Else
        sql = sql & " And (Qty - " & sqlResult & " <= 0 " & _
            "Or Complete_Cls =1)"
    End If
    
    sql = sql & "order by Schedule_Date,a.Item_Code,Lot_no "
    Set rsProd = Db.Execute(sql)
    
    i = 1
    If Not (rsProd.EOF) Then
        Do While Not rsProd.EOF
            .Rows = .Rows + 1
            .Cell(flexcpChecked, i, bteColSelect) = flexUnchecked
            .Cell(flexcpBackColor, i, bteColSelect) = vbWhite
            .TextMatrix(i, bteColDate) = Format(Trim(rsProd("Schedule_Date")), "dd MMM yyyy")
            .TextMatrix(i, bteColPart) = Trim(rsProd("MakerItem_Code"))
            .TextMatrix(i, bteColProdCode) = Trim(rsProd("Item_Code"))
            .TextMatrix(i, bteColDesc) = Trim(rsProd("Descr"))
            .TextMatrix(i, bteColLotNo) = Trim(rsProd("Lot_no"))
        ' Add 20090210
            .TextMatrix(i, BteColSerialFrom) = IIf(IsNull(Trim(rsProd("SerialNoFrom"))), "", Trim(rsProd("SerialNoFrom")))
            .TextMatrix(i, BteColSerialTo) = IIf(IsNull(Trim(rsProd("SerialNoTo"))), "", Trim(rsProd("SerialNoTo")))
        ' -----------
            .TextMatrix(i, bteColPlan) = Format(rsProd("Qty"), gs_formatQty)
            .TextMatrix(i, bteColResult) = Format(rsProd("Result"), gs_formatQty)
            .TextMatrix(i, bteColRemain) = Format(rsProd("Sisa"), gs_formatQty)
            .Cell(flexcpChecked, i, bteColComplete) = IIf(rsProd("Complete_Cls") = 1, flexChecked, flexUnchecked)
            .TextMatrix(i, bteColWHCode) = Trim(rsProd("WH_Code"))
            .TextMatrix(i, bteColSeqNo) = rsProd("Seq_No")
            .TextMatrix(i, bteColUnitCls) = rsProd("Unit_Cls")
        
            If Val(rsProd("auto_cls") & "") = 0 Then .TextMatrix(i, bteColAuto) = "No" Else .TextMatrix(i, bteColAuto) = "Yes"
            
            .TextMatrix(i, bteColCustCode) = rsProd("cust_code") & ""
            .TextMatrix(i, bteColCustName) = rsProd("trade_name") & ""
            .TextMatrix(i, bteColPONo) = rsProd("po_no") & ""
            
            i = i + 1
            rsProd.MoveNext
        Loop
    Else
        LblErrMsg = DisplayMsg(4006)
    End If
    Set rsProd = Nothing
    Me.MousePointer = vbDefault

End With
End Sub



Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col <> bteColSelect Then Cancel = 1
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim tampung As Long
With grid
    If Row <> 0 And Col = bteColSelect Then
        If .Cell(flexcpChecked, Row, bteColSelect) = flexChecked Then
            tampung = Row
'            For i = 1 To .Rows - 1
'                .Cell(flexcpChecked, i, bteColSelect) = flexUnchecked
'            Next i
            .Cell(flexcpChecked, tampung, bteColSelect) = flexChecked
        End If
    End If
End With
End Sub

Private Sub Command1_Click(Index As Integer)
Dim cek As Integer
Dim rsCek As New ADODB.Recordset

    If cbo(0) = "" And Index <> 1 Then
        LblErrMsg = DisplayMsg(1040)
        cbo(0).SetFocus
    Else
        cbo(0) = cbo(0)
        cbo(1) = cbo(1)
        If cbo(0).MatchFound = False And Index <> 1 Then
            LblErrMsg = DisplayMsg(4016)
            cbo(0).SetFocus
        ElseIf cbo(1) <> "" And cbo(1).MatchFound = False And Index <> 1 Then
            LblErrMsg = DisplayMsg(4017)
            cbo(1).SetFocus
        ElseIf Index = 1 Then
            'WORKSHEET
            viewWorksheet
        Else
            With grid
                cek = 0
                For i = 1 To .Rows - 1
                    If .Cell(flexcpChecked, i, bteColSelect) = flexChecked Then
                        cek = i
                        Exit For
                    End If
                Next i
                
                If cek = 0 Then
                    '#If from supply request then close form Prod
                    'If Index = 1 Then GoTo frmMaterialRequest Else
                    LblErrMsg = DisplayMsg(8011)
                Else
                    DoEvents
                    Select Case Index
                        Case 0: 'Prod Result
                            If hakAkses("frmProdResult") = 0 Then LblErrMsg = DisplayMsg(3007): Exit Sub

                            If .Cell(flexcpChecked, cek, bteColWHCode) = flexChecked Then
                                LblErrMsg = DisplayMsg(1113): Exit Sub 'Daily Already Completed
                            End If
                            frmProdResult.is_LoadByItemCode = Trim(.TextMatrix(cek, bteColProdCode))
                            frmProdResult.cbo(0) = cbo(0) 'Factory CD
                            frmProdResult.cbo(1) = cbo(1) 'Line Code
                            frmProdResult.cbo(2) = Trim(.TextMatrix(cek, bteColWHCode)) ' WH Code 20090212 - Default WH Code = Factory Code
                            frmProdResult.cboResultCls.ListIndex = 0
                            frmProdResult.tglProd = Format(.TextMatrix(cek, bteColDate), "yyyy-MM-dd") 'Tgl Prod
                            frmProdResult.cbo(3) = Trim(.TextMatrix(cek, bteColProdCode)) 'Item Code
                            frmProdResult.txtLot = Trim(.TextMatrix(cek, bteColLotNo)) 'Lot NO
                            frmProdResult.txtQty = Format(0, gs_formatQty)
                            frmProdResult.txtremarks = ""
                            frmProdResult.dailyseqno = .TextMatrix(cek, bteColSeqNo) 'Daily Seq No
                            frmProdResult.qtyDaily = .TextMatrix(cek, bteColPlan) 'Daily Qty
                            frmProdResult.txtUnit = uf_GetUnitDescription(Trim(.TextMatrix(cek, bteColUnitCls)))
                            frmProdResult.UnitCls = .TextMatrix(cek, bteColUnitCls) 'Unit Cls
                            frmProdResult.qtyAllResult = .TextMatrix(cek, bteColResult) 'QTy Result
                            frmProdResult.completeCls = IIf(.Cell(flexcpChecked, cek, bteColComplete) = flexChecked, 1, 0)
                            frmProdResult.QtyRemaining = .TextMatrix(cek, bteColRemain) ' Remaining
                            
                            Call frmProdResult.tampilData
                            
                            frmProdResult.cbo(0).locked = True
                            frmProdResult.cbo(1).locked = True
                            frmProdResult.cmdsubmenu.Caption = "&Back"
                            frmProdResult.Show
                            Me.Hide
                    End Select
                    DoEvents
                End If
            End With
        End If
    End If
End Sub

Private Sub cmdReport_Click()
    previewReport
End Sub


'************ Unload **********
Private Sub CmdSubMenu_Click()
    If cmdsubmenu.Caption = "&Back" Then
        Call Command1_Click(1)
    Else
        Unload frmProdResult
        DoEvents
        frmMainMenu.Show
        DoEvents
        Unload Me
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    LblErrMsg.Caption = ErrMsg
End If
End Sub
'**************

