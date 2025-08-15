VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmFinishGoodWIPManufactureSummary 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Finish Good / WIP Manufacture (Summary)"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10065
   Icon            =   "FrmFinishGoodWIPManufactureSummary.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   10065
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H0080FFFF&
      Caption         =   "Process &Batch"
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
      Left            =   6915
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5700
      Width           =   1470
   End
   Begin VB.PictureBox PicProg 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   360
      ScaleHeight     =   495
      ScaleWidth      =   9465
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3810
      Width           =   9495
      Begin MSComctlLib.ProgressBar Prg1 
         Height          =   315
         Left            =   105
         TabIndex        =   18
         Top             =   90
         Width           =   9195
         _ExtentX        =   16219
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5670
      Width           =   1230
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   360
      TabIndex        =   15
      Top             =   4845
      Width           =   9450
      Begin VB.Label LblErrMsg 
         Alignment       =   2  'Center
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
         Height          =   270
         Left            =   150
         TabIndex        =   16
         Top             =   210
         Width           =   9210
      End
   End
   Begin VB.CommandButton cmdAction 
      BackColor       =   &H0080FFFF&
      Caption         =   "E&xcel"
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
      Left            =   8595
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5700
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1770
      Left            =   360
      TabIndex        =   8
      Top             =   1575
      Width           =   9525
      Begin VB.CommandButton cmdBrowser 
         Caption         =   "..."
         Height          =   300
         Left            =   4080
         TabIndex        =   22
         Top             =   787
         Width           =   300
      End
      Begin MSComCtl2.DTPicker dtawal 
         Height          =   315
         Left            =   1290
         TabIndex        =   2
         Top             =   1200
         Width           =   1530
         _ExtentX        =   2699
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
         Format          =   334692355
         CurrentDate     =   39173
      End
      Begin VB.Label Label1 
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
         Left            =   195
         TabIndex        =   14
         Top             =   840
         Width           =   915
      End
      Begin VB.Label lblitem 
         BackStyle       =   0  'Transparent
         Caption         =   "lblitem"
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
         Left            =   4515
         TabIndex        =   13
         Top             =   840
         Width           =   2070
      End
      Begin VB.Label lblgroup 
         BackStyle       =   0  'Transparent
         Caption         =   "lblgroup"
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
         Left            =   4560
         TabIndex        =   12
         Top             =   405
         Width           =   3810
      End
      Begin VB.Line Line2 
         X1              =   4515
         X2              =   6600
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         Left            =   195
         TabIndex        =   11
         Top             =   1260
         Width           =   405
      End
      Begin MSForms.ComboBox cboitem 
         Height          =   315
         Left            =   1290
         TabIndex        =   1
         Top             =   780
         Width           =   2685
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "4736;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cbogroup 
         Height          =   315
         Left            =   1290
         TabIndex        =   0
         Top             =   345
         Width           =   2010
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "3545;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Group"
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
         Left            =   195
         TabIndex        =   10
         Top             =   405
         Width           =   525
      End
      Begin VB.Line Line1 
         X1              =   4530
         X2              =   8460
         Y1              =   645
         Y2              =   645
      End
      Begin VB.Label lblitem 
         BackStyle       =   0  'Transparent
         Caption         =   "lblitem"
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
         Left            =   6840
         TabIndex        =   9
         Top             =   840
         Width           =   2565
      End
      Begin VB.Line Line3 
         X1              =   6825
         X2              =   9240
         Y1              =   1080
         Y2              =   1080
      End
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   8055
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   270
      Width           =   1845
      _extentx        =   3254
      _extenty        =   714
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
      Height          =   285
      Left            =   360
      TabIndex        =   21
      Top             =   4380
      Width           =   120
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
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
      Left            =   8265
      TabIndex        =   20
      Top             =   3540
      Width           =   60
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Batch Process   :"
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
      Left            =   4635
      TabIndex        =   19
      Top             =   3540
      Width           =   1860
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackColor       =   &H00FDDFE3&
      Caption         =   "Finish Good / WIP Manufacture (Summary)"
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
      TabIndex        =   7
      Top             =   600
      Width           =   9540
   End
End
Attribute VB_Name = "FrmFinishGoodWIPManufactureSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sql As String
Dim sqlLastBatch As String
Dim sqlDuty As String
Dim Curr(4) As String
Dim Idx As Long ' index utk row of excel
Dim xlColModel As String 'a
Dim xlColDesc As String 'b
Dim xlColIDRMaterial As String 'c
Dim xlColJPYMaterial As String 'd
Dim xlColUSDMaterial As String 'e
Dim xlColEURMaterial As String 'f
Dim xlColIDRContract As String 'g
Dim xlColUSDContract As String 'h
Dim rr As Long

'Materials
Dim TotIDRMaterial As Double
Dim TotJPYMaterial As Double
Dim TotUSDMaterial As Double
Dim TotEURMaterial As Double
'Contract
Dim TotIDRContract As Double, TotUSDContract As Double
'NonContract
Dim TotAMPLI As Double
Dim TotSPK1 As Double
Dim TotSPK2 As Double
Dim TotPCAT As Double
Dim TotPDIA As Double
Dim TotPINJ As Double
Dim TotPMTM As Double
Dim TotUnknown As Double
'Duty Price
Dim TotIDRDuty As Double
Dim TotJPYDuty As Double
Dim TotUSDDuty As Double
'Times
Dim TotTimes As Double

Private Sub cmdAction_Click(Index As Integer)
If cboGroup = "" Then
   LblErrMsg = DisplayMsg(8081)
   cboGroup.SetFocus
ElseIf CboItem = "" Then
   LblErrMsg = DisplayMsg(8082)
   CboItem.SetFocus
Else
   cboGroup = cboGroup
   CboItem = CboItem
               
   If cboGroup.MatchFound = False Then
     LblErrMsg = DisplayMsg(8083)
     cboGroup.SetFocus
   ElseIf CboItem.MatchFound = False Then
     LblErrMsg = DisplayMsg(8084)
     CboItem.SetFocus
   Else
    Me.MousePointer = vbHourglass
    LblErrMsg.Caption = ""
   
    Select Case Index
     Case 0: 'Process Batch
      On Error GoTo ErrHandlerBatch
      LblErrMsg.Caption = ""
      lblKet.Caption = ""
      Call SetControl(False)
      
      Dim db_Cek As New ADODB.Connection
      Dim db_Batch As New ADODB.Connection
      Dim rsCek As New ADODB.Recordset
      Dim rsBatch As New ADODB.Recordset
      
      db_Cek.Open Db.ConnectionString
      db_Batch.Open Db.ConnectionString
      
      sql = "select im.item_code, ml.Manufacture_Code mc from item_master im " & _
            "left join manufacture_line ml " & _
            "on im.Manufacture_Code = ml.Manufacture_Code and im.line_code = ml.line_code " & _
            "where production_cls = '01' and finishgoodpart_cls = '01' and use_endday >= convert(char(8), getdate(), 112) "
     
      If CboItem.ListIndex = 0 Then
       If cboGroup.ListIndex <> 0 Then sql = sql & "and group_cls = '" & Trim(cboGroup.Text) & "'"
      Else
       sql = sql & "and im.item_code = '" & Trim(CboItem.Text) & "' "
  
      End If
     
      If rsCek.State <> adStateClosed Then rsCek.Close
      rsCek.Open sql, db_Cek, adOpenKeyset, adLockOptimistic
            
      If Not rsCek.EOF Then
       Prg1.Max = rsCek.RecordCount
       
       Do While Not (rsCek.EOF)
        Prg1.Value = rsCek.AbsolutePosition
        lblKet.Caption = "Process Batch ... " & Round((rsCek.AbsolutePosition / Prg1.Max) * 100, 0) & " %"
        
        DoEvents
        Call SetTotal
        If Left((UCase(Trim(rsCek!Item_Code))), 1) = "L" Then 'Local---> Kena Pajak
         sqlDuty = "tax =  (Case (select count(bm2.item_code) from bom_master bm2 where bm2.parent_itemcode = bm.Item_Code) " & _
         Chr(10) & "when 0 then Isnull(hm.tax,0) else 0 end), "
        Else 'Export ----> Ngga Kena Pajak
         sqlDuty = "tax = 0, "
        End If
        Call HitungSummary(Trim(rsCek!Item_Code), 0, 1)
        
        Select Case (Trim(IIf(IsNull(rsCek!MC), "", rsCek!MC)))
         Case "AMPLI"
          TotAMPLI = TotAMPLI + TotUnknown
         Case "SPK1"
          TotSPK1 = TotSPK1 + TotUnknown
         Case "SPK2"
          TotSPK2 = TotSPK2 + TotUnknown
         Case "P-CAT"
          TotPCAT = TotPCAT + TotUnknown
         Case "P-DIA"
          TotPDIA = TotPDIA + TotUnknown
         Case "P-INJ"
          TotPINJ = TotPINJ + TotUnknown
         Case "P-MTM"
          TotPMTM = TotPMTM + TotUnknown
        End Select
       
        sql = "select * from Batch_Manufacture bt where bt.item_code = '" & Trim(rsCek!Item_Code) & "' "
        If rsBatch.State <> adStateClosed Then rsBatch.Close
        rsBatch.Open sql, db_Batch, adOpenKeyset, adLockOptimistic
        
        db_Batch.BeginTrans
        If rsBatch.EOF = True Then
        'Insert
         rsBatch.AddNew
         rsBatch!Item_Code = Trim(rsCek!Item_Code)
        End If
        
        'Update
        rsBatch!IDRMat = IIf(TotIDRMaterial = 0, Null, TotIDRMaterial)
        rsBatch!JPYMat = IIf(TotJPYMaterial = 0, Null, TotJPYMaterial)
        rsBatch!USDMat = IIf(TotUSDMaterial = 0, Null, TotUSDMaterial)
        rsBatch!EURMat = IIf(TotEURMaterial = 0, Null, TotEURMaterial)
        rsBatch!IDRCon = IIf(TotIDRContract = 0, Null, TotIDRContract)
        rsBatch!USDCon = IIf(TotUSDContract = 0, Null, TotUSDContract)
        rsBatch!AMPLI = IIf(TotAMPLI = 0, Null, TotAMPLI)
        rsBatch!SPK1 = IIf(TotSPK1 = 0, Null, TotSPK1)
        rsBatch!SPK2 = IIf(TotSPK2 = 0, Null, TotSPK2)
        rsBatch!PCAT = IIf(TotPCAT = 0, Null, TotPCAT)
        rsBatch!PDIA = IIf(TotPDIA = 0, Null, TotPDIA)
        rsBatch!PINJ = IIf(TotPINJ = 0, Null, TotPINJ)
        rsBatch!PMTM = IIf(TotPMTM = 0, Null, TotPMTM)
        rsBatch!IDRDuty = IIf(TotIDRDuty = 0, Null, TotIDRDuty)
        rsBatch!JPYDuty = IIf(TotJPYDuty = 0, Null, TotJPYDuty)
        rsBatch!USDDuty = IIf(TotUSDDuty = 0, Null, TotUSDDuty)
        rsBatch!Times = IIf(TotTimes = 0, Null, TotTimes)
        rsBatch!Last_Batch = Now
        rsBatch!Last_Update = Now
        rsBatch!last_user = Trim(userLogin)
        rsBatch.update
        
        db_Batch.CommitTrans
        If rsBatch.State <> adStateClosed Then rsBatch.Close
        rsCek.MoveNext
       Loop
       
       lblKet.Caption = "Process Complete"
       Call MaxLastBatch
      End If
ErrExitBatch:
        Me.MousePointer = vbDefault
        If rsCek.State <> adStateClosed Then rsCek.Close
        db_Cek.Close
        db_Batch.Close
        Call SetControl(True)
        Exit Sub
ErrHandlerBatch:
        db_Batch.RollbackTrans
        lblKet.Caption = ""
        LblErrMsg.Caption = "[" & err.number & "] " & err.Description
        err.clear
        Resume ErrExitBatch
     '***************************************************************************
     
     Case 1: 'Report To Excel
     
     On Error GoTo ErrHandlerExcel
     LblErrMsg.Caption = ""
     lblKet.Caption = ""
     Call SetControl(False)
 
     Dim rsTampil As New ADODB.Recordset
   
     sql = "select " & _
           "isnull(im.makeritem_code,'') model, bt.item_code, im.item_name, isnull(bt.IDRMat,0) IDRMat, isnull(bt.JPYMat,0) JPYMat, isnull(bt.USDMat,0) USDMat, isnull(bt.EURMat,0) EURMat, " & _
           "isnull(bt.IDRCon,0) IDRCon, isnull(bt.USDCon,0) USDCon, isnull(bt.AMPLI,0) AMPLI, isnull(bt.SPK1,0) SPK1, " & _
           "isnull(bt.SPK2,0) SPK2, isnull(bt.PCAT,0) PCAT, isnull(bt.PDIA,0) PDIA, isnull(bt.PINJ,0) PINJ, isnull(bt.PMTM,0) PMTM, " & _
           "isnull(bt.IDRDuty,0) IDRDuty, isnull(bt.JPYDuty,0) JPYDuty, isnull(bt.USDDuty,0) USDDuty, " & _
           "isnull(bt.Times,0) Times, bt.Last_Batch " & _
           "From " & _
           "batch_manufacture bt left join item_master im " & _
           "on bt.item_code = im.item_code " & _
           "where production_cls = '01' and finishgoodpart_cls = '01' and use_endday >= convert(char(8), getdate(), 112) "
     
     If CboItem.ListIndex = 0 Then
      If cboGroup.ListIndex <> 0 Then sql = sql & "and im.group_cls = '" & Trim(cboGroup.Text) & "'"
     Else
      sql = sql & "and bt.item_code = '" & Trim(CboItem.Text) & "' "
     End If
          
     Set rsTampil = Db.Execute(sql)
     
     If rsTampil.EOF Then
      LblErrMsg.Caption = DisplayMsg(4006)
      Set rsTampil = Nothing
      Call SetControl(True)
      Me.MousePointer = vbDefault
      Exit Sub
     End If
    
          
     Dim RsCurr As New ADODB.Recordset
     Dim rsManufacture As New ADODB.Recordset
     
     Dim xlapp As Excel.application
     Dim xlBook As Excel.Workbook
     Dim xlSheet As Excel.Worksheet

     Set xlapp = CreateObject("Excel.Application")
     Set xlBook = xlapp.Workbooks.Add
     Set xlSheet = xlBook.Worksheets("Sheet1")

     Screen.MousePointer = vbHourglass
     
     xlColModel = "a"
     xlColDesc = "b"
     xlColIDRMaterial = "c"
     xlColJPYMaterial = "d"
     xlColUSDMaterial = "e"
     xlColEURMaterial = "f"
     xlColIDRContract = "g"
     xlColUSDContract = "h"
     
     sql = "Select cc.Curr_Cls, cc.Description from curr_cls cc"
     Set RsCurr = Db.Execute(sql)

     'Currency Description
     rr = 0
     Do While Not (RsCurr.EOF)
      Curr(rr) = Trim(RsCurr!Description)
      RsCurr.MoveNext
      rr = rr + 1
     Loop

     With xlSheet

      '******************Fieldname************************
      Idx = 5

     .Range(xlColModel & Idx, xlColModel & Idx + 2).Merge
     .Range(xlColModel & Idx) = "Model"

     .Range(xlColDesc & Idx, xlColDesc & Idx + 2).Merge
     .Range(xlColDesc & Idx) = "Description"

     .Range(xlColIDRMaterial & Idx, xlColEURMaterial & Idx + 1).Merge
     .Range(xlColIDRMaterial & Idx) = "Materials"
     .Range(xlColIDRMaterial & Idx + 2) = Curr(2)
     .Range(xlColJPYMaterial & Idx + 2) = Curr(0)
     .Range(xlColUSDMaterial & Idx + 2) = Curr(1)
     .Range(xlColEURMaterial & Idx + 2) = Curr(3)

     .Range(xlColIDRContract & Idx, xlColUSDContract & Idx + 1).Merge
     .Range(xlColIDRContract & Idx) = "Contract"
     .Range(xlColIDRContract & Idx + 2) = Curr(2)
     .Range(xlColUSDContract & Idx + 2) = Curr(1)

     sql = "Select * from Manufacture_Line order by line_code "
     Set rsManufacture = Db.Execute(sql)

     rr = 0
     Do While Not (rsManufacture.EOF)
      .Cells(Idx + 1, rr + 9) = Trim(rsManufacture!Manufacture_Code)
      .Cells(Idx + 2, rr + 9) = Curr(2)
      rsManufacture.MoveNext
      rr = rr + 1
     Loop
        
     'rr skrg jd 7

     .Range(.Cells(Idx, 9), .Cells(Idx, rr + 8)).Merge
     .Cells(Idx, 9) = "Process Cost"
     
     .Cells(Idx + 2, rr + 9) = Curr(2)
     .Cells(Idx + 2, rr + 10) = Curr(0)
     .Cells(Idx + 2, rr + 11) = Curr(1)
     .Range(.Cells(Idx, rr + 9), .Cells(Idx + 1, rr + 11)).Merge
     .Cells(Idx, rr + 9) = "Duty Price"
     
     .Range(.Cells(Idx, rr + 12), .Cells(Idx + 2, rr + 12)).Merge
     .Cells(Idx, rr + 12) = "Product" & Chr(10) & "Times"

     'Bold
     .Range(xlColModel & Idx, .Cells(Idx + 2, rr + 12)).Font.Bold = True

     '****************************TITLE***************************************
     .Range(xlColModel & "2", .Cells(2, rr + 12)).Merge
     .Range(xlColModel & "2") = "FINISH GOOD / WIP MANUFACTURE"
     .Range(xlColModel & "3", .Cells(3, rr + 12)).Merge
     .Range(xlColModel & "3") = "(SUMMARY)"
     .Range(xlColModel & "2", .Cells(3, rr + 12)).Columns.Font.Size = "10"
     .Range(xlColModel & "2", .Cells(3, rr + 12)).Columns.Font.Bold = True
     .Range(xlColModel & "2", .Cells(3, rr + 12)).Columns.Font.Name = "Arial"
     .Range(xlColModel & "2", .Cells(3, rr + 12)).HorizontalAlignment = xlCenter

     '********************** border ******************************
     .Range(xlColModel & Idx, .Cells(Idx + 2, rr + 12)).Borders(xlEdgeLeft).LineStyle = xlContinuous
     .Range(xlColModel & Idx, .Cells(Idx + 2, rr + 12)).Borders(xlEdgeTop).LineStyle = xlContinuous
     .Range(xlColModel & Idx, .Cells(Idx + 2, rr + 12)).Borders(xlEdgeBottom).LineStyle = xlContinuous
     .Range(xlColModel & Idx, .Cells(Idx + 2, rr + 12)).Borders(xlEdgeRight).LineStyle = xlContinuous
     .Range(xlColModel & Idx, .Cells(Idx + 2, rr + 12)).Borders(xlInsideVertical).LineStyle = xlContinuous
     .Range(xlColModel & Idx, .Cells(Idx + 2, rr + 12)).Borders(xlInsideHorizontal).LineStyle = xlContinuous

     '********************** alignment field *************************
     .Range(xlColModel & Idx, .Cells(Idx + 2, rr + 12)).HorizontalAlignment = xlCenter
     .Range(xlColModel & Idx, .Cells(Idx + 2, rr + 12)).VerticalAlignment = xlCenter

     Dim i As Long
     i = 0

     DoEvents
     Do While Not (rsTampil.EOF)
      .Range(xlColModel & Idx + 3 + i) = Trim(rsTampil!Model)
      .Range(xlColDesc & Idx + 3 + i) = Trim(rsTampil!Item_Code) & " / " & Trim(rsTampil!item_name)
      .Range(xlColIDRMaterial & Idx + 3 + i) = IIf(rsTampil!IDRMat = 0, "", rsTampil!IDRMat)
      .Range(xlColJPYMaterial & Idx + 3 + i) = IIf(rsTampil!JPYMat = 0, "", rsTampil!JPYMat)
      .Range(xlColUSDMaterial & Idx + 3 + i) = IIf(rsTampil!USDMat = 0, "", rsTampil!USDMat)
      .Range(xlColEURMaterial & Idx + 3 + i) = IIf(rsTampil!EURMat = 0, "", rsTampil!EURMat)
      .Range(xlColIDRContract & Idx + 3 + i) = IIf(rsTampil!IDRCon = 0, "", rsTampil!IDRCon)
      .Range(xlColUSDContract & Idx + 3 + i) = IIf(rsTampil!USDCon = 0, "", rsTampil!USDCon)

      .Cells(Idx + 3 + i, rr + 2) = IIf(rsTampil!AMPLI = 0, "", rsTampil!AMPLI)
      .Cells(Idx + 3 + i, rr + 3) = IIf(rsTampil!SPK1 = 0, "", rsTampil!SPK1)
      .Cells(Idx + 3 + i, rr + 4) = IIf(rsTampil!SPK2 = 0, "", rsTampil!SPK2)
      .Cells(Idx + 3 + i, rr + 5) = IIf(rsTampil!PCAT = 0, "", rsTampil!PCAT)
      .Cells(Idx + 3 + i, rr + 6) = IIf(rsTampil!PDIA = 0, "", rsTampil!PDIA)
      .Cells(Idx + 3 + i, rr + 7) = IIf(rsTampil!PINJ = 0, "", rsTampil!PINJ)
      .Cells(Idx + 3 + i, rr + 8) = IIf(rsTampil!PMTM = 0, "", rsTampil!PMTM)

      .Cells(Idx + 3 + i, rr + 9) = IIf(rsTampil!IDRDuty = 0, "", rsTampil!IDRDuty)
      .Cells(Idx + 3 + i, rr + 10) = IIf(rsTampil!JPYDuty = 0, "", rsTampil!JPYDuty)
      .Cells(Idx + 3 + i, rr + 11) = IIf(rsTampil!USDDuty = 0, "", rsTampil!USDDuty)
      
      .Cells(Idx + 3 + i, rr + 12) = IIf(rsTampil!Times = 0, "", rsTampil!Times)
      rsTampil.MoveNext
       i = i + 1
      Loop

     '*********************FormatTampilan*******************************************
     .Range(xlColIDRMaterial & Idx + 3 & ":" & xlColIDRMaterial & Idx + 2 + i & "," & xlColIDRContract & Idx + 3 & ":" & xlColIDRContract & Idx + 2 + i).NumberFormat = gs_formatAmountIDR
     .Range(xlColJPYMaterial & Idx + 3 & ":" & xlColEURMaterial & Idx + 2 + i & "," & xlColUSDContract & Idx + 3 & ":" & xlColUSDContract & Idx + 2 + i).NumberFormat = gs_formatAmount
     .Range(.Cells(Idx + 3, 9), .Cells(Idx + 2 + i, rr + 8)).NumberFormat = gs_formatAmountIDR 'Process Cost
     .Range(.Cells(Idx + 3, rr + 9), .Cells(Idx + 2 + i, rr + 9)).NumberFormat = gs_formatAmountIDR 'IDRDuty
     .Range(.Cells(Idx + 3, rr + 10), .Cells(Idx + 2 + i, rr + 11)).NumberFormat = gs_formatAmount 'JPYDuty n USDDuty
     .Range(.Cells(Idx + 3, rr + 12), .Cells(Idx + 2 + i, rr + 12)).NumberFormat = gs_formatQtyBOM 'Times
     .Range(xlColModel & Idx, .Cells(Idx + 2 + i, rr + 12)).Columns.Font.Name = "Arial"
     .Range(xlColModel & Idx, .Cells(Idx + 2 + i, rr + 12)).Columns.Font.Size = 8
     
     '************ width ******************
     '.Range(xlColIDRMaterial & ":" & Chr(Asc("h") + rr + 11)).ColumnWidth = 12
     .Range(xlColModel & ":" & xlColDesc).Columns.AutoFit
     '.Columns(xlColModel).AutoFit
     '.Columns(xlColDesc).AutoFit
         
   'Border
     
    With .Range(xlColModel & Idx + 3, .Cells(Idx + 2 + i, rr + 12))
     .Borders(xlEdgeLeft).LineStyle = xlContinuous
     .Borders(xlEdgeTop).LineStyle = xlContinuous
     .Borders(xlEdgeBottom).LineStyle = xlContinuous
     .Borders(xlEdgeRight).LineStyle = xlContinuous
     .Borders(xlInsideVertical).LineStyle = xlContinuous
     If (Idx + 2 + i) > 8 Then .Borders(xlInsideHorizontal).LineStyle = xlContinuous
    End With
        
   With .PageSetup
    .LeftMargin = application.InchesToPoints(0.75)
    .RightMargin = application.InchesToPoints(0.75)
    .TopMargin = application.InchesToPoints(1)
    .BottomMargin = application.InchesToPoints(1)
    .HeaderMargin = application.InchesToPoints(0.5)
    .FooterMargin = application.InchesToPoints(0.5)
    .PrintHeadings = False
    .PrintGridlines = False
    .CenterHorizontally = True
    .CenterVertically = False
    .Orientation = xlLandscape
   End With
 End With
 
 xlapp.WindowState = xlMaximized
 xlapp.Visible = True

ErrExitExcel:
    Me.MousePointer = vbDefault
    Screen.MousePointer = vbDefault
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlapp = Nothing
    Set RsCurr = Nothing
    Set rsManufacture = Nothing
    Set rsTampil = Nothing
    Call SetControl(True)
    Exit Sub
ErrHandlerExcel:
    LblErrMsg.Caption = "[" & err.number & "] " & err.Description
    err.clear
    Resume ErrExitExcel
    '******************************************************************************
    End Select
   End If
 End If
End Sub
Sub SetTotal()
'Total Materials
TotIDRMaterial = 0
TotJPYMaterial = 0
TotUSDMaterial = 0
TotEURMaterial = 0
'Total Contract
TotIDRContract = 0
TotUSDContract = 0
'Total Process Cost
TotAMPLI = 0
TotSPK1 = 0
TotSPK2 = 0
TotPCAT = 0
TotPDIA = 0
TotPINJ = 0
TotPMTM = 0
TotUnknown = 0
'Duty Price
TotIDRDuty = 0
TotJPYDuty = 0
TotUSDDuty = 0
'Times
TotTimes = 0
End Sub
Sub HitungSummary(ibu As String, lvl As Integer, qtyinduk As Double)
Dim anak As String
Dim rsAnak As New ADODB.Recordset
Dim SqlExchRate As String

'********VarPenampung********
Dim qtyawal As Double
Dim qtyakhir As Double
Dim cost_minute As Double
Dim part_cost As Double
Dim process_cost As Double
Dim tax As Double
Dim duty_price As Double
Dim amountpartcost As Double
Dim amountprocesscost As Double
Dim ExchRate As Double
'*****************************

'******VarPembantu*********
Dim formatprice As String
Dim FormatAmount As String
'**************************

     '******************************ExchangeRate*************************/*********
    SqlExchRate = Chr(10) & "IsNull((Select ber.Exch0" & dtAwal.Month & " as ExchRate " & _
                  Chr(10) & "from Book_ExchangeRate ber,Company_Profile cp " & _
                  Chr(10) & "Where ber.Term_Cls = cp.ValuationPrice_ExchTerm " & _
                  Chr(10) & "and ber.Exch_Year = '" & dtAwal.Year & "' " & _
                  Chr(10) & "and ber.Currency_Code = z.Curr_Code),0) "
    '********************************Material**************************************
    sql = "Select *, ExchRate = " & SqlExchRate & "from ( "
    sql = sql & _
          Chr(10) & "--Material" & _
          Chr(10) & "Select '0' idx, bm.parent_itemcode, bm.Item_Code,im.Item_Name, " & _
          Chr(10) & "bm.qty, bm.unit_cls, " & _
          Chr(10) & "trade_code = (Case (select count(bm2.item_code) " & _
          Chr(10) & "from bom_master bm2 where bm2.parent_itemcode = bm.Item_Code) " & _
          Chr(10) & "when 0 then  im.supplier_code Else '' end), " & _
          Chr(10) & "trade_name = '', cost_minute=0, process_cls='',process_cost = 0, "
    sql = sql & _
          Chr(10) & "part_cost = (Case (select count(bm2.item_code) from bom_master bm2 where bm2.parent_itemcode = bm.Item_Code) " & _
          Chr(10) & "when 0 then " & _
          Chr(10) & "(isnull((select top 1 prim.price " & _
          Chr(10) & "from price_master prim " & _
          Chr(10) & "Where prim.item_code = bm.item_code " & _
          Chr(10) & "and prim.price_cls = '01' " & _
          Chr(10) & "and prim.trade_code = im.Supplier_Code " & _
          Chr(10) & "and (prim.start_date <= '" & Format(dtAwal.Value, "yyyymmdd") & "' " & _
          Chr(10) & "and prim.end_date >= '" & Format(dtAwal.Value, "yyyymmdd") & "') " & _
          Chr(10) & "order by prim.priority_cls desc),0) ) " & _
          Chr(10) & "else 0 end ), "
    sql = sql & _
          Chr(10) & "curr_code = (Case (select count(bm2.item_code) from bom_master bm2 where bm2.parent_itemcode = bm.Item_Code) " & _
          Chr(10) & "when 0 then (select top 1 prim.currency_code from price_master prim " & _
          Chr(10) & "where prim.item_code = bm.item_code and prim.price_cls = '01' " & _
          Chr(10) & "and prim.trade_code = im.Supplier_Code " & _
          Chr(10) & "and (prim.start_date <= '" & Format(dtAwal.Value, "yyyymmdd") & "' " & _
          Chr(10) & "and prim.end_date >= '" & Format(dtAwal.Value, "yyyymmdd") & "') " & _
          Chr(10) & "order by prim.priority_cls desc) else '' end), "
    sql = sql & _
          Chr(10) & sqlDuty & _
          Chr(10) & "isnull(ml.manufacture_code,'') manufacture_code "
    sql = sql & _
          Chr(10) & "from BOM_Master bm,Item_Master im, HS_Master hm, Manufacture_Line ml where " & _
          Chr(10) & "im.HS_Code *= hm.HS_Code " & _
          Chr(10) & "and im.Manufacture_Code *= ml.Manufacture_Code " & _
          Chr(10) & "and im.Line_Code *= ml.Line_code " & _
          Chr(10) & "and bm.Item_Code = im.Item_Code and bm.parent_itemcode = '" & ibu & "' " & _
          Chr(10) & "and (bm.start_date <= '" & Format(dtAwal.Value, "yyyymmdd") & "' " & _
          Chr(10) & "and bm.end_date >= '" & Format(dtAwal.Value, "yyyymmdd") & "') "
    '********************************* WIP **************************************************
    sql = sql & _
          Chr(10) & _
          Chr(10) & "UNION ALL " & _
          Chr(10)
    sql = sql & _
          Chr(10) & "--WIP" & _
          Chr(10) & "select '1' idx, prom.item_code, item_code = '', pros.description process_name, " & _
          Chr(10) & "qty = prom.standard_time, " & _
          Chr(10) & "unit_cls = (case len(isnull(prom.trade_code,'')) when 0 then '0' else '1' end) , " & _
          Chr(10) & "prom.trade_code, tm.trade_name, isnull(prom.cost_minute,0) cost_minute, prom.process_cls, "
    sql = sql & _
          Chr(10) & "process_cost = " & _
          Chr(10) & "--purchase " & _
          Chr(10) & "(case len(isnull(prom.trade_code,'')) when 0 then prom.cost_minute " & _
          Chr(10) & "else ( " & _
          Chr(10) & "isnull((select top 1 prim.price from price_master prim where prim.trade_code = prom.trade_code and " & _
          Chr(10) & "prim.price_cls = '01' and prim.item_code=prom.item_code and (prim.start_date <= '" & Format(dtAwal.Value, "yyyymmdd") & "' and " & _
          Chr(10) & "prim.end_date >= '" & Format(dtAwal.Value, "yyyymmdd") & "') order by prim.priority_cls desc),0) " & _
          Chr(10) & "* " & _
          Chr(10) & "(case " & _
          Chr(10) & "(isnull((select top 1 prim.price from price_master prim where prim.trade_code = prom.trade_code and " & _
          Chr(10) & "prim.price_cls = '05' and prim.item_code=prom.item_code and (prim.start_date <= '" & Format(dtAwal.Value, "yyyymmdd") & "' and " & _
          Chr(10) & "prim.end_date >= '" & Format(dtAwal.Value, "yyyymmdd") & "') order by prim.priority_cls desc),0) " & _
          Chr(10) & ") when 0 Then 1 "
    sql = sql & _
          Chr(10) & "else ( " & _
          Chr(10) & "-- jika tdk ada price di service pake curr sendr, sebalikny pake curr di service (1) jika sama currnya " & _
          Chr(10) & "-- jika curr di purchase beda dengan service, dianggap (0) " & _
          Chr(10) & "case " & _
          Chr(10) & "(select top 1 prim.currency_code from price_master prim " & _
          Chr(10) & "where prim.trade_code = prom.trade_code and prim.price_cls = '01' and prim.item_code=prom.item_code " & _
          Chr(10) & "and (prim.start_date <= '" & Format(dtAwal.Value, "yyyymmdd") & "' and prim.end_date >= '" & Format(dtAwal.Value, "yyyymmdd") & "') " & _
          Chr(10) & "order by prim.priority_cls desc) " & _
          Chr(10) & "when ( " & _
          Chr(10) & "--cek sama ato ngga dg curr yg di service " & _
          Chr(10) & "select top 1 prim.currency_code from price_master prim " & _
          Chr(10) & "where prim.trade_code = prom.trade_code and prim.price_cls = '05' and prim.item_code=prom.item_code " & _
          Chr(10) & "and (prim.start_date <= '" & Format(dtAwal.Value, "yyyymmdd") & "' and prim.end_date >= '" & Format(dtAwal.Value, "yyyymmdd") & "') " & _
          Chr(10) & "order by prim.priority_cls desc " & _
          Chr(10) & ") then 1 else 0 end)  end) " & _
          Chr(10) & "+ "
    sql = sql & _
          Chr(10) & "--service " & _
          Chr(10) & "(isnull((select top 1 prim.price from price_master prim where prim.trade_code = prom.trade_code and " & _
          Chr(10) & "prim.price_cls = '05' and prim.item_code=prom.item_code and (prim.start_date <= '" & Format(dtAwal.Value, "yyyymmdd") & "' and " & _
          Chr(10) & "prim.end_date >= '" & Format(dtAwal.Value, "yyyymmdd") & "') order by prim.priority_cls desc),0)) " & _
          Chr(10) & ") end), "
    sql = sql & _
          Chr(10) & "part_cost = 0, "
    sql = sql & _
          Chr(10) & "curr_code = " & _
          Chr(10) & "(case len(isnull(prom.trade_code,'')) when 0 then prom.Currency_Code " & _
          Chr(10) & "else ( " & _
          Chr(10) & "case " & _
          Chr(10) & "(select count (Y.currency_code) from " & _
          Chr(10) & "(select top 1 prim.currency_code from price_master prim " & _
          Chr(10) & "where prim.trade_code = prom.trade_code and prim.price_cls = '05' and prim.item_code=prom.item_code " & _
          Chr(10) & "and (prim.start_date <= '" & Format(dtAwal.Value, "yyyymmdd") & "' and prim.end_date >= '" & Format(dtAwal.Value, "yyyymmdd") & "') " & _
          Chr(10) & "order by prim.priority_cls desc) Y) "
    sql = sql & _
          Chr(10) & "when  0 then " & _
          Chr(10) & "( " & _
          Chr(10) & "select top 1 prim.currency_code from price_master prim " & _
          Chr(10) & "where prim.trade_code = prom.trade_code and prim.price_cls = '01' and prim.item_code=prom.item_code " & _
          Chr(10) & "and (prim.start_date <= '" & Format(dtAwal.Value, "yyyymmdd") & "' and prim.end_date >= '" & Format(dtAwal.Value, "yyyymmdd") & "') " & _
          Chr(10) & "order by prim.priority_cls desc) " & _
          Chr(10) & "Else " & _
          Chr(10) & "( " & _
          Chr(10) & "select top 1 prim.currency_code from price_master prim " & _
          Chr(10) & "where prim.trade_code = prom.trade_code and prim.price_cls = '05' and prim.item_code=prom.item_code " & _
          Chr(10) & "and (prim.start_date <= '" & Format(dtAwal.Value, "yyyymmdd") & "' and prim.end_date >= '" & Format(dtAwal.Value, "yyyymmdd") & "') " & _
          Chr(10) & "order by prim.priority_cls desc) end) " & _
          Chr(10) & "end), "
    sql = sql & _
          Chr(10) & "tax = 0, " & _
          Chr(10) & "isnull(ml.manufacture_code,'') manufacture_code " & _
          Chr(10) & "from process_master prom, process_cls pros, trade_master tm, item_master im, manufacture_line ml  where " & _
          Chr(10) & "prom.process_cls = pros.process_cls and " & _
          Chr(10) & "prom.trade_code *= tm.trade_code and " & _
          Chr(10) & "prom.item_code = im.item_code and " & _
          Chr(10) & "im.manufacture_Code *= ml.manufacture_code and " & _
          Chr(10) & "im.line_code *= ml.Line_Code and " & _
          Chr(10) & "prom.item_code = '" & ibu & "' "
    '********************************************************************************
    sql = sql & _
          Chr(10) & ")Z order by idx, item_code "
          
    Set rsAnak = Db.Execute(sql)
       
    
  If Not rsAnak.EOF Then
    
    Do While Not rsAnak.EOF
      
      If Trim(rsAnak!curr_code) = "03" Then
        FormatAmount = gs_formatAmountIDR
        formatprice = gs_formatPriceIDR
      Else
        FormatAmount = gs_formatAmount
        formatprice = gs_formatPrice
      End If
      
      anak = Trim(rsAnak!Item_Code)
      qtyawal = Format(rsAnak!Qty, gs_formatQtyBOM)
      qtyakhir = Format((qtyinduk * qtyawal), gs_formatQtyBOM)
      cost_minute = Format(rsAnak!cost_minute, formatprice)
      part_cost = Format(rsAnak!part_cost, formatprice)
      process_cost = Format(rsAnak!process_cost, formatprice)
      tax = Format(((rsAnak!tax) / 100), gs_formatPercentage)
      duty_price = Format((tax * part_cost), formatprice)
      ExchRate = Format(rsAnak!ExchRate, gs_formatExchangeRate)
             
     If rsAnak!part_cost <> 0 Then
      amountpartcost = Format((qtyakhir * part_cost), FormatAmount)
     Else
      amountpartcost = 0
     End If

     If (rsAnak!process_cost <> 0) Then
      'subcon
      If (Trim(IIf(IsNull(rsAnak!Trade_Code), "", rsAnak!Trade_Code)) <> "") Then
       'processcost utk 1 minutenya
       process_cost = Format((process_cost / qtyawal), formatprice)
       TotTimes = 0
      Else
       TotTimes = TotTimes + qtyakhir 'times utk nonsubcon
      End If
      'perlu qtyakhir krn processcostnya hanya utk 1 minute
      amountprocesscost = Format((qtyakhir * process_cost), FormatAmount)
     Else
      amountprocesscost = 0
      TotTimes = 0
     End If
     
     If rsAnak!part_cost = 0 And rsAnak!process_cost = 0 Then
      amountpartcost = 0
      amountprocesscost = 0
     End If
     
     Select Case Trim(rsAnak!curr_code)
      Case "01"
       TotJPYMaterial = TotJPYMaterial + Format(amountpartcost, FormatAmount)
       TotJPYDuty = TotJPYDuty + Format((qtyakhir * duty_price), FormatAmount)
       If (Trim(IIf(IsNull(rsAnak!Trade_Code), "", rsAnak!Trade_Code)) <> "") Then
        TotIDRContract = TotIDRContract + Format((amountprocesscost * ExchRate), gs_formatAmountIDR)
       End If
      Case "02"
       TotUSDMaterial = TotUSDMaterial + Format(amountpartcost, FormatAmount)
       TotUSDDuty = TotUSDDuty + Format((qtyakhir * duty_price), FormatAmount)
       If (Trim(IIf(IsNull(rsAnak!Trade_Code), "", rsAnak!Trade_Code)) <> "") Then
        TotUSDContract = TotUSDContract + Format(amountprocesscost, FormatAmount)
       End If
      Case "03"
       TotIDRMaterial = TotIDRMaterial + Format(amountpartcost, FormatAmount)
       TotIDRDuty = TotIDRDuty + Format((qtyakhir * duty_price), FormatAmount)
       If (Trim(IIf(IsNull(rsAnak!Trade_Code), "", rsAnak!Trade_Code)) <> "") Then
        TotIDRContract = TotIDRContract + Format(amountprocesscost, FormatAmount)
       End If
      Case "04"
       TotEURMaterial = TotEURMaterial + Format(amountpartcost, FormatAmount)
       'TotIDRMaterial = TotIDRMaterial + Format((amountpartcost * ExchRate), gs_formatAmountIDR)
       TotIDRDuty = TotIDRDuty + Format((qtyakhir * duty_price * ExchRate), gs_formatAmountIDR)
       If (Trim(IIf(IsNull(rsAnak!Trade_Code), "", rsAnak!Trade_Code)) <> "") Then
        TotIDRContract = TotIDRContract + Format((amountprocesscost * ExchRate), gs_formatAmountIDR)
       End If
      Case "05"
       TotIDRMaterial = TotIDRMaterial + Format((amountpartcost * ExchRate), gs_formatAmountIDR)
       TotIDRDuty = TotIDRDuty + Format((qtyakhir * duty_price * ExchRate), gs_formatAmountIDR)
       If (Trim(IIf(IsNull(rsAnak!Trade_Code), "", rsAnak!Trade_Code)) <> "") Then
        TotIDRContract = TotIDRContract + Format((amountprocesscost * ExchRate), gs_formatAmountIDR)
       End If
      End Select
        
     Select Case (Trim(rsAnak!Manufacture_Code))
      Case "AMPLI"
       If (Trim(IIf(IsNull(rsAnak!Trade_Code), "", rsAnak!Trade_Code)) = "") Then
         TotAMPLI = TotAMPLI + Format((amountprocesscost * ExchRate), gs_formatAmountIDR)
       End If
      Case "SPK1"
       If (Trim(IIf(IsNull(rsAnak!Trade_Code), "", rsAnak!Trade_Code)) = "") Then
         TotSPK1 = TotSPK1 + Format((amountprocesscost * ExchRate), gs_formatAmountIDR)
       End If
      Case "SPK2"
       If (Trim(IIf(IsNull(rsAnak!Trade_Code), "", rsAnak!Trade_Code)) = "") Then
         TotSPK2 = TotSPK2 + Format((amountprocesscost * ExchRate), gs_formatAmountIDR)
       End If
      Case "P-CAT"
       If (Trim(IIf(IsNull(rsAnak!Trade_Code), "", rsAnak!Trade_Code)) = "") Then
         TotPCAT = TotPCAT + Format((amountprocesscost * ExchRate), gs_formatAmountIDR)
       End If
      Case "P-DIA"
       If (Trim(IIf(IsNull(rsAnak!Trade_Code), "", rsAnak!Trade_Code)) = "") Then
         TotPDIA = TotPDIA + Format((amountprocesscost * ExchRate), gs_formatAmountIDR)
       End If
      Case "P-INJ"
       If (Trim(IIf(IsNull(rsAnak!Trade_Code), "", rsAnak!Trade_Code)) = "") Then
         TotPINJ = TotPINJ + Format((amountprocesscost * ExchRate), gs_formatAmountIDR)
       End If
      Case "P-MTM"
       If (Trim(IIf(IsNull(rsAnak!Trade_Code), "", rsAnak!Trade_Code)) = "") Then
         TotPMTM = TotPMTM + Format((amountprocesscost * ExchRate), gs_formatAmountIDR)
       End If
      Case Else
       If (Trim(IIf(IsNull(rsAnak!Trade_Code), "", rsAnak!Trade_Code)) = "") Then
         TotUnknown = TotUnknown + Format((amountprocesscost * ExchRate), gs_formatAmountIDR)
       End If
     End Select

     
     Call HitungSummary(anak, lvl + 1, rsAnak!Qty * qtyinduk)
     
     If Not (rsAnak.EOF) Then rsAnak.MoveNext
            
     Loop
  End If
  Set rsAnak = Nothing

End Sub

Private Sub cmdBrowser_Click()
 Me.MousePointer = vbHourglass
 frm_BrowseItem.getItemCode = CboItem.Text
 frm_BrowseItem.Show 1
 CboItem.Text = frm_BrowseItem.getItemCode
 Me.MousePointer = vbDefault
End Sub

Private Sub CmdSubMenu_Click()
DoEvents
frmMainMenu.Show
DoEvents
Unload Me
End Sub
Private Sub Form_Load()
If gb_Simulation = True Then Call up_InitSimulation(Me)
CtrlMenu1.FormName = Me.Name
Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
Call Kosong
Call adtocombo
dtAwal = Now
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
Sub Kosong()
lblgroup = ""
lblitem(0) = ""
lblitem(1) = ""
End Sub

Sub adtocombo()
'*******Group Cls**********
Call up_FillCombo(cboGroup, "Group_Cls", , , True)
cboGroup.ListWidth = 150
cboGroup.ColumnWidths = "30 pt;120 pt"
cboGroup.ListIndex = 0
End Sub

Private Sub adtocomboitem()
Dim adoRs As New ADODB.Recordset
With CboItem
    
    .clear
    .columnCount = 3
    
    sql = "select item_code, makeritem_code, item_name from item_master " & _
        "where production_cls = '01' and finishgoodpart_cls = '01' and use_endday >= convert(char(8), getdate(), 112) "
    If cboGroup.ListIndex <> 0 Then sql = sql & "and group_cls = '" & Trim(cboGroup.Text) & "'"

    adoRs.Open sql, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    .AddItem ""
    .List(.ListCount - 1, 0) = strAll
    .List(.ListCount - 1, 1) = strAll
    .List(.ListCount - 1, 2) = strAll
    
    While Not adoRs.EOF
        .AddItem ""
        .List(.ListCount - 1, 0) = Trim(adoRs.Fields("item_code"))
        .List(.ListCount - 1, 1) = Trim(adoRs.Fields("makeritem_code"))
        .List(.ListCount - 1, 2) = Trim(adoRs.Fields("item_name"))
        adoRs.MoveNext
    Wend
    adoRs.Close
    
    .ListWidth = 410
    .ColumnWidths = "130 pt;130 pt;150 pt"
    
End With
Set adoRs = Nothing
End Sub

Private Sub cboGroup_Change()
lblgroup = ""
LblErrMsg = ""
End Sub

Private Sub CboGroup_Click()
    If cboGroup.MatchFound Then
        lblgroup = cboGroup.Column(1)
        LblErrMsg = ""
    Else
        lblgroup = ""
        LblErrMsg = DisplayMsg(8083)
        cboGroup.SetFocus
    End If
    adtocomboitem
    CboItem.ListIndex = 0
    Call MaxLastBatch
End Sub

Private Sub CboGroup_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 13 Then Call CboGroup_Click
End Sub

Private Sub CboItem_Change()
lblitem(0) = ""
lblitem(1) = ""
LblErrMsg = ""
End Sub

Private Sub cboitem_Click()
    If CboItem.MatchFound Then
        lblitem(0) = CboItem.Column(1)
        lblitem(1) = CboItem.Column(2)
        LblErrMsg = ""
    Else
        lblitem(0) = ""
        lblitem(1) = ""
        LblErrMsg = DisplayMsg(8084)
    End If
     
    Call MaxLastBatch
End Sub
Private Sub cboitem_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 13 Then Call cboitem_Click
End Sub

Sub MaxLastBatch()
Label4.Caption = ""
Dim rsMaxLastBatch As New ADODB.Recordset
If CboItem.ListIndex = 0 Then
  If cboGroup.ListIndex > 0 Then
   sqlLastBatch = "select max(bt.last_batch) last_batch " & _
                  "from Batch_Manufacture bt left join Item_Master im " & _
                  "on bt.Item_Code = im.Item_Code " & _
                  "where im.group_cls = '" & cboGroup.Text & "' " & _
                  "group by im.Group_Cls "
  Else
   sqlLastBatch = "select max(last_batch) last_batch from Batch_Manufacture "
  End If
Else
 sqlLastBatch = "select last_batch from batch_manufacture where item_code = '" & CboItem.Text & "' "
End If

Set rsMaxLastBatch = Db.Execute(sqlLastBatch)

If Not rsMaxLastBatch.EOF Then
 Label4.Caption = Format((Trim(rsMaxLastBatch!Last_Batch)), "DD MMM YYYY hh:mm")
Else
 Label4.Caption = ""
End If

Set rsMaxLastBatch = Nothing
End Sub

Sub SetControl(Flag As Boolean)
 Prg1.Value = 0
 cmdsubmenu.Enabled = Flag
 cmdAction(0).Enabled = Flag
 cmdAction(1).Enabled = Flag
End Sub

