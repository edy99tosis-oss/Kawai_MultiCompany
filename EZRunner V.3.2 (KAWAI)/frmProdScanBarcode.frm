VERSION 5.00
Begin VB.Form frmProdScanBarcode 
   BackColor       =   &H00FDDFE3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scan Barcode"
   ClientHeight    =   2040
   ClientLeft      =   3645
   ClientTop       =   6705
   ClientWidth     =   6885
   Icon            =   "frmProdScanBarcode.frx":0000
   KeyPreview      =   -1  'True
   LinkMode        =   1  'Source
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   6885
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameMsg 
      BackColor       =   &H00FDDFE3&
      Height          =   600
      Left            =   225
      TabIndex        =   2
      Top             =   1215
      Width           =   6435
      Begin VB.Label lblErrMsg 
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
         Height          =   285
         Left            =   105
         TabIndex        =   3
         Top             =   195
         Width           =   6210
      End
   End
   Begin VB.TextBox txtBarcode 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   225
      MaxLength       =   25
      TabIndex        =   1
      Top             =   720
      Width           =   6435
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Scan Barcode"
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
      Left            =   2685
      TabIndex        =   0
      Top             =   135
      Width           =   1515
   End
End
Attribute VB_Name = "frmProdScanBarcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ShowData()
    Dim adoRs As New ADODB.Recordset
    
    LblErrMsg.Caption = ""
    
    sql = "select dp.seq_no, dp.factory_code, dp.line_code, dp.item_code, dp.lot_no, dp.qty, dp.unit_cls, dp.schedule_date, dp.complete_cls, " & _
        "im.item_name, im.wh_code, im.makeritem_code, cc.description unit_desc, " & _
        "qty_result = isnull((select sum(qty) from part_receipt where receipt_cls = 'P1' and dailyseq_no = dp.seq_no), 0) " & _
        "from daily_production dp " & _
        "inner join item_master im on dp.item_code = im.item_code " & _
        "left join unit_cls cc on dp.unit_cls = cc.unit_cls " & _
        "where dp.prod_barcode = '" & txtBarcode.Text & "' "
        
    adoRs.Open sql, Db, adOpenStatic, adLockReadOnly, adCmdText
    If Not adoRs.EOF Then
        frmProdScanBarcode.Hide
        With frmProdResult
            .Cbo(0) = Trim(adoRs.Fields("factory_code"))
            .Cbo(1) = Trim(adoRs.Fields("Line_code"))
            .Cbo(2) = Trim(adoRs.Fields("wh_code"))
            .Cbo(3) = Trim(adoRs.Fields("item_code"))

            .cboResultCls.ListIndex = 0

            .txtLot = adoRs.Fields("lot_no")
            .txtQty = Format(0, gs_formatQty)
            .txtRemarks = ""
            .txtUnit = adoRs.Fields("unit_desc")

            .tglProd = adoRs.Fields("schedule_date")
            .is_LoadByItemCode = Trim(adoRs.Fields("item_code"))
            .dailyseqno = adoRs.Fields("seq_no")
            .qtyDaily = adoRs.Fields("qty")
            .UnitCls = adoRs.Fields("unit_cls")
            .qtyAllResult = adoRs.Fields("qty_result")
            .completeCls = Val(adoRs.Fields("complete_cls") & "")
        End With
        Unload Me
    Else
        LblErrMsg.Caption = DisplayMsg("0013")
        txtBarcode.Text = ""
    End If
End Sub

Private Sub txtBarcode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyReturn: ShowData
    Case vbKeyEscape: Unload Me
    End Select
End Sub
