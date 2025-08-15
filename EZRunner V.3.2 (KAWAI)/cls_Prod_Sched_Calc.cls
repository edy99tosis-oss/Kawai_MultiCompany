VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_Prod_Sched_Calc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim lcd_startDate As Date
Dim lcd_endDate As Date

Private Type LT_paRam
    custCode As String
    itemCode As String
    orderSeq As Integer
    poNo As String
    Lot As Double
    prodCls As String
    deldate As Date
    Qty As Double
    Lead As Integer
End Type

Dim T_item() As LT_paRam    'dibuat array utk support calc multi item / per PoNo

Public Property Get StartDate() As Date
    StartDate = lcd_startDate
End Property

Public Property Let StartDate(ByVal vNewValue As Date)
    lcd_startDate = vNewValue
End Property

Public Property Get EndDate() As Date
    EndDate = lcd_endDate
End Property

Public Property Let EndDate(ByVal vNewValue As Date)
    lcd_endDate = vNewValue
End Property

Public Function CF_processItem(ByVal pCust As String, ByVal pItem As String, ByVal pOrderSeq As Long, ByVal pPo As String, _
                                                Optional pProdCls As String = "", Optional pLot As Double = -1, Optional pDelDate As Date = 0, _
                                                Optional pQty As Double = -1, Optional pLead As Integer = -1) As String
    ReDim T_item(0) As LT_paRam
    Dim ls_msG As String
    
    If pProdCls = "" Or pLot = -1 Or Format(pDelDate, "yyyymmdd") = Format(0, "yyyymmdd") Or pQty = -1 Then
        'getItemInformation
        CF_processItem = "gada data"
        Exit Function
    Else
        T_item(0).custCode = pCust
        T_item(0).itemCode = pItem
        T_item(0).orderSeq = pOrderSeq
        T_item(0).poNo = pPo
        T_item(0).Lot = pLot
        T_item(0).prodCls = pProdCls
        T_item(0).deldate = pDelDate
        T_item(0).Qty = pQty
        T_item(0).Lead = pLead
    End If

    ls_msG = CF_insertMaster(T_item(0).custCode, T_item(0).itemCode, T_item(0).orderSeq, T_item(0).poNo, _
                                        T_item(0).Qty, T_item(0).deldate, T_item(0).Lot, T_item(0).Lead, IIf(T_item(0).prodCls = "01", True, False), 0) & ""
    If ls_msG <> "" Then
        CF_processItem = ls_msG
        Exit Function
    End If
    
    'update calc_CLS di orderentry_detail
    ls_msG = CF_updateCalcAtOrderEntryDetail(T_item(0).custCode, T_item(0).poNo, T_item(0).itemCode, T_item(0).orderSeq) & ""
    If ls_msG <> "" Then
        CF_processItem = ls_msG
        Exit Function
    End If
    
End Function

Private Function CF_insertMaster(ByVal pCust As String, ByVal pItem As String, ByVal pOrderSeq As Long, ByVal pPo As String, _
                                                ByVal pQty As Double, ByVal pDelDate As Date, ByVal pLot As Double, ByVal pLead As Integer, _
                                                ByVal pInsertParent As Boolean, ByVal pBomLevel As Integer) As String
Dim ls_Que As String
Dim ls_msG As String

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'1. Insert to calc_Master
    ls_Que = "DELETE [ProductionCalculate_Master] " & vbCrLf & _
            " where Cust_Code = '" & pCust & "' " & vbCrLf & _
            " and po_no = '" & pPo & "' " & vbCrLf & _
            " and planitem_code = '" & pItem & "' " & vbCrLf & _
            " and seq_no = '" & pOrderSeq & "' "
    Db.Execute ls_Que
    
    ls_Que = " INSERT INTO [ProductionCalculate_Master] " & vbCrLf & _
                "            ([Cust_Code] " & vbCrLf & _
                "            ,[PO_No] " & vbCrLf & _
                "            ,[Seq_No] " & vbCrLf & _
                "            ,[PlanItem_Code] " & vbCrLf & _
                "            ,[Last_Update] " & vbCrLf & _
                "            ,[Last_User] " & vbCrLf & _
                "            ,[Register_Date]) " & vbCrLf & _
                "      VALUES " & vbCrLf & _
                "            ('" & pCust & "' " & vbCrLf & _
                "            ,'" & pPo & "' " & vbCrLf & _
                "            ,'" & pOrderSeq & "' " & vbCrLf & _
                "            ,'" & pItem & "' " & vbCrLf & _
                "            ,null " & vbCrLf & _
                "            ,'" & userLogin & "' " & vbCrLf & _
                "            ,getdate() ) "
    
    Db.Execute (ls_Que)

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'2. Calculate Data -> masukin data ke table calc_detail berdasarkan BOM
'    ls_msG = CF_insertParent(pCust, pItem, pOrderSeq, pPo, pQty, pDelDate, pLot, pLead, pInsertParent, pBomLevel) & ""

    Dim rsParent As New ADODB.Recordset
    ls_msG = CF_insertParent(pCust, pItem, pOrderSeq, pPo, pQty, pDelDate, pLot, pLead, pInsertParent, pBomLevel) & ""
    If ls_msG <> "" Then
        CF_insertMaster = ls_msG & " - bomLv: " & pBomLevel
        Exit Function
    End If
    
'    ls_Que = " declare @tanggal char(6) " & vbCrLf & _
'                " set @tanggal = '" & Format(pDelDate, "yyyymmdd") & "' " & vbCrLf & _
'                "  " & vbCrLf & _
'                " select bom.item_code child_itemcode, bom.qty bom_Qty, " & vbCrLf & _
'                "       im.production_Cls, isnull(im.product_readtime,0) Lead_Time, isnull(im.lot_Qty,0) Lot_Qty " & vbCrLf & _
'                " from bom_master bom " & vbCrLf & _
'                " inner join item_Master im " & vbCrLf & _
'                "   on bom.item_Code = im.item_Code " & vbCrLf & _
'                " where parent_itemcode = '" & pItem & "' " & vbCrLf & _
'                " and @tanggal >= start_Date and @tanggal<= end_date "
'    Set rsParent = Db.Execute(ls_Que)
'    If Not (rsParent.EOF Or rsParent.BOF) Then
'        While Not rsParent.EOF
'            ls_msG = CF_insertParent(pCust, Trim$(rsParent("child_itemCode") & ""), pOrderSeq, pPo, pQty, pDelDate, rsParent("bom_Qty"), rsParent("Lead_Time"), IIf(rsParent("Production_Cls") = "01", True, False), pBomLevel) & ""
'            If ls_msG <> "" Then
'                CF_insertMaster = ls_msG & " - bomLv: " & pBomLevel
'                Exit Function
'            End If
'            rsParent.MoveNext
'        Wend
'    Else
'    End If
    
    CF_insertMaster = ls_msG
    If rsParent.State = adStateOpen Then rsParent.Close
    Set rsParent = Nothing
End Function

Private Function CF_insertParent(ByVal pCust As String, ByVal pItem As String, ByVal pOrderSeq As Long, ByVal pPo As String, _
                                                ByVal pQty As Double, ByVal pDelDate As Date, ByVal pLot As Double, ByVal pLead As Integer, _
                                                ByVal pInsertParent As Boolean, ByVal pBomLevel As Integer) As String

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'   Plan_Cls :  1=Delivery             2=Production Schedule
'                   3=Consumption       4=Purchase
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Dim ls_PlanCls As String
Dim ld_Qty As Double
Dim ld_Date As Date

Dim ls_msG As String

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'1. Insert Delivery Data into calc_Detail
   
'    If pLot <> 0 Then
'        ld_Qty = RoundUp(pQty / pLot) * pLot
'    Else
        ld_Qty = pQty
'    End If
    
    If pBomLevel = 0 Then
        ls_PlanCls = 1  'Plan Delivery
    Else
        ls_PlanCls = 3 'Plan Consumption
    End If
    
    ls_msG = CF_processInsertDetail(pCust, pPo, pOrderSeq, pItem, ls_PlanCls, pDelDate, ld_Qty)
    If ls_msG <> "" Then
        CF_insertParent = ls_msG & " - bomLv: " & pBomLevel
        Exit Function
    End If
    
    
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'2. Insert prod sched or purchase into calc_Detail
    
    ls_msG = CF_CalcDate(pDelDate, pLead, ld_Date) & ""
    If ls_msG <> "" Then
        CF_insertParent = ls_msG & " - bomLv: " & pBomLevel
        Exit Function
    End If

    If pInsertParent Then
        ls_PlanCls = 2  'insert Prod sched
    Else
        ls_PlanCls = 4 'insert Purchase
    End If

    ls_msG = CF_processInsertDetail(pCust, pPo, pOrderSeq, pItem, ls_PlanCls, ld_Date, ld_Qty)
    If ls_msG <> "" Then
        CF_insertParent = ls_msG & " - bomLv: " & pBomLevel
        Exit Function
    End If

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'3. Check child
    Dim ls_Que As String
    Dim rsChild As New ADODB.Recordset

    ls_Que = " declare @tanggal char(6) " & vbCrLf & _
                " set @tanggal = '" & Format(ld_Date, "yyyymmdd") & "' " & vbCrLf & _
                "  " & vbCrLf & _
                " select bom.item_code child_itemcode, bom.qty bom_Qty, " & vbCrLf & _
                "       im.production_Cls, isnull(im.product_readtime,0) Lead_Time, isnull(im.lot_Qty,0) Lot_Qty " & vbCrLf & _
                " from bom_master bom " & vbCrLf & _
                " inner join item_Master im " & vbCrLf & _
                "   on bom.item_Code = im.item_Code " & vbCrLf & _
                " where parent_itemcode = '" & pItem & "' " & vbCrLf & _
                " and @tanggal >= start_Date and @tanggal<= end_date "
    Set rsChild = Db.Execute(ls_Que)
    If Not (rsChild.EOF Or rsChild.BOF) Then
        While Not rsChild.EOF
            ls_msG = CF_insertParent(pCust, Trim$(rsChild("child_itemCode") & ""), pOrderSeq, pPo, ld_Qty, ld_Date, rsChild("bom_Qty"), rsChild("Lead_Time"), IIf(rsChild("Production_Cls") = "01", True, False), pBomLevel + 1)
            If ls_msG <> "" Then
                CF_insertParent = ls_msG & " - bomLv: " & pBomLevel + 1
                Exit Function
            End If
            rsChild.MoveNext
        Wend
    Else
    End If

    If rsChild.State = adStateOpen Then rsChild.Close
    Set rsChild = Nothing
End Function

Private Function CF_CalcDate(ByVal pDelDate As Date, ByVal pLeadTime As Integer, ByRef pNewDate As Date) As String
Dim li_rowAffected As Integer
Dim rsDate As New ADODB.Recordset
Dim ls_Que As String
Dim li_u As Integer
    
On Local Error GoTo errHandler

    pNewDate = pDelDate
    If pLeadTime = 0 Then GoTo normalExit
    For li_u = 1 To pLeadTime
        li_rowAffected = 1
        While li_rowAffected = 1
            pNewDate = DateAdd("d", -1, pNewDate)
            ls_Que = "select * from calendar_master where cal_date = '" & Format(pNewDate, "yyyy-mm-dd") & "' and factory_code = '999' "
            Set rsDate = Db.Execute(ls_Que)
            li_rowAffected = IIf(rsDate.BOF Or rsDate.EOF, 0, 1)
        Wend
    Next li_u
    
normalExit:
If rsDate.State = adStateOpen Then rsDate.Close
Set rsDate = Nothing
Exit Function
errHandler:
CF_CalcDate = Err.Description
Resume normalExit
End Function

Private Function CF_generateSeq(ByVal pCust As String, ByVal pPo As String, ByVal pOrderSeq As Long, ByVal pItem As String, ByVal pPlanCls As String, ByRef li_SeqNo) As String
Dim li_rowAffected As Integer
Dim rsSeq As New ADODB.Recordset
Dim ls_Que As String
    
On Local Error GoTo errHandler

    ls_Que = "select new_Seq = isnull(max(Plan_SeqNo),0) + 1  from ProductionCalculate_Detail " & vbCrLf & _
                " where cust_Code = '" & pCust & "' " & vbCrLf & _
                "   and po_no = '" & pPo & "' " & vbCrLf & _
                "   and seq_no = '" & pOrderSeq & "' " & vbCrLf & _
                "   /*and plan_cls = '" & pPlanCls & "'*/ " & vbCrLf & _
                "   and planitem_code = '" & pItem & "' "
    Set rsSeq = Db.Execute(ls_Que)
            
    li_SeqNo = rsSeq("new_seq")
    
normalExit:
If rsSeq.State = adStateOpen Then rsSeq.Close
Set rsSeq = Nothing
Exit Function
errHandler:
CF_generateSeq = Err.Description
Resume normalExit
End Function

Private Function CF_deleteDetail(ByVal pCust As String, ByVal pPo As String, ByVal pOrderSeq As Long, ByVal pItem As String, ByVal pPlanCls As String) As String
On Local Error GoTo errHandler
Dim ls_Que As String

    ls_Que = "DELETE [ProductionCalculate_Detail] " & vbCrLf & _
            " where Cust_Code = '" & pCust & "' " & vbCrLf & _
            " and po_no = '" & pPo & "' " & vbCrLf & _
            " and planitem_code = '" & pItem & "' " & vbCrLf & _
            " and plan_cls = '" & pPlanCls & "' " & vbCrLf & _
            " and seq_no = '" & pOrderSeq & "' "
    Db.Execute ls_Que
    
normalExit:
    Exit Function
errHandler:
    CF_deleteDetail = Err.Description
    Resume normalExit
End Function

Private Function CF_insertDetail(ByVal pCust As String, ByVal pPo As String, ByVal pOrderSeq As Long, ByVal pItem As String, _
                                                ByVal pPlanSeq As Long, ByVal pPlanCls As String, ByVal pPlanDate As Date, ByVal pQty As Double) As String
On Local Error GoTo errHandler
Dim ls_Que As String

    ls_Que = " INSERT INTO [ProductionCalculate_Detail] " & vbCrLf & _
                "            ([Cust_Code] " & vbCrLf & _
                "            ,[PO_No] " & vbCrLf & _
                "            ,[Seq_No] " & vbCrLf & _
                "            ,[PlanItem_Code] " & vbCrLf & _
                "            ,[Plan_SeqNo] " & vbCrLf & _
                "            ,[Plan_Cls] " & vbCrLf & _
                "            ,[Plan_Date] " & vbCrLf & _
                "            ,[Plan_Qty] " & vbCrLf & _
                "            ,[Last_Update] " & vbCrLf & _
                "            ,[Last_User] " & vbCrLf & _
                "            ,[Register_Date]) " & vbCrLf & _
                "      VALUES " & vbCrLf & _
                "            ('" & pCust & "' " & vbCrLf & _
                "            ,'" & pPo & "' " & vbCrLf & _
                "            ,'" & pOrderSeq & "' " & vbCrLf & _
                "            ,'" & pItem & "' " & vbCrLf & _
                "            ,'" & pPlanSeq & "' " & vbCrLf & _
                "            ,'" & pPlanCls & "' " & vbCrLf & _
                "            ,'" & Format(pPlanDate, "yyyy-mm-dd") & "' " & vbCrLf & _
                "            ,'" & pQty & "' " & vbCrLf & _
                "            ,null " & vbCrLf & _
                "            ,'" & userLogin & "' " & vbCrLf & _
                "            ,getdate() ) "
    Db.Execute ls_Que
    
normalExit:
    Exit Function
errHandler:
    CF_insertDetail = Err.Description
    Resume normalExit
End Function

Private Function CF_processInsertDetail(ByVal pCust As String, ByVal pPo As String, ByVal pOrderSeq As Long, ByVal pItem As String, _
                                                            ByVal pPlanCls As String, ByVal pPlanDate As Date, ByVal pQty As Double) As String
On Local Error GoTo errHandler
Dim li_SeqNo As Integer
Dim ls_msG As String

    ls_msG = CF_deleteDetail(pCust, pPo, pOrderSeq, pItem, pPlanCls) & ""
    If ls_msG <> "" Then
        CF_processInsertDetail = ls_msG
        Exit Function
    End If
    
    ls_msG = CF_generateSeq(pCust, pPo, pOrderSeq, pItem, pPlanCls, li_SeqNo) & ""
    If ls_msG <> "" Then
        CF_processInsertDetail = ls_msG
        Exit Function
    End If
    
    ls_msG = CF_insertDetail(pCust, pPo, pOrderSeq, pItem, li_SeqNo, pPlanCls, pPlanDate, pQty) & ""
    If ls_msG <> "" Then
        CF_processInsertDetail = ls_msG
        Exit Function
    End If
    
normalExit:
    Exit Function
errHandler:
    CF_processInsertDetail = Err.Description
    Resume normalExit
End Function

Private Function CF_updateCalcAtOrderEntryDetail(ByVal pCust As String, ByVal pPo As String, ByVal pItem As String, ByVal pOrderSeq As Long) As String
Dim ls_Que As String
    ls_Que = "update orderentry_Detail " & vbCrLf & _
    " set calculate_Cls = '1' " & vbCrLf & _
    " where Cust_Code = '" & pCust & "' " & vbCrLf & _
    "   and Po_No  = '" & pPo & "' " & vbCrLf & _
    "   and item_Code  = '" & pItem & "' " & vbCrLf & _
    "   and Seq_No = '" & pOrderSeq & "' "
    Db.Execute ls_Que
End Function
