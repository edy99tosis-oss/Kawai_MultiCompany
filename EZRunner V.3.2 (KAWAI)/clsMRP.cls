VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim sql As String

Function CompanyName() As String
    Dim rsCompany As New ADODB.Recordset
    sql = "Select Company_Name From Company_Profile"
    Set rsCompany = Db.Execute(sql)
    If rsCompany.EOF Then CompanyName = "" Else CompanyName = Trim(rsCompany!company_name)
End Function

Function blnAkhir() As String
    Dim rsbln As New ADODB.Recordset
    sql = "select isnull(max(inventory_month),0) as bln," & _
        "isnull(max(inventory_year),0) as thn " & _
        "from inventory_control where " & _
        "inventory_year = (select max(inventory_year) from Inventory_control)"
    Set rsbln = Db.Execute(sql)
    blnAkhir = rsbln("bln") & "," & rsbln("thn")
    Set rsbln = Nothing
End Function

Sub HapusRequirement(nmDb As Connection, thn As String, bln As String)
    '********** Hapus Requirement Berd Lot No yg ada di Prod Planning pd date yg dipilih **********
    sql = "delete Requirement_Master where " & _
        "ChildRequirement_Year = '" & thn & _
        "' and ChildRequirement_Month = '" & bln & "'"
    nmDb.Execute sql
    
    sql = "delete Requirement " & _
        "Where Year(Production_Date) = '" & thn & _
        "' and Month(Production_Date) = '" & bln & _
        "' And (Complete_Cls is NULL or Complete_Cls = 0)"
    nmDb.Execute sql
    
    sql = "Update Requirement " & _
        "set Qty = 0, " & _
        "Off_Qty = 0, " & _
        "ChildRequirement_Qty = 0, " & _
        "OffChildRequirement_Qty = 0, " & _
        "ChildRequirementResult_Qty = 0, " & _
        "Last_Update = getdate(), " & _
        "Last_User = '" & userLogin & "' " & _
        "Where Year(Production_Date) = " & thn & " " & _
        "And Month(Production_Date) = " & bln
        nmDb.Execute sql
    '********************
End Sub

Sub HapusRequirementKosong(nmDb As Connection, thn As String, bln As String)
    sql = "Delete Requirement " & _
        "Where Qty = 0 " & _
        "And Off_Qty = 0 " & _
        "And ChildRequirement_Qty = 0 " & _
        "And OffChildRequirement_Qty = 0 " & _
        "And ChildRequirementResult_Qty = 0 " & _
        "And Year(Production_Date) = " & thn & _
        "And Month(Production_Date) = " & bln
    nmDb.Execute sql
End Sub

Sub UpdateRequirement(nmDb As Connection, thn As String, bln As String)
    '******** Update RequirementResult ***********
    sql = "update Requirement set Last_Update = getdate(), Last_User = '" & userLogin & "'," & _
        "ChildRequirementResult_Qty = " & _
        "(select isnull(sum(childRequirement_Qty),0) from Part_Supply " & _
        "where ParentItem_Code = Requirement.ParentITem_Code " & _
        "and Production_Date = Requirement.Production_Date " & _
        "and Lot_No = Requirement.Lot_No " & _
        "and ChildItem_Code = Requirement.ChildItem_Code " & _
        "And Supply_cls = 'S') " & _
        "Where Year(Production_Date) = '" & thn & _
        "' And Month(Production_Date) = '" & bln & "'"
    nmDb.Execute sql
    '*******************
    
    sql = "update Requirement_Master set Last_Update = getdate(), Last_User = '" & userLogin & "'," & _
        "ChildRequirementResult_Qty = " & _
        "(select isnull(sum(ChildRequirementResult_Qty),0) from Requirement " & _
        "where ChildItem_Code = Requirement_Master.ChildItem_Code " & _
        "and Year(Production_Date) =Requirement_Master.ChildRequirement_Year " & _
        "and Month(Production_Date) = Requirement_Master.ChildRequirement_Month) " & _
        "Where ChildRequirement_Year = " & CInt(thn) & _
        " And ChildRequirement_Month = " & CInt(bln)
    nmDb.Execute sql
    '*******************
End Sub

Sub UpdateRequirementResult(nmDb As Connection, _
    tglProd As String, ParentItemCD As String, lotno As String, ChildItemCD As String, Optional kondisi As String)
    
    '******** Update RequirementResult ***********
    sql = "Update Requirement set Last_Update = getdate(), Last_User = '" & userLogin & "'," & _
        "ChildRequirementResult_Qty = " & _
        "(Select ISNULL(Sum(ChildRequirement_Qty),0) From Part_Supply " & _
        "Where Supply_Cls = 'S' " & _
        "And Production_Date = Requirement.Production_Date " & _
        "And ParentItem_Code = Requirement.ParentItem_Code " & _
        "And Lot_No = Requirement.Lot_No " & _
        "And ChildItem_Code = Requirement.ChildItem_Code) " & _
        "Where "
    
    If kondisi = "" Then
        kondisi = "Production_Date = '" & tglProd & _
            "' And ParentItem_Code = " & ParentItemCD & _
            " And Lot_No = '" & lotno & _
            "' And ChildItem_Code in ( " & ChildItemCD & ")"
    End If
    
    sql = sql & kondisi
    nmDb.Execute sql
End Sub

Sub UpdateCompleteReq(nmDb As Connection, itemCD As String, lotno As String, tglProd As String)
    Dim rsItemReq As New ADODB.Recordset
    Dim strItemReq As String
    Dim keupdate As Long
    
    sql = "Update Requirement set Complete_Cls = 1, Last_Update = getdate(), Last_User = '" & userLogin & "' " & _
        "Where Qty <= " & _
        "(Select Sum(Qty) From Part_Receipt " & _
        "Where Item_Code = Requirement.ParentItem_Code " & _
        "And Lot_No = Requirement.Lot_No " & _
        "And DailySeq_No in " & _
        "(Select Seq_No From Daily_Production " & _
        "Where Item_Code = Part_Receipt.Item_Code " & _
        "And Lot_No = Part_Receipt.SuratJalan_No " & _
        "And Schedule_Date = Requirement.Production_Date)) " & _
        "And ParentItem_Code = '" & itemCD & _
        "' And Lot_No = '" & lotno & _
        "' And Production_Date = '" & tglProd & "'"
    nmDb.Execute sql, keupdate
    
    If keupdate > 0 Then
        sql = "Select ChildItem_Code From Requirement " & _
            "Where ParentItem_Code = '" & itemCD & _
            "' And Lot_No = '" & lotno & _
            "' ANd PRoduction_Date = '" & tglProd & "'"
        Set rsItemReq = nmDb.Execute(sql)
        If Not rsItemReq.EOF Then
            strItemReq = ""
            Do While Not rsItemReq.EOF
                strItemReq = strItemReq & "'" & rsItemReq!childitem_code & "',"
                rsItemReq.MoveNext
            Loop
            Call UpdateRequirementResultMaster(nmDb, CInt(Mid(tglProd, 6, 2)), CInt(Left(tglProd, 4)), Left(strItemReq, Len(strItemReq) - 1))
        End If
    End If
End Sub

Sub InputRequirementMaster(nmDb As Connection, thn As String, bln As String)
    sql = "insert into requirement_master " & _
        "(ChildRequirement_Year,ChildRequirement_Month, ChildItem_Code,ChildRequirement_Qty,ChildRequirementResult_Qty,ChildUnit_Cls, Last_Update, Last_User) " & _
        "select year(Production_Date),month(Production_Date),ChildItem_Code," & _
        "Isnull(sum(childrequirement_Qty),0)-Isnull(sum(Offchildrequirement_Qty),0) ," & _
        "Isnull(sum(childrequirementResult_Qty),0),ChildUnit_Cls, getdate(), '" & userLogin & "' " & _
        "From Requirement where " & _
        "year(Production_Date) = '" & thn & _
        "' And month(Production_Date) = '" & bln & _
        "' And (Complete_Cls = 0 Or Complete_Cls is NULL) " & _
        "group by ChildItem_Code,year(Production_Date),month(Production_Date),ChildUnit_Cls"
    nmDb.Execute sql
End Sub

Sub UpdateRequirementResultMaster(nmDb As Connection, bln As Integer, thn As Integer, ChildItemCD As String)
    '******** Update RequirementResult Master***********
    sql = "update Requirement_Master set Last_Update = getdate(), Last_User = '" & userLogin & "'," & _
        "ChildRequirement_Qty = " & _
        "(select isnull(sum(ChildRequirement_Qty),0) - isnull(sum(OffChildRequirement_Qty),0) " & _
        "from Requirement " & _
        "where ChildItem_Code = Requirement_Master.ChildItem_Code " & _
        "and Year(Production_Date) = Requirement_Master.ChildRequirement_Year " & _
        "and Month(Production_Date) = Requirement_Master.ChildRequirement_Month " & _
        "and (Complete_Cls = 0 Or Complete_Cls is NULL)) " & _
        ",ChildRequirementResult_Qty = " & _
        "(select isnull(sum(ChildRequirementResult_Qty),0) from Requirement " & _
        "where ChildItem_Code = Requirement_Master.ChildItem_Code " & _
        "and Year(Production_Date) = Requirement_Master.ChildRequirement_Year " & _
        "and Month(Production_Date) = Requirement_Master.ChildRequirement_Month " & _
        "and (Complete_Cls = 0 Or Complete_Cls is NULL)) " & _
        "Where ChildRequirement_Year = " & thn & _
        " And ChildRequirement_Month = " & bln & _
        " And ChildItem_Code in (" & ChildItemCD & ")"
    nmDb.Execute sql
    '*******************
End Sub

Sub updateStock(fromWHCode As String, itemAnak As String, qtyAnak As Double, lotno As String, tglDO As String, blnFix As Integer, thnFix As Integer, nmDb As Connection, stData As String, Optional qtySblmnya As Double, Optional stTambah As Integer)
    Dim blnSJ As Integer, thnSJ As Integer, batasBln As Integer
    Dim kondisi As String, updateStock As String, keupdate As Long
    Dim rsCek As New ADODB.Recordset
    
    'Utk update ke Stock harus lebih besar = dr Fix en Lebih besar 2 bln
    blnSJ = CInt(Mid(tglDO, 6, 2))
    thnSJ = CInt(Left(tglDO, 4))
    
    If (thnSJ = thnFix) Or (thnSJ - thnFix = 1) Then
        If thnSJ > thnFix Then batasBln = blnSJ + 12 Else batasBln = blnSJ
    
        If batasBln >= blnFix And (batasBln - blnFix) < 3 Then
            kondisi = "Warehouse_Code = '" & fromWHCode & "' and Item_Code ='" & itemAnak & "'"
    
            If batasBln = blnFix Then
                updateStock = "LM_" & stData
            ElseIf batasBln - blnFix = 1 Then
                updateStock = "TM_" & stData
            ElseIf batasBln - blnFix = 2 Then
                updateStock = "NM_" & stData
            End If
    
            sql = "update stock_master set Last_Update = getdate(), Last_User = '" & userLogin & "'," & _
                updateStock & "= (" & updateStock & "-(" & qtySblmnya & "))" & _
                IIf((stTambah = 1), "+", "-") & "(" & CDbl(qtyAnak) & _
                ") where " & kondisi
            nmDb.Execute sql, keupdate
    
            If keupdate = 0 Then
                sql = "insert into Stock_Master(Warehouse_Code,Item_Code,Last_Update, Last_User, " & updateStock & ") " & _
                    "values ('" & fromWHCode & "','" & itemAnak & "', getdate(), '" & userLogin & "'," & IIf((stTambah = 1), "+", "-") & CDbl(qtyAnak) & ")"
                nmDb.Execute sql
            End If
    
            If updateStock = "LM_" & stData Then
                sql = "Update Stock_Master set Last_Update = getdate(), Last_User = '" & userLogin & "'," & _
                    "LM_Current = LM_PreMonth + LM_Receipt - LM_Supply - LM_LossReject " & _
                    "where " & kondisi
                nmDb.Execute sql
            ElseIf updateStock = "TM_" & stData Then
                sql = "Update Stock_Master set Last_Update = getdate(), Last_User = '" & userLogin & "'," & _
                    "TM_Current = TM_Premonth + TM_Receipt - TM_Supply - TM_LossReject, " & _
                    "NM_PreMonth = TM_Premonth + TM_Receipt - TM_Supply - TM_LossReject, " & _
                    "NM_Current = (TM_Premonth + TM_Receipt - TM_Supply - TM_LossReject) + " & _
                    "NM_Receipt - NM_Supply - NM_LossReject " & _
                    "where " & kondisi
                nmDb.Execute sql
            ElseIf updateStock = "NM_" & stData Then
                sql = "Update Stock_Master set Last_Update = getdate(), Last_User = '" & userLogin & "'," & _
                    "NM_Current = TM_Current + NM_Receipt - NM_Supply - NM_LossReject " & _
                    "where " & kondisi
                nmDb.Execute sql
            End If
    
            sql = "Delete stock_master Where " & kondisi & _
                " And LM_Premonth=0 and TM_Premonth=0 and NM_premonth=0 " & _
                "And LM_Receipt=0 and TM_Receipt=0 and NM_receipt=0 " & _
                "And LM_Supply=0 and TM_Supply=0 and NM_supply=0 " & _
                "And LM_LossReject=0 and TM_LossReject=0 and NM_LossReject=0 " & _
                "And LM_Current=0 and TM_Current=0 and NM_Current=0 " & _
                "And (LM_Inventory=0 or LM_Inventory is NULL) " & _
                "And (TM_Inventory = 0 or TM_Inventory is NULL) " & _
                "And (NM_Inventory = 0  or NM_Inventory is NULL) "
            nmDb.Execute sql
        
        End If
    End If
End Sub

Public Sub inputSupply(nmDb As Connection, seqNo As Double, fromWHCode As String, machineNo As String, toWHCode As String, _
    childSupplyDt As String, ChildItemCD As String, ChildLotNo As String, _
    SupplyCls As String, Stockqty As Double, consumpQty, UnitCls As String, ParentItemCD As String, LotNoParent As String, _
    KeyProd As String, Optional tglProd As String, Optional matConsumpCls As String)

    Dim keupdate As Long
    
    sql = "Update Part_Supply " & _
        "Set Last_Update = getdate(), Last_User = '" & userLogin & "'," & _
        "ChildRequirement_Qty = " & Stockqty & ", " & _
        "Consumption_Qty =  " & consumpQty & ", " & _
        "Remarks =  '" & ChildLotNo & "', " & _
        "ChildUnit_Cls = '" & UnitCls & "', " & _
        "ParentItem_Code = '" & ParentItemCD & "', " & _
        "Lot_No = '" & LotNoParent & "', " & _
        "Production_Date = " & IIf(tglProd = "Null", "Null", "'" & tglProd & "'") & ", " & _
        "MaterialConsump_Cls = " & IIf(matConsumpCls = "", "Null", "'" & matConsumpCls & "'") & " " & _
        "Where Seq_No = " & CStr(seqNo) & _
        " And ChildItem_Code='" & ChildItemCD & "'"
    nmDb.Execute sql, keupdate
    
    If keupdate = 0 Then
        sql = "insert into Part_Supply(FromWarehouse_Code,From_Address,ToWarehouse_Code," & _
            "ChildSupply_date,ChildItem_Code,Supply_Cls,ChildRequirement_Qty,Consumption_Qty, ChildUnit_Cls," & _
            "ParentItem_Code,Lot_No,Production_Date,Do_No,Remarks,MaterialConsump_Cls, Last_Update, Last_User) " & _
            "Values ('" & fromWHCode & "','" & machineNo & "','" & toWHCode & "','" & _
            childSupplyDt & "','" & ChildItemCD & "','" & _
            SupplyCls & "'," & Stockqty & ",'" & consumpQty & "', '" & UnitCls & "','" & _
            ParentItemCD & "','" & LotNoParent & "'," & _
            IIf(tglProd = "Null", "Null", "'" & tglProd & "'") & ",'" & KeyProd & "','" & ChildLotNo & "'," & _
            IIf(matConsumpCls = "", "Null", "'" & matConsumpCls & "'") & ",getdate(),'" & userLogin & "')"
        nmDb.Execute sql
    End If
End Sub

Function keyReceipt() As Double
    Dim rsSeq As New ADODB.Recordset
    sql = "select Isnull(Max(Seq_No),0) + 1  as seqNo  from part_Receipt"
    Set rsSeq = Db.Execute(sql)
    keyReceipt = rsSeq!seqNo
End Function

Function keySupply() As Double
    Dim rsSeq As New ADODB.Recordset
    sql = "select Isnull(Max(Seq_No),0) as seqNo  from part_Supply"
    Set rsSeq = Db.Execute(sql)
    keySupply = rsSeq!seqNo
End Function

Function formatnilai(nilai As Double) As String
    Dim strNilai As String
    If nilai <> 0 Then
        If Round(CDec(nilai)) / CDec(nilai) <> 1 Then
            formatnilai = Format(nilai, "#,##0.####")
        Else
            formatnilai = Format(nilai, "#,##0")
        End If
    Else
        formatnilai = 0
    End If
End Function


Public Sub HapusDataSupp(nmDb As Connection, noDO As String, blnFix As Integer, thnFix As Integer, Optional bySeqNo As Byte, Optional updateAllocate As Byte)
    Dim dbSupp As New Connection
    Dim rsDelSupp As New ADODB.Recordset
    Dim rsPlanActual As New ADODB.Recordset
    Dim qtyPlan As Double, qtyAllocate As Double
    
    dbSupp.Open Db.ConnectionString
    
    sql = "Select Supp.*,Item.StockControl_Cls stockItem, " & vbCrLf & _
        "Isnull(WH.StockControl_Cls,'01') StockWH " & vbCrLf & _
        "from Part_Supply Supp  " & vbCrLf & _
        "   Left Join Item_Master Item on Supp.ChildItem_Code = Item.Item_Code " & vbCrLf & _
        "   Left Join Warehouse_Master WH on FromWarehouse_Code = WH.WH_Code" & vbCrLf & _
        "where 'A'='A' --Supp.ChildItem_Code = Item.Item_Code " & vbCrLf & _
        "--And FromWarehouse_Code *= WH.WH_Code " & vbCrLf

    If bySeqNo = 1 Then 'dr Seq No saja
        sql = sql & "And Seq_No in (" & noDO & ")"
    Else
        sql = sql & "And DO_NO in (" & noDO & ")"
    End If

    Set rsDelSupp = dbSupp.Execute(sql)
'
    With rsDelSupp
        Do While Not .EOF
            If !stockItem = "01" And !stockWH = "01" Then
                Call updateStock(!fromwarehouse_code, _
                !childitem_code, !ChildRequirement_qty, "", Format(!childsupply_date, "yyyy-MM-dd"), _
                blnFix, thnFix, nmDb, "Supply", 0, 0)
            End If
            .MoveNext
        Loop
    End With
    
    sql = "delete  Part_Supply where "
    If bySeqNo = 1 Then 'dr Seq No saja
        sql = sql & "Seq_No in (" & noDO & ")"
    Else
        sql = sql & "DO_NO in (" & noDO & ")"
    End If
    nmDb.Execute sql
End Sub
