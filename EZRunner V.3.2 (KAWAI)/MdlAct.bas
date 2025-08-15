Attribute VB_Name = "MdlAct"
Option Explicit

Public gd_StockMaster As Double
Public gs_unitDesc  As String
Public gs_year As String
Public FromControlCls As String
Public ItemControlCls As String
Public ToControlCls As String
Public gs_accUS As String, gs_accYen As String, gs_accIDR As String, gs_bank As String
Public gi_countBaris As Integer
Public gs_cartoonList As String
Public li_diff As Integer

Public Function uf_Description(ls_Code As String, ls_tablename As String, ls_fieldname As String) As String
    Dim RS As New ADODB.Recordset

    sql = "select description from " & ls_tablename & " where " & ls_fieldname & "  = '" & Trim(ls_Code) & "'"
    If RS.State <> adStateClosed Then RS.Close
    RS.Open sql, Db, adOpenKeyset, adLockOptimistic
    If Not (RS.EOF And RS.BOF) Then
        uf_Description = Trim(RS!Description)
    End If
    If RS.State <> adStateClosed Then RS.Close
End Function
Public Sub up_FillCombo(nmCombo, ls_tablename As String, Optional ls_field As String, Optional ls_Condition As String, Optional lb_AllSelection As Boolean)
    
    Dim ls_sql As String
    Dim i As Long
    Dim lrs As New ADODB.Recordset
    
    With nmCombo
        .clear
        .columnCount = 2
        
        If lrs.State <> adStateClosed Then lrs.Close
        
        If Trim(ls_field) = "" Then
            ls_sql = "select * from " & ls_tablename
        Else
            ls_sql = "select " & ls_field & " from " & ls_tablename & " " & ls_Condition
        End If
        
        lrs.Open ls_sql, Db, adOpenKeyset, adLockOptimistic
        i = 0
        
        If lb_AllSelection = True Then
            .AddItem ""
            .List(i, 0) = strAll
            .List(i, 1) = strAll
            i = 1
        End If
        
        While lrs.EOF = False
            .AddItem ""
            .List(i, 0) = Trim(lrs(0))
            .List(i, 1) = Trim(lrs(1))
            lrs.MoveNext
            i = i + 1
        Wend
        
        If lrs.State <> adStateClosed Then lrs.Close
        
        .ListWidth = 60
        .ColumnWidths = "20 pt;40 pt"
    End With

End Sub

Public Function uf_GetSupplyLastUpdate(WHCode As String, ItemCode As String, lotno As String) As String
    
    Dim rsmax As New ADODB.Recordset
    Dim sqlMax As String
    
    '# Check Data Last Update from part_supply
    sqlMax = "select max(supplyDate)supplyDate  from ( " & _
        "select max(childsupply_Date)SupplyDate from part_Supply  where " & _
        "FromWarehouse_code='" & Trim(WHCode) & "' " & _
        "and Childitem_code='" & Trim(ItemCode) & "' " & _
        "and Lot_no='" & Trim(lotno) & "' " & _
        "union all " & _
        "select max(receipt_Date)SupplyDate from part_receipt  where " & _
        "Warehouse_code='" & Trim(WHCode) & "' " & _
        "and item_code='" & Trim(ItemCode) & "' " & _
        "and suratjalan_no='" & Trim(lotno) & "' )tb"
    
    If rsmax.State <> adStateClosed Then rsmax.Close
    rsmax.Open sqlMax, Db, adOpenKeyset, adLockOptimistic
    
    uf_GetSupplyLastUpdate = IIf(IsNull(rsmax!SupplyDate) = True, "", rsmax!SupplyDate)
    If rsmax.State <> adStateClosed Then rsmax.Close

End Function

Public Function uf_GenerateSupplyRequestNo(Month As String, Year As String) As String

    Dim RS As New ADODB.Recordset

    RS.Open " select * from partSupplyRequest_Master where " & _
        "( month(childsupply_date)='" & Trim(Month) & "' and  year(childsupply_date)='" & Trim(Year) & "' ) " & _
        "or right(rtrim(supplyRec_No),7)='" & Trim(Month) & "/" & Trim(Year) & "' " & _
        "order by left(supplyRec_no,5) desc ", Db

    If RS.EOF = False Then '# previous number exits
        RS.MoveFirst
        uf_GenerateSupplyRequestNo = Format(Val(Left(RS!supplyRec_No, 5) + 1), "00000") & "/" & Month & "/" & Year
    Else '# previous number not exits
        uf_GenerateSupplyRequestNo = "00001/" & Format(Month, "00") & "/" & Format(Year, "0000")
    End If
    RS.Close

End Function

Public Function uf_Ceiling(number As Double) As Double

    If InStr(1, CStr(number), ".") > 0 Then
        uf_Ceiling = CDbl(Left(CStr(number), InStr(1, CStr(number), ".") - 1)) + 1
    Else
        uf_Ceiling = number
    End If

End Function

Public Function uf_GetMaterialDescription(material_code As String) As String

    Dim RS As New ADODB.Recordset
    
    RS.Open "select * from material_cls where material_cls='" & Trim(material_code) & "'", Db
    If RS.EOF = False Then '# Material Exist
        uf_GetMaterialDescription = Trim(RS!Description)
    Else '# Material Exist
        uf_GetMaterialDescription = ""
    End If
    RS.Close

End Function

Public Function uf_GetQueryDescription(SearchInfo As String, fieldCode As String) As String

    Select Case Trim(SearchInfo)
    Case "ItemDesc": uf_GetQueryDescription = "case rtrim(" & Trim(fieldCode) & ") when '01' then  rtrim(im.item_name) + ' (T' + rtrim(im.thickness)  + '  x W' + rtrim(im.Width) + ' x L' + rtrim(im.length ) +')' else im.item_name end "
    End Select
    
End Function

Public Function uf_GetLastClosing(Request As String) As String
'
    '###############################################################
    '#                                                                                                                          #
    '#  Notes : To Get Last Closing Month,Year, or full date                                 #
    '#                                                                                                                          #
    '###############################################################
    
    Dim sqlControl As String, RsInvControl As New ADODB.Recordset
    Dim InvYear As String
    Dim InvMonth As String
    Dim lotno As String
    
    sqlControl = "select * from inventory_control where fix_cls='1' order by inventory_year desc ,inventory_month desc"
    
    If Request = "fulldate" Then
        sqlControl = "   select    " & _
        "        cast (   " & _
        "        cast(year as varchar(4) ) +case when month <10 then '0' else'' end +cast (month as varchar(2) )+'01'    " & _
        "            as dateTime)ClosingDate     " & _
        "        from    " & _
        "        (   " & _
        "        select top 1 max(inventory_month)month,inventory_year year   " & _
        "         from inventory_control    " & _
        "        where fix_cls='1'   " & _
        "        group by inventory_year   " & _
        "        order by inventory_year desc   " & _
        "        )tbA  "
    End If
    
    If RsInvControl.State <> adStateClosed Then RsInvControl.Close
    RsInvControl.Open sqlControl, Db, adOpenForwardOnly, adLockReadOnly
    
    If RsInvControl.EOF = False Then '#Inventory CLosing Data exist
        If Request <> "fulldate" Then
            RsInvControl.MoveFirst
            InvYear = Trim(RsInvControl!Inventory_Year)
            InvMonth = Trim(RsInvControl!Inventory_Month)
        End If
    End If
     
    If Request = "month" Then '#Request for month
        uf_GetLastClosing = InvMonth
    ElseIf Request = "year" Then '#Request for year
        uf_GetLastClosing = InvYear
    ElseIf Request = "fulldate" Then '#Request for fulldate
        uf_GetLastClosing = IIf(IsNull(RsInvControl!closingdate), 0, Format(RsInvControl!closingdate, "yyyy-MM-dd"))
    End If
    
    RsInvControl.Close

End Function

Public Sub up_UpdateStockMaster(dateDT As Date, DateInv As String, yearInv As String, wCode As String, toCode As String, iCode As String, Qty As Double, inStatus As String, toStatus As String, lotNoCode As String, Status As String, maxDateFrom As String, maxDateTo As String, skip As Boolean, Receipt As Boolean, Optional DBSp As Boolean, Optional dbCon As ADODB.Connection)

    Dim ls_sqlIns As String
    Dim ls_sqlTo As String
    Dim RSIns As New ADODB.Recordset
    Dim rsTo As New ADODB.Recordset
    
    '#From Warehouse
    ls_sqlIns = " select * from stock_master where " & _
        "warehouse_code='" & Trim(wCode) & "' " & _
        "and item_code='" & Trim(iCode) & "' "
        
    If RSIns.State <> adStateClosed Then RSIns.Close
    
    If DBSp = True Then
        RSIns.Open ls_sqlIns, dbCon, adOpenKeyset, adLockOptimistic
    Else
        RSIns.Open ls_sqlIns, Db, adOpenKeyset, adLockOptimistic
    End If
    
    If RSIns.EOF = True Then
    
        '# Insert data into stock Master for FromWarehouse if there is no present data
        Db.Execute ("insert into stock_master (Warehouse_Code, Item_Code, " & _
            "LM_PreMonth, LM_Receipt, LM_Supply, LM_LossReject, LM_Current, LM_Inventory, " & _
            "TM_PreMonth, TM_Receipt, TM_Supply, TM_LossReject, TM_Current, TM_Inventory, " & _
            "NM_PreMonth, NM_Receipt, NM_Supply, NM_LossReject, NM_Current, NM_Inventory, " & _
            "Last_Update, Last_User) " & _
            "values ('" & Trim(wCode) & "','" & Trim(iCode) & "',0,0,0,0,0,null,0,0,0,0,0,null,0,0,0,0,0,null,getdate(),'" & userLogin & "')")
        RSIns.Requery
        GoTo insDet
    
    Else
    
insDet:
    If skip = False Then
    
        '#To Warehouse
        ls_sqlTo = " select * from stock_master where " & _
            "warehouse_code='" & Trim(toCode) & "' " & _
            "and item_code='" & Trim(iCode) & "' "
    
        If rsTo.State <> adStateClosed Then rsTo.Close
            If DBSp = True Then
                rsTo.Open ls_sqlTo, dbCon, adOpenKeyset, adLockOptimistic
            Else
                rsTo.Open ls_sqlTo, Db, adOpenKeyset, adLockOptimistic
            End If
        End If
    
        If skip = False Then
            rsTo.Requery
            '# Insert data into stock Master for ToWarehouse if there is no present data
            If rsTo.EOF = True Then
                Db.Execute ("insert into stock_master (Warehouse_Code, Item_Code, " & _
                    "LM_PreMonth, LM_Receipt, LM_Supply, LM_LossReject, LM_Current, LM_Inventory, " & _
                    "TM_PreMonth, TM_Receipt, TM_Supply, TM_LossReject, TM_Current, TM_Inventory, " & _
                    "NM_PreMonth, NM_Receipt, NM_Supply, NM_LossReject, NM_Current, NM_Inventory, " & _
                    "Last_Update, Last_User) " & _
                    "values ('" & Trim(toCode) & "','" & Trim(iCode) & "',0,0,0,0,0,null,0,0,0,0,0,null,0,0,0,0,0,null,getdate(),'" & userLogin & "')")
            End If
            rsTo.Requery
        End If
    
        '###############################################################
        '#                                                             #
        '#  Notes : from receipt in the code below indicate that data  #
        '#          is used for Receipt Transaction                    #
        '#                                                             #
        '###############################################################
    
       Select Case up_GetDateRange(dateDT)
        
        Case 0:
    
            If Receipt = False Then '#From Supply
            
                '#Check if Supply influence stock in from warehouse or not
                If FromControlCls = "01" Then
                    RSIns!lm_current = RSIns!lm_current - CDbl(Trim(Qty))
                    RSIns!lm_supply = RSIns!lm_supply + IIf(Trim(inStatus) <> "L" And Trim(inStatus) <> "RJ", CDbl(Trim(Qty)), 0)
                    RSIns!lm_lossreject = RSIns!lm_lossreject + IIf(Trim(inStatus) = "L" Or Trim(inStatus) = "RJ", CDbl(Trim(Qty)), 0)
                    RSIns!Last_Update = Now
                    RSIns!last_user = userLogin
                    RSIns.update
                End If
                
                '#Skip influence To Warehouse Stock or Not
                If skip = False Then
                
                    '#If not Supply-Supply Transaction then dn't influence Stock in toWarehouse
                    If Trim(inStatus) <> "S1" Then Exit Sub
                    
                    '#Check if Supply influence stock in To Warehouse or not
                    If Trim(toStatus) = "01" Then
                        rsTo!lm_receipt = rsTo!lm_receipt + CDbl(Trim(Qty))
                        rsTo!lm_current = rsTo!lm_current + CDbl(Trim(Qty))
                        rsTo!Last_Update = Now
                        rsTo!last_user = userLogin
                        rsTo.update
                    End If
                        
                End If
                    
            End If
            
        Case 1:
        
            If Receipt = False Then '#From Supply
            
                '#Check if Supply influence stock in from warehouse or not
                If FromControlCls = "01" Then
                    RSIns!tm_current = RSIns!tm_current - CDbl(Trim(Qty))
                    RSIns!tm_supply = RSIns!tm_supply + IIf(Trim(inStatus) <> "L" And Trim(inStatus) <> "RJ", CDbl(Trim(Qty)), 0)
                    RSIns!tm_lossreject = RSIns!tm_lossreject + IIf(Trim(inStatus) = "L" Or Trim(inStatus) = "RJ", CDbl(Trim(Qty)), 0)
                    RSIns!nm_premonth = RSIns!tm_current
                    RSIns!nm_current = RSIns!nm_premonth + RSIns!nm_receipt - RSIns!nm_lossreject - RSIns!nm_supply
                    RSIns!Last_Update = Now
                    RSIns!last_user = userLogin
                    RSIns.update
                End If
                
                '#Skip influence To Warehouse Stock or Not
                If skip = False Then
                
                    '#If not Supply-Supply Transaction then dn't influence Stock in toWarehouse
                    If Trim(inStatus) <> "S1" Then Exit Sub
                    
                    '#Check if Supply influence stock in To Warehouse or not
                    If Trim(toStatus) = "01" Then
                        rsTo!tm_receipt = rsTo!tm_receipt + CDbl(Trim(Qty))
                        rsTo!tm_current = rsTo!tm_current + CDbl(Trim(Qty))
                        rsTo!nm_premonth = rsTo!tm_current
                        rsTo!nm_current = rsTo!nm_premonth + rsTo!nm_receipt - rsTo!nm_lossreject - rsTo!nm_supply
                        rsTo!Last_Update = Now
                        rsTo!last_user = userLogin
                        rsTo.update
                    End If
                    
                End If
                
            End If
            
        Case 2:
        
            If Receipt = False Then '#From supply
            
                '#Check if Supply influence stock in from warehouse or not
                If FromControlCls = "01" Then
                    RSIns!nm_current = RSIns!nm_current - CDbl(Trim(Qty))
                    RSIns!nm_supply = RSIns!nm_supply + IIf(Trim(inStatus) <> "L" And Trim(inStatus) <> "RJ", CDbl(Trim(Qty)), 0)
                    RSIns!nm_lossreject = RSIns!nm_lossreject + IIf(Trim(inStatus) = "L" Or Trim(inStatus) = "RJ", CDbl(Trim(Qty)), 0)
                    RSIns!Last_Update = Now
                    RSIns!last_user = userLogin
                    RSIns.update
                End If
                
                '#Skip influence To Warehouse Stock or Not
                If skip = False Then
                
                    '#If not Supply-Supply Transaction then dn't influence Stock in toWarehouse
                    If Trim(inStatus) <> "S1" Then Exit Sub
                    
                    '#Check if Supply influence stock in To Warehouse or not
                    If Trim(toStatus) = "01" Then
                        rsTo!nm_receipt = rsTo!nm_receipt + CDbl(Trim(Qty))
                        rsTo!nm_current = rsTo!nm_current + CDbl(Trim(Qty))
                        rsTo!Last_Update = Now
                        rsTo!last_user = userLogin
                        rsTo.update
                    End If
                    
                End If
                
            End If
            
        End Select
    
    End If

End Sub

Public Sub up_EraseBlankDataInStockMaster(WHCode As String, ItemCode As String, lotno As String)
'    Exit Sub
'    Dim sqldel2 As String
'    Dim recAff2
'
'    '# Erase data from stockMaster ( base on WareHouse code)
'    sqldel2 = "update  stock_master " & _
'        "set item_code=item_code, last_update = getdate(), last_user = '" & userLogin & "' " & _
'        "where item_code='" & Trim(ItemCode) & "' " & _
'        "and warehouse_code='" & Trim(WHCode) & "' " & _
'        "and lm_premonth=0 and nm_premonth=0 and tm_premonth=0 " & _
'        "and lm_receipt=0 and tm_receipt=0 and nm_receipt=0 " & _
'        "and lm_supply=0 and tm_supply=0 and nm_supply=0 " & _
'        "and lm_lossreject=0 and tm_lossreject=0 and nm_lossreject=0 " & _
'        "and lm_current=0 and tm_current=0 and nm_current=0 " & _
'        "and (lm_inventory=0 or lm_inventory is null) " & _
'        "and (tm_inventory=0 or tm_inventory is null) " & _
'        "and (nm_inventory=0  or nm_inventory is null) "
'    Db.Execute sqldel2, recAff2
'    If recAff2 <> 0 Then
'        Db.Execute "delete from stock_master with (updlock) " & _
'            "where item_code='" & Trim(ItemCode) & "' " & _
'            "and warehouse_code='" & Trim(WHCode) & "'"
'    End If
    
End Sub


Public Function up_ValidateDateRange(n As Date, booUpdate As Boolean) As String

    'Dim li_diff As Integer
    li_diff = DateDiff("M", uf_GetLastClosing("fulldate"), n)
    If booUpdate Then
'        Disable untuk standar
'        If StatusAdmin = 1 Then
            If li_diff > 2 Or li_diff < 0 Then
                up_ValidateDateRange = DisplayMsg(8022) 'Please input valid daterange
            Else
                up_ValidateDateRange = ""
            End If
    Else
        If li_diff > 2 Or li_diff < 0 Then
            up_ValidateDateRange = DisplayMsg(8022) 'Please input valid daterange
            
        Else
            up_ValidateDateRange = ""
        End If
    End If

End Function

Public Function up_GetDateRange(n As Date) As Integer

    up_GetDateRange = DateDiff("M", uf_GetLastClosing("fulldate"), n)

End Function

Public Function uf_GetItemDescription(n As String) As String
    
    Dim RS As New ADODB.Recordset
    
    sql = "select im.*,description from item_master im left join sheetcoil_cls sh on im.sheetcoil_cls=sh.sheetcoil_cls  where item_code='" & Trim(n) & "'"
    Set RS = Db.Execute(sql)
    If Not (RS.EOF And RS.BOF) Then
        If Trim(RS!sheetcoil_cls) <> "" Then
            uf_GetItemDescription = Trim(RS!item_name) '& " (" & Trim(rs!Description) & ", T" & rs!Thickness & " x W" & rs!Width & " x L" & rs!Length & ")"
        Else
            uf_GetItemDescription = Trim(RS!item_name)
        End If
    End If
    RS.Close

End Function

Public Function uf_GetCurrencyDescription(ls_Code As String) As String

    Dim RS As New ADODB.Recordset

    sql = "select description from curr_cls where curr_cls='" & Trim(ls_Code) & "'"
    If RS.State <> adStateClosed Then RS.Close
    RS.Open sql, Db, adOpenKeyset, adLockOptimistic
    If Not (RS.EOF And RS.BOF) Then
        uf_GetCurrencyDescription = Trim(RS!Description)
    End If
    If RS.State <> adStateClosed Then RS.Close

End Function

Public Function uf_GetUnitDescription(ls_Code As String) As String
    
    Dim RS As New ADODB.Recordset
    
    sql = "select description from unit_cls where unit_cls='" & Trim(ls_Code) & "'"
    If RS.State <> adStateClosed Then RS.Close
    RS.Open sql, Db, adOpenKeyset, adLockOptimistic
    If Not (RS.EOF And RS.BOF) Then
        uf_GetUnitDescription = Trim(RS!Description)
    End If
    If RS.State <> adStateClosed Then RS.Close
    
End Function

Public Function uf_GetCurrencyCode(ls_description As String) As String
    
    Dim RS As New ADODB.Recordset
    
    sql = "select curr_cls from curr_cls where description='" & Trim(ls_description) & "'"
    If RS.State <> adStateClosed Then RS.Close
    RS.Open sql, Db, adOpenKeyset, adLockOptimistic
    If Not (RS.EOF And RS.BOF) Then
        uf_GetCurrencyCode = Trim(RS!curr_cls)
    Else
        uf_GetCurrencyCode = "02"
    End If
    If RS.State <> adStateClosed Then RS.Close
    
End Function

Public Function uf_GetUnitCode(ls_description As String) As String
    
    Dim RS As New ADODB.Recordset
    
    sql = "select unit_cls from unit_cls where description='" & Trim(ls_description) & "'"
    If RS.State <> adStateClosed Then RS.Close
    RS.Open sql, Db, adOpenKeyset, adLockOptimistic
    If Not (RS.EOF And RS.BOF) Then
            uf_GetUnitCode = Trim(RS!Unit_cls)
    End If
    If RS.State <> adStateClosed Then RS.Close
    
End Function

Public Function uf_GetComboListIndex(cboTemp As MSForms.ComboBox, strValue As String) As Integer
    
    Dim IntIndex As Integer
    
    uf_GetComboListIndex = -1
    For IntIndex = 0 To cboTemp.ListCount - 1
        If Trim(cboTemp.List(IntIndex)) = strValue Then
            uf_GetComboListIndex = IntIndex
            Exit For
        End If
    Next
End Function

Public Function uf_ValidateComboData(cboTemp As MSForms.ComboBox, ErrCode As String, LblErr As Object, LblDescription As Object) As Boolean
    Dim Index As Integer
    Index = uf_GetComboListIndex(cboTemp, cboTemp.Text)
    If Index < 0 Then
        LblErr = DisplayMsg(ErrCode)
            LblDescription = ""
        uf_ValidateComboData = False
        
    Else
        LblErr = ""
            LblDescription = cboTemp.List(Index, 1)
        uf_ValidateComboData = True
    End If
End Function

Public Function uf_Trunc(dblValue As Double, intDigit As Integer) As Double
    If dblValue > 0 Then
        uf_Trunc = Fix(dblValue) + Left$(Round(dblValue - Fix(dblValue), 10), intDigit + 2)
    Else
        uf_Trunc = Fix(dblValue) + Left$(Round(dblValue - Fix(dblValue), 10), intDigit + 3)
    End If
End Function

Public Function uf_ValidClosingReceipt(dtDate As Date) As String
    Dim adoRs As New ADODB.Recordset
    sql = "select closing_date from closing_receipt where closing_year = '" & Year(dtDate) & "' and closing_month = '" & Month(dtDate) & "' and status = '1'"
    adoRs.Open sql, Db, adOpenForwardOnly, adLockReadOnly
    If adoRs.EOF Then
        uf_ValidClosingReceipt = DisplayMsg(8108)
    ElseIf StatusAdmin <> 1 Then
        If dtDate > adoRs.Fields("closing_date") Then uf_ValidClosingReceipt = DisplayMsg(8109) & " " & Format(adoRs.Fields("closing_date"), "dd MMM yyyy")
    End If
    adoRs.Close
    'fungsi di disable untuk standar
    uf_ValidClosingReceipt = ""
End Function

' Function To Create Serial No To automatic
' Add 20090204

Public Function GetSerialTo(SerialFrom As String, Qty)
Dim VTo As Long
Dim JDigit As Integer

If Qty > 0 Then
    JDigit = Len(SerialFrom)
    VTo = Val(Mid(SerialFrom, 2, JDigit - 1)) + (CDbl(Qty) - 1)
    GetSerialTo = Left(SerialFrom, 1) & Format(VTo, String(JDigit - 1, "0"))
Else
    GetSerialTo = ""
End If

End Function

Public Function G_CekExcelApp() As Boolean
On Error Resume Next
    Dim ExlApp As New Excel.application, strVer As String
    strVer = ExlApp.version
    If err.number <> 0 Then
        err.clear
    Else
        G_CekExcelApp = True
    End If
    Set ExlApp = Nothing
End Function
