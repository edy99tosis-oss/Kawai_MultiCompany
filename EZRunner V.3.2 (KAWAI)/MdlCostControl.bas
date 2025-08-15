Attribute VB_Name = "MdlCostControl"
Option Explicit

Public Sub up_CreateSQLFunctionALL()
    
    'Exit Sub
    
    Call up_CreateSQLFunctionGetBookExchangeRate
    Call up_CreateSQLFunctionGetStockMutation
    Call up_CreateSQLFunctionGetAverageMaterialConsumption
    Call up_CreateSQLFunctionGetPreviousValuationPrice
    Call up_CreateSQLFunctionGetAverageProcessCost
    Call up_CreateSQLFunctionGetAverageProcessPrice
    Call up_CreateSQLFunctionGetAverageProductionCost
    Call up_CreateSQLFunctionGetValuationPrice

End Sub

Public Sub up_DropSQLFunctionALL()

    Dim ls_sql As String
    
    'Exit Sub
    
    ls_sql = "IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[UF_GetAverageMaterialConsumption]') AND type in (N'FN', N'IF', N'TF', N'FS', N'FT')) " & _
        "DROP FUNCTION [dbo].[UF_GetAverageMaterialConsumption]"
    Db.Execute ls_sql
    
    ls_sql = "IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[UF_GetAverageProcessCost]') AND type in (N'FN', N'IF', N'TF', N'FS', N'FT')) " & _
        "DROP FUNCTION [dbo].[UF_GetAverageProcessCost]"
    Db.Execute ls_sql
    
    ls_sql = "IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[UF_GetAverageProcessPrice]') AND type in (N'FN', N'IF', N'TF', N'FS', N'FT')) " & _
        "DROP FUNCTION [dbo].[UF_GetAverageProcessPrice]"
    Db.Execute ls_sql
    
    ls_sql = "IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[UF_GetAverageProductionCost]') AND type in (N'FN', N'IF', N'TF', N'FS', N'FT')) " & _
        "DROP FUNCTION [dbo].[UF_GetAverageProductionCost]"
    Db.Execute ls_sql
    
    ls_sql = "IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[UF_GetBookExchangeRate]') AND type in (N'FN', N'IF', N'TF', N'FS', N'FT')) " & _
        "DROP FUNCTION [dbo].[UF_GetBookExchangeRate]"
    Db.Execute ls_sql
    
    ls_sql = "IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[UF_GetPreviousValuationPrice]') AND type in (N'FN', N'IF', N'TF', N'FS', N'FT')) " & _
        "DROP FUNCTION [dbo].[UF_GetPreviousValuationPrice]"
    Db.Execute ls_sql
    
    ls_sql = "IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[UF_GetStockMutation]') AND type in (N'FN', N'IF', N'TF', N'FS', N'FT')) " & _
        "DROP FUNCTION [dbo].[UF_GetStockMutation]"
    Db.Execute ls_sql
    
    ls_sql = "IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[UF_GetValuationPrice]') AND type in (N'FN', N'IF', N'TF', N'FS', N'FT')) " & _
        "DROP FUNCTION [dbo].[UF_GetValuationPrice]"
    Db.Execute ls_sql

End Sub

Public Sub up_CreateSQLFunctionGetBookExchangeRate()

    Dim ls_sql As String
    
    ls_sql = " create function UF_GetBookExchangeRate (@Year char(4),@Month varchar(2),@CurrencyCode char(2)) " & vbCrLf & _
                      " Returns numeric(18,5) " & vbCrLf & _
                      " AS " & vbCrLf & _
                      " BEGIN " & vbCrLf & _
                      "     Declare @ExchangeRate  numeric(18,5) " & vbCrLf & _
                      "     Set @ExchangeRate= " & vbCrLf & _
                      "     ( " & vbCrLf & _
                      "         Select Case @Month " & vbCrLf & _
                      "             WHEN '01' THEN exch01  WHEN '02' THEN exch02  WHEN '03' THEN exch03  WHEN '04' THEN exch04   " & vbCrLf & _
                      "             WHEN '05' THEN exch05 WHEN '06' THEN exch06 WHEN '07' THEN exch07  WHEN '08' THEN exch08   " & vbCrLf & _
                      "             WHEN '09' THEN exch09 WHEN '10' THEN exch010 WHEN '11' THEN exch011  WHEN '12' THEN exch012  " & vbCrLf
    
    ls_sql = ls_sql + "             End  ExRate " & vbCrLf & _
                      "         from Book_ExchangeRate    " & vbCrLf & _
                      "         where rtrim(Book_ExchangeRate.exch_year) = @Year " & vbCrLf & _
                      "         and  Book_ExchangeRate.term_cls=(select ValuationPrice_ExchTerm from company_profile) " & vbCrLf & _
                      "         and  Book_ExchangeRate.Currency_Code=@CurrencyCode " & vbCrLf & _
                      "     ) " & vbCrLf & _
                      "     IF (@CurrencyCode='03')  " & vbCrLf & _
                      "     BEGIN " & vbCrLf & _
                      "         Set @ExchangeRate=1 " & vbCrLf & _
                      "     END " & vbCrLf & _
                      "     Return isnull(@ExchangeRate,0) " & vbCrLf
    
    ls_sql = ls_sql + " END "
    Db.Execute ls_sql

End Sub

Public Sub up_CreateSQLFunctionGetStockMutation()

    Dim ls_sql As String
    
    ls_sql = " Create Function UF_GetStockMutation(@Year char(4) ,@Month varchar(2),@ItemCode char(25),@Data char(20)) " & vbCrLf & _
                      " Returns numeric (18,5) " & vbCrLf & _
                      " AS " & vbCrLf & _
                      " BEGIN    " & vbCrLf & _
                      "     Declare @DiffWithClosingPeriod int " & vbCrLf & _
                      "     Declare @Stock numeric (18,5) " & vbCrLf & _
                      "     Declare @InventoryClosingYear char(4) " & vbCrLf & _
                      "     Declare @InventoryClosingMonth varchar(2) " & vbCrLf & _
                      "     select @InventoryClosingMonth = ( select top 1 inventory_month  from inventory_control  order by inventory_year desc , inventory_month desc )   " & vbCrLf & _
                      "     select @InventoryClosingYear = ( select top 1 inventory_year from inventory_control  order by inventory_year desc , inventory_month desc )   " & vbCrLf & _
                      "     Set @DiffWithClosingPeriod= datediff(month,cast(@InventoryClosingYear as varchar(4)) + '-' +  cast(@InventoryClosingMonth as varchar(2)) + '-01', " & vbCrLf
    
    ls_sql = ls_sql + "                             cast(@Year as varchar(4)) + '-' +  cast(@Month as varchar(2)) + '-01') " & vbCrLf & _
                      "     Set @Stock =( " & vbCrLf & _
                      "         Select sum( " & vbCrLf & _
                      "             case @DiffWithClosingPeriod when 0 then " & vbCrLf & _
                      "                 Case  when @Data= 'in' then " & vbCrLf & _
                      "                     isnull(lm_receipt,0) " & vbCrLf & _
                      "                 when @Data='out' then " & vbCrLf & _
                      "                     isnull(lm_supply,0) + isnull(lm_lossreject,0) " & vbCrLf & _
                      "                 when @Data='InOther' or @Data='OutOther' then " & vbCrLf & _
                      "                     isnull(lm_Inventory,0) - isnull(lm_Current,0) " & vbCrLf & _
                      "                 end              " & vbCrLf
    
    ls_sql = ls_sql + "             When 1 then " & vbCrLf & _
                      "                 Case  when @Data= 'in' then " & vbCrLf & _
                      "                     isnull(tm_receipt,0) " & vbCrLf & _
                      "                 when @Data='out' then " & vbCrLf & _
                      "                     isnull(tm_supply,0) + isnull(tm_lossreject,0) " & vbCrLf & _
                      "                 when @Data='InOther' or @Data='OutOther' then " & vbCrLf & _
                      "                     0--isnull(tm_Inventory,0) - isnull(tm_Current,0) " & vbCrLf & _
                      "                 end          " & vbCrLf & _
                      "             When 2 then " & vbCrLf & _
                      "                 Case  when @Data= 'in' then " & vbCrLf & _
                      "                     isnull(nm_receipt,0) " & vbCrLf
    
    ls_sql = ls_sql + "                 when @Data='out' then " & vbCrLf & _
                      "                     isnull(nm_supply,0) + isnull(nm_lossreject,0) " & vbCrLf & _
                      "                 when @Data='InOther' or @Data='OutOther' then " & vbCrLf & _
                      "                     0--isnull(nm_Inventory,0) - isnull(nm_Current,0) " & vbCrLf & _
                      "                 end      " & vbCrLf & _
                      "             End " & vbCrLf & _
                      "             ) " & vbCrLf & _
                      "         From Stock_Master  where Item_Code=@ItemCode " & vbCrLf & _
                      "         Group by Item_Code " & vbCrLf & _
                      "         ) " & vbCrLf & _
                      "     if  (@Data='InOther' and   isnull(@Stock,0) <=0 ) " & vbCrLf
    
    ls_sql = ls_sql + "     BEGIN " & vbCrLf & _
                      "         set @Stock= 0 " & vbCrLf & _
                      "     END " & vbCrLf & _
                      "     Else if  (@Data='OutOther' and   isnull(@Stock,0) >0 ) " & vbCrLf & _
                      "     BEGIN " & vbCrLf & _
                      "         set @Stock= 0 " & vbCrLf & _
                      "     END " & vbCrLf & _
                      "  " & vbCrLf & _
                      "     Return abs(isnull(@Stock,0)) " & vbCrLf & _
                      " END " & vbCrLf
    Db.Execute ls_sql

End Sub

Public Sub up_CreateSQLFunctionGetAverageMaterialConsumption()

    Dim ls_sql As String
    
    ls_sql = "  " & vbCrLf & _
                      " Create Function UF_GetAverageMaterialConsumption(@Year char(4) ,@Month varchar(2),@ParentItemCode char(25),@ChildItemCode char(15),@Optional_ReceiptSeqNo as char(18)) " & vbCrLf & _
                      " Returns Numeric (18,5) " & vbCrLf & _
                      " AS " & vbCrLf & _
                      " BEGIN " & vbCrLf & _
                      "     Declare @AverageConsumption numeric (18,5) " & vbCrLf & _
                      "     IF (rtrim(@Optional_ReceiptSeqNo)='ALL')  " & vbCrLf & _
                      "         BEGIN " & vbCrLf & _
                      "         --Average ALL Consumsumption Data " & vbCrLf & _
                      "         Set @AverageConsumption =( " & vbCrLf & _
                      "             Select  " & vbCrLf
    
    ls_sql = ls_sql + "                 case when  Sum(isnull(PR.Qty, 0)) = 0 then " & vbCrLf & _
                      "                     0 " & vbCrLf & _
                      "                 else " & vbCrLf & _
                      "                     Sum(isnull(PS.Consumption_Qty, 0)) / Sum(isnull(PR.Qty, 0)) " & vbCrLf & _
                      "                 End  " & vbCrLf & _
                      "             from part_supply PS  " & vbCrLf & _
                      "             Left Join Part_Receipt PR on PR.Seq_No= PS.DO_No " & vbCrLf & _
                      "             where PS.Supply_Cls='S'  " & vbCrLf & _
                      "             and PS.Consumption_qty is not null " & vbCrLf & _
                      "             and PS.ParentItem_Code=@ParentItemCode " & vbCrLf & _
                      "             and PS.ChildItem_Code=@ChildItemCode " & vbCrLf
    
    ls_sql = ls_sql + "             and month (PS.ChildSupply_Date)=@Month " & vbCrLf & _
                      "             and Year (PS.ChildSupply_Date)=@Year " & vbCrLf & _
                      "             ) " & vbCrLf & _
                      "         END  " & vbCrLf & _
                      "     ELSE " & vbCrLf & _
                      "         BEGIN " & vbCrLf & _
                      "         --Average  Consumsumption Data for selected Production Result " & vbCrLf & _
                      "         Set @AverageConsumption =isnull(( " & vbCrLf & _
                      "             Select  " & vbCrLf & _
                      "                 case when  Sum(isnull(PR.Qty, 0)) = 0 then " & vbCrLf & _
                      "                     0 " & vbCrLf
    
    ls_sql = ls_sql + "                 else " & vbCrLf & _
                      "                     Sum(isnull(PS.Consumption_Qty, 0)) / Sum(isnull(PR.Qty, 0)) " & vbCrLf & _
                      "                 End  " & vbCrLf & _
                      "             from part_supply PS  " & vbCrLf & _
                      "             Left Join Part_Receipt PR on PR.Seq_No= PS.DO_No " & vbCrLf & _
                      "             where PS.Supply_Cls='S'  " & vbCrLf & _
                      "             and PS.Consumption_qty is not null " & vbCrLf & _
                      "             and PS.ParentItem_Code=@ParentItemCode " & vbCrLf & _
                      "             and PS.ChildItem_Code=@ChildItemCode " & vbCrLf & _
                      "             and month (PS.ChildSupply_Date)=@Month " & vbCrLf & _
                      "             and Year (PS.ChildSupply_Date)=@Year " & vbCrLf
    
    ls_sql = ls_sql + "             and PR.Seq_No = rtrim(@Optional_ReceiptSeqNo) " & vbCrLf & _
                      "             ), 0) " & vbCrLf & _
                      "         END  " & vbCrLf & _
                      "  " & vbCrLf & _
                      "     IF (isnull(@AverageConsumption,0)=0) " & vbCrLf & _
                      "     Begin " & vbCrLf & _
                      "         Set @AverageConsumption=(  " & vbCrLf & _
                      "             Select Qty from BOM_Master BM  " & vbCrLf & _
                      "             where BM.Parent_ItemCode =@ParentItemCode  " & vbCrLf & _
                      "             and BM.Item_Code=@ChildItemCode)     " & vbCrLf & _
                      "     End " & vbCrLf
    
    ls_sql = ls_sql + "     Return isnull(@AverageConsumption,0) " & vbCrLf & _
                      " END " & vbCrLf & _
                      "  "
    
    Db.Execute ls_sql

End Sub

Public Sub up_CreateSQLFunctionGetPreviousValuationPrice()

    Dim ls_sql As String
    
    ls_sql = "  " & vbCrLf & _
                      " Create Function UF_GetPreviousValuationPrice(@Year as char(4),@Month as varchar(2),@itemCode as char(25)) " & vbCrLf & _
                      " RETURNS numeric(18,5) " & vbCrLf & _
                      " AS " & vbCrLf & _
                      " BEGIN " & vbCrLf & _
                      "     Declare @Price numeric(18,5) " & vbCrLf & _
                      "     Set @Price =( Select top 1 Inventory_price from  " & vbCrLf & _
                      "             (Select Inventory_Price ,Inventory_Year ,Inventory_Month from Inventory_Price InventoryPricePrevious  " & vbCrLf & _
                      "             where InventoryPricePrevious.item_code =@ItemCode " & vbCrLf & _
                      "             and InventoryPricePrevious.inventory_year <> @Year " & vbCrLf & _
                      "             and InventoryPricePrevious.inventory_month <> @Month " & vbCrLf
    
    ls_sql = ls_sql + "             UNION ALL " & vbCrLf & _
                      "             Select Inventory_Price ,Inventory_Year ,Inventory_Month from InventoryPrice_HISTORY InventoryPricePrevious  " & vbCrLf & _
                      "             where InventoryPricePrevious.item_code =@ItemCode " & vbCrLf & _
                      "             and InventoryPricePrevious.inventory_year <> @Year " & vbCrLf & _
                      "             and InventoryPricePrevious.inventory_month <> @Month) tbA order by " & vbCrLf & _
                      "              Inventory_Year desc ,Inventory_Month desc " & vbCrLf & _
                      "                ) " & vbCrLf & _
                      "     Return Isnull(@Price,0) " & vbCrLf & _
                      " END " & vbCrLf & _
                      "  "
    Db.Execute ls_sql

End Sub

Public Sub up_CreateSQLFunctionGetAverageProcessCost()

    Dim ls_sql As String
    
    ls_sql = "  " & vbCrLf & _
                      " Create Function UF_GetAverageProcessCost(@Year as char(4),@Month as varchar(2),@itemCode as char(25),@Optional_ReceiptSeqNo as char(18),@Optional_CalculateWhat as char(18)) " & vbCrLf & _
                      " RETURNS numeric(18,5) " & vbCrLf & _
                      " AS " & vbCrLf & _
                      " BEGIN " & vbCrLf & _
                      "     Declare @EndMonthDate char(10) " & vbCrLf & _
                      "     Declare @Time numeric(18,5) " & vbCrLf & _
                      "     Declare @AdditionalCost numeric (18,5) " & vbCrLf & _
                      "  " & vbCrLf & _
                      "     Set @EndMonthDate = convert(char(10),dateadd(day,-1,dateadd(month, 1,cast(@Year as char(4)) + '-' +  cast(@Month as varchar(2)) + '-01')),120) " & vbCrLf & _
                      "     Set @AdditionalCost=(select sum(Addcost) from( " & vbCrLf
    
    ls_sql = ls_sql + "                  select  " & vbCrLf & _
                      "                      " & vbCrLf & _
                      "                         case when (   " & vbCrLf & _
                      "                             select isnull(MI.Amount,0) * dbo.UF_GetBookExchangeRate(@Year,@Month,MI.Currency_Code)   " & vbCrLf & _
                      "                             from InventoryCost_Item MI   " & vbCrLf & _
                      "                             where MI.Item_Code=@ItemCode  " & vbCrLf & _
                      "                             and Cost_Cls=MC.Cost_Cls   " & vbCrLf & _
                      "                             and Start_Date <=replace(@EndMonthDate,'-','')   " & vbCrLf & _
                      "                             and End_Date>= replace(@EndMonthDate,'-','')   " & vbCrLf & _
                      "                         ) = null then " & vbCrLf & _
                      "                         isnull((   " & vbCrLf
    
    ls_sql = ls_sql + "                             select isnull(MI.Amount,0) * dbo.UF_GetBookExchangeRate(@Year,@Month,MI.Currency_Code)   " & vbCrLf & _
                      "                             from InventoryCost_Group MI   " & vbCrLf & _
                      "                             Where Cost_Cls=MC.Cost_Cls   " & vbCrLf & _
                      "                             and Start_Date <=replace(@EndMonthDate,'-','')   " & vbCrLf & _
                      "                             and End_Date>=replace(@EndMonthDate,'-','')   " & vbCrLf & _
                      "                         ),0) End AddCost " & vbCrLf & _
                      "                                          " & vbCrLf & _
                      "                 from inventorycost_master MC   " & vbCrLf & _
                      "                 where  MC.additional_cls='0'    )tbA             " & vbCrLf & _
                      "             ) " & vbCrLf & _
                      "     IF (RTRIM(@Optional_CalculateWhat)='ALL' ) " & vbCrLf
    
    ls_sql = ls_sql + "     BEGIN " & vbCrLf & _
                      "         IF (rtrim(@Optional_ReceiptSeqNo)='ALL')  " & vbCrLf & _
                      "             BEGIN " & vbCrLf & _
                      "                 Set @Time =(select case when sum(isnull(PR.Qty,0)) =0 then 0 else  " & vbCrLf & _
                      "                         sum(isnull(WTM.TotalWorking_Time,0))/sum(isnull(PR.Qty,0)) End            " & vbCrLf & _
                      "                         from Part_Receipt PR " & vbCrLf & _
                      "                         left join WorkingTime_Master WTM on PR.Seq_No=WTM.ProductionSeq_No " & vbCrLf & _
                      "                         where PR.Receipt_cls='P1'  " & vbCrLf & _
                      "                         and month(PR.Receipt_Date)=@Month " & vbCrLf & _
                      "                         and year(PR.Receipt_Date)=@Year) " & vbCrLf & _
                      "             END " & vbCrLf
    
    ls_sql = ls_sql + "         ELSE " & vbCrLf & _
                      "             BEGIN " & vbCrLf & _
                      "                 Set @Time =(select case when sum(isnull(PR.Qty,0)) =0 then 0 else  " & vbCrLf & _
                      "                     sum(isnull(WTM.TotalWorking_Time,0))/sum(isnull(PR.Qty,0)) End            " & vbCrLf & _
                      "                     from Part_Receipt PR " & vbCrLf & _
                      "                     left join WorkingTime_Master WTM on PR.Seq_No=WTM.ProductionSeq_No " & vbCrLf & _
                      "                     where PR.Receipt_cls='P1'  " & vbCrLf & _
                      "                     and month(PR.Receipt_Date)=@Month " & vbCrLf & _
                      "                     and year(PR.Receipt_Date)=@Year " & vbCrLf & _
                      "                     and PR.Seq_No=rtrim(@Optional_ReceiptSeqNo)) " & vbCrLf & _
                      "             END " & vbCrLf
    
    ls_sql = ls_sql + "  " & vbCrLf & _
                      "         IF (Isnull(@Time,0)=0)  " & vbCrLf & _
                      "         BEGIN " & vbCrLf & _
                      "             Set @Time =(Select sum(isnull(Standard_Time,0)) from Process_Master where Item_Code=@ItemCode) " & vbCrLf & _
                      "         END  " & vbCrLf & _
                      "     END " & vbCrLf & _
                      "  " & vbCrLf & _
                      "     IF (RTRIM(@Optional_CalculateWhat)='ALL' ) " & vbCrLf & _
                      "         BEGIN " & vbCrLf & _
                      "         --Return  Process Cost + Additional Cost " & vbCrLf & _
                      "         Set @Time=  (Isnull(@Time,0) * dbo.UF_GetAverageProcessPrice(@EndMonthDate,@itemCode )) +@AdditionalCost " & vbCrLf
    
    ls_sql = ls_sql + "         END " & vbCrLf & _
                      "     ELSE    IF (RTRIM(@Optional_CalculateWhat)='ProcessCost' ) " & vbCrLf & _
                      "         BEGIN " & vbCrLf & _
                      "         --Return  Process Cost " & vbCrLf & _
                      "         Set @Time=  (Isnull(@Time,0) * dbo.UF_GetAverageProcessPrice(@EndMonthDate,@itemCode )) " & vbCrLf & _
                      "         END " & vbCrLf & _
                      "  " & vbCrLf & _
                      "     ELSE IF (RTRIM(@Optional_CalculateWhat)='AdditionalCost' ) " & vbCrLf & _
                      "         BEGIN " & vbCrLf & _
                      "         --Return   Additional Cost " & vbCrLf & _
                      "         Set @Time=  @AdditionalCost " & vbCrLf
    
    ls_sql = ls_sql + "         END " & vbCrLf & _
                      "      " & vbCrLf & _
                      "     Return isnull(@Time, 0) " & vbCrLf & _
                      " END " & vbCrLf
    Db.Execute ls_sql

End Sub

Public Sub up_CreateSQLFunctionGetAverageProcessPrice()

    Dim ls_sql As String
    
    ls_sql = " Create Function UF_GetAverageProcessPrice(@Date as char(10),@itemCode as char(25)) " & vbCrLf & _
                      " RETURNS numeric(18,5) " & vbCrLf & _
                      " AS " & vbCrLf & _
                      " BEGIN " & vbCrLf & _
                      "     Declare @Price numeric(18,5) " & vbCrLf & _
                      "     Set @Price=( " & vbCrLf & _
                      "         Select  " & vbCrLf & _
                      "             case when Sum(isnull(Standard_Time, 0)) =0 then 0 else " & vbCrLf & _
                      "             Sum(case when PCM.Trade_code is null then    " & vbCrLf & _
                      "                 isnull(Standard_time,0) * isnull(Cost_Minute,0) *  " & vbCrLf & _
                      "                 dbo.UF_GetBookExchangeRate(Year(@Date),Month(@Date),PCM.Currency_Code)  " & vbCrLf
    
    ls_sql = ls_sql + "             else " & vbCrLf & _
                      "                 PM1.PriceSubCon " & vbCrLf & _
                      "             end )/       " & vbCrLf & _
                      "             Sum(Standard_Time) End " & vbCrLf & _
                      "         from Process_Master PCM      " & vbCrLf & _
                      "         Left Join " & vbCrLf & _
                      "         ( " & vbCrLf & _
                      "             Select PM1.Trade_Code,PM1.Item_Code,sum(Isnull(Price,0) *  " & vbCrLf & _
                      "                 dbo.UF_GetBookExchangeRate(Year(@Date),Month(@Date), PM1.Currency_Code)  " & vbCrLf & _
                      "                 ) PriceSubcon from Price_Master PM1  " & vbCrLf & _
                      "             Where PM1.Price_Cls in ('01','05') " & vbCrLf
    
    ls_sql = ls_sql + "             and PM1.Start_Date<= Replace(@Date,'-','') " & vbCrLf & _
                      "             and PM1.End_Date>= Replace(@Date,'-','') " & vbCrLf & _
                      "             and PM1.Priority_Cls='1' " & vbCrLf & _
                      "             Group by PM1.Trade_Code ,PM1.Item_Code " & vbCrLf & _
                      "         )PM1 " & vbCrLf & _
                      "         on PM1.Trade_Code=  PCM.Trade_Code  " & vbCrLf & _
                      "         and PM1.Item_Code=  PCM.Item_Code  " & vbCrLf & _
                      "         Where PCM.Item_Code=@ItemCode " & vbCrLf & _
                      "     ) " & vbCrLf & _
                      "     Return Isnull(@Price,0) " & vbCrLf & _
                      " END " & vbCrLf
    
    Db.Execute ls_sql

End Sub

Public Sub up_CreateSQLFunctionGetAverageProductionCost()

    Dim ls_sql As String
    
    ls_sql = "  " & vbCrLf & _
                      " Create Function UF_GetAverageProductionCost(@Year char(4) ,@Month varchar(2),@ItemCode char(25),@Optional_ReceiptSeqNo as char(18),@Optional_CalculateWhat as char(18)) " & vbCrLf & _
                      " Returns Numeric (18,5) " & vbCrLf & _
                      " AS " & vbCrLf & _
                      " BEGIN " & vbCrLf & _
                      "  " & vbCrLf & _
                      "     Declare @Pos as integer " & vbCrLf & _
                      "     Declare @ParentItemCode char(25) " & vbCrLf & _
                      "     Declare @ChildItemCode char(25) " & vbCrLf & _
                      "     Declare @ProductionCost numeric (18,5) " & vbCrLf & _
                      "     Declare @Cost numeric (18,5) " & vbCrLf
    
    ls_sql = ls_sql + "     Declare @Table Table (item_Code char(25)) " & vbCrLf & _
                      "     Declare @Table2 Table (item_Code char(25)) " & vbCrLf & _
                      "     Declare @EndMonthDate char(10) " & vbCrLf & _
                      "     Declare @LastPeriod  datetime " & vbCrLf & _
                      "     Set @EndMonthDate = convert(char(10),dateadd(day,-1,dateadd(month, 1,cast(@Year as char(4)) + '-' +  cast(@Month as varchar(2)) + '-01')),120) " & vbCrLf & _
                      "     Set @LastPeriod = dateadd(month,-1,cast (@EndMonthDate as datetime)) " & vbCrLf & _
                      "  " & vbCrLf & _
                      "  " & vbCrLf & _
                      "     Declare RS Cursor For " & vbCrLf & _
                      "     Select distinct ParentItem_Code,ChildItem_Code                                       " & vbCrLf & _
                      "     from part_supply PS  " & vbCrLf
    
    ls_sql = ls_sql + "     Left Join Part_Receipt PR on PR.Seq_No= PS.DO_No " & vbCrLf & _
                      "     where PS.Supply_Cls='S'  " & vbCrLf & _
                      "     and PS.Consumption_qty is not null " & vbCrLf & _
                      "     and PS.ParentItem_Code = @ItemCode " & vbCrLf & _
                      "     and month (PS.ChildSupply_Date)=@Month " & vbCrLf & _
                      "     and Year (PS.ChildSupply_Date)=@Year " & vbCrLf & _
                      "  " & vbCrLf & _
                      "  " & vbCrLf & _
                      "     Open RS " & vbCrLf & _
                      "     Fetch Next from RS into " & vbCrLf & _
                      "         @ParentItemCode, " & vbCrLf
    
    ls_sql = ls_sql + "         @ChildItemCode   " & vbCrLf & _
                      "     Set @ProductionCost=0 " & vbCrLf & _
                      "     Set @Pos=0 " & vbCrLf & _
                      "  " & vbCrLf & _
                      "     --####################################################################### " & vbCrLf & _
                      "     --Collect ChildItemCode      " & vbCrLf & _
                      "     Insert Into @Table(Item_Code) " & vbCrLf & _
                      "     ( " & vbCrLf & _
                      "         Select distinct ChildItem_Code                                       " & vbCrLf & _
                      "         from part_supply PS  " & vbCrLf & _
                      "         Left Join Part_Receipt PR on PR.Seq_No= PS.DO_No " & vbCrLf
    
    ls_sql = ls_sql + "         where PS.Supply_Cls='S'  " & vbCrLf & _
                      "         and PS.Consumption_qty is not null " & vbCrLf & _
                      "         and PS.ParentItem_Code = @ItemCode " & vbCrLf & _
                      "         and month (PS.ChildSupply_Date)=@Month " & vbCrLf & _
                      "         and Year (PS.ChildSupply_Date)=@Year " & vbCrLf & _
                      "     ) " & vbCrLf & _
                      "     --####################################################################### " & vbCrLf & _
                      "  " & vbCrLf & _
                      "     --FinishGood/WIP (Calculate Average Process Cost) --> For Parent " & vbCrLf & _
                      "     Set @ProductionCost =@ProductionCost + dbo.UF_GetAverageProcessCost(@Year,@Month ,@ItemCode,@Optional_ReceiptSeqNo,@Optional_CalculateWhat) " & vbCrLf & _
                      "  " & vbCrLf
    
    ls_sql = ls_sql + "     While (@Pos =0 ) " & vbCrLf & _
                      "     BEGIN " & vbCrLf & _
                      "         If( @@Fetch_Status=0 ) " & vbCrLf & _
                      "             BEGIN " & vbCrLf & _
                      "                  " & vbCrLf & _
                      "                 Set @Pos=0 " & vbCrLf & _
                      "                 Set @Cost=0 " & vbCrLf & _
                      "                 --Check Wheter Item is FinishGoodWIP or Material " & vbCrLf & _
                      "                 IF (isnull((Select count(Parent_ItemCode) From BOM_Master where Parent_ItemCode=@ChildItemCode),0)<=0) " & vbCrLf & _
                      "                 BEGIN " & vbCrLf & _
                      "                     IF (RTRIM(@Optional_CalculateWhat)='ALL' or RTRIM(@Optional_CalculateWhat)='Material') " & vbCrLf
    
    ls_sql = ls_sql + "                     BEGIN " & vbCrLf & _
                      "                         --Material (Calculate Average Receipt Price) " & vbCrLf & _
                      "                         Set @Cost =dbo.UF_GetAverageMaterialConsumption(@Year ,@Month,@ParentItemCode,@ChildItemCode,@Optional_ReceiptSeqNo) * " & vbCrLf & _
                      "                                   dbo.UF_GetValuationPrice(@Year ,@Month ,@ChildItemCode ,'in','1',@Optional_ReceiptSeqNo) " & vbCrLf & _
                      "  " & vbCrLf & _
                      "                     END " & vbCrLf & _
                      "                     ELSE " & vbCrLf & _
                      "                     BEGIN " & vbCrLf & _
                      "                         Set @Cost=0 " & vbCrLf & _
                      "                     END " & vbCrLf & _
                      "                 END " & vbCrLf
    
    ls_sql = ls_sql + "                 ELSE " & vbCrLf & _
                      "                 BEGIN " & vbCrLf & _
                      "                     IF (RTRIM(@Optional_CalculateWhat)='ALL' or RTRIM(@Optional_CalculateWhat)='ProcessCost' or RTRIM(@Optional_CalculateWhat)='AdditionalCost') " & vbCrLf & _
                      "                     BEGIN " & vbCrLf & _
                      "                         --FinishGood/WIP (Calculate Average Process Cost) " & vbCrLf & _
                      "                         Set @Cost =dbo.UF_GetAverageProcessCost(@Year,@Month ,@ParentItemCode,@Optional_ReceiptSeqNo,@Optional_CalculateWhat) " & vbCrLf & _
                      "                     END " & vbCrLf & _
                      "                     ELSE " & vbCrLf & _
                      "                     BEGIN " & vbCrLf & _
                      "                         Set @Cost=0 " & vbCrLf & _
                      "                     END " & vbCrLf
    
    ls_sql = ls_sql + "                 END " & vbCrLf & _
                      "  " & vbCrLf & _
                      "                 Set @ProductionCost=@ProductionCost + @Cost " & vbCrLf & _
                      "  " & vbCrLf & _
                      "                 Fetch Next from RS into " & vbCrLf & _
                      "                 @ParentItemCode,@ChildItemCode " & vbCrLf & _
                      "             END " & vbCrLf & _
                      "         ELSE " & vbCrLf & _
                      "             BEGIN " & vbCrLf & _
                      "                 IF (( " & vbCrLf & _
                      "                     select count(ParentItem_Code) from " & vbCrLf
    
    ls_sql = ls_sql + "                     (Select distinct ParentItem_Code,ChildItem_Code                                      " & vbCrLf & _
                      "                     from part_supply PS  " & vbCrLf & _
                      "                     Left Join Part_Receipt PR on PR.Seq_No= PS.DO_No " & vbCrLf & _
                      "                     where PS.Supply_Cls='S'  " & vbCrLf & _
                      "                     and PS.Consumption_qty is not null " & vbCrLf & _
                      "                     and PS.ParentItem_Code in (Select Item_Code from @Table ) " & vbCrLf & _
                      "                     and month (PS.ChildSupply_Date)=@Month " & vbCrLf & _
                      "                     and Year (PS.ChildSupply_Date)=@Year)tbA " & vbCrLf & _
                      "                     )<=0) " & vbCrLf & _
                      "  " & vbCrLf & _
                      "                     BEGIN " & vbCrLf
    
    ls_sql = ls_sql + "                         --End Calculation " & vbCrLf & _
                      "                         Set @Pos=1 " & vbCrLf & _
                      "                     END " & vbCrLf & _
                      "                 ELSE " & vbCrLf & _
                      "                     BEGIN " & vbCrLf & _
                      "                         --Set Calculation to Next Level " & vbCrLf & _
                      "                         Close RS " & vbCrLf & _
                      "                         Deallocate RS " & vbCrLf & _
                      "                         Declare RS Cursor for " & vbCrLf & _
                      "                         Select BM2.Parent_ItemCode,BM2.Item_Code " & vbCrLf & _
                      "                         From BOM_Master BM2 " & vbCrLf
    
    ls_sql = ls_sql + "                         where BM2.Parent_ItemCode in " & vbCrLf & _
                      "                         ( Select Item_Code from @Table) " & vbCrLf & _
                      "                         and BM2.Start_Date<= cast(@Year as char(4))+'' + right(cast (100+@Month as char(3)),2) +'01' " & vbCrLf & _
                      "                         and BM2.End_Date >= replace(@EndMonthDate,'-','')                    " & vbCrLf & _
                      "  " & vbCrLf & _
                      "                         --####################################################################### " & vbCrLf & _
                      "                         --Collect ChildItemCode " & vbCrLf & _
                      "                         delete from @Table2 " & vbCrLf & _
                      "                         insert into @Table2 (item_Code) (select item_code from @Table) " & vbCrLf & _
                      "                         delete from @Table " & vbCrLf & _
                      "                         Insert Into @Table(Item_Code)   " & vbCrLf
    
    ls_sql = ls_sql + "                         ( " & vbCrLf & _
                      "                             Select distinct ChildItem_Code                                       " & vbCrLf & _
                      "                             from part_supply PS  " & vbCrLf & _
                      "                             Left Join Part_Receipt PR on PR.Seq_No= PS.DO_No " & vbCrLf & _
                      "                             where PS.Supply_Cls='S'  " & vbCrLf & _
                      "                             and PS.Consumption_qty is not null " & vbCrLf & _
                      "                             and PS.ParentItem_Code in (Select Item_Code from @Table2) " & vbCrLf & _
                      "                             and month (PS.ChildSupply_Date)=@Month " & vbCrLf & _
                      "                             and Year (PS.ChildSupply_Date)=@Year " & vbCrLf & _
                      "                         ) " & vbCrLf & _
                      "                         --####################################################################### " & vbCrLf
    
    ls_sql = ls_sql + "              " & vbCrLf & _
                      "                         Set @Pos=0           " & vbCrLf & _
                      "                         Open RS " & vbCrLf & _
                      "                         Fetch Next from RS into " & vbCrLf & _
                      "                         @ParentItemCode, " & vbCrLf & _
                      "                         @ChildItemCode " & vbCrLf & _
                      "              " & vbCrLf & _
                      "                     END " & vbCrLf & _
                      "             END " & vbCrLf & _
                      "     END " & vbCrLf & _
                      "  "
    
    ls_sql = ls_sql + "     Close RS " & vbCrLf & _
                      "     Deallocate RS " & vbCrLf & _
                      "  " & vbCrLf & _
                      "     IF ( isnull(@ProductionCost,0) =0)  " & vbCrLf & _
                      "     BEGIN " & vbCrLf & _
                      "         Set @ProductionCost=dbo.UF_GetPreviousValuationPrice (year(@LastPeriod),month(@LastPeriod),@ItemCode) " & vbCrLf & _
                      "     END " & vbCrLf & _
                      "     Return  isnull(@ProductionCost,0)  " & vbCrLf & _
                      " END " & vbCrLf
    
    'ls_sql = "  " & vbCrLf & _
    '                  " Create Function UF_GetAverageProductionCost(@Year char(4) ,@Month varchar(2),@ItemCode char(15),@Optional_ReceiptSeqNo as char(18),@Optional_CalculateWhat as char(18)) " & vbCrLf & _
    '                  " Returns Numeric (18,4) " & vbCrLf & _
    '                  " AS " & vbCrLf & _
    '                  " BEGIN " & vbCrLf & _
    '                  " return 0 " & vbCrLf & _
    '                  " End "
    
    Db.Execute ls_sql

End Sub
Public Sub up_CreateSQLFunctionGetValuationPrice()

    Dim ls_sql As String
    
    ls_sql = " Create Function UF_GetValuationPrice(@Year char(4) ,@Month varchar(2),@ItemCode char(25),@Data as char(20),@ReceiptDataOnly as char(1),@Optional_ReceiptSeqNo as char(18)) " & vbCrLf & _
                      " Returns Numeric(18,5) " & vbCrLf & _
                      " AS " & vbCrLf & _
                      " BEGIN " & vbCrLf & _
                      "     Declare @Price numeric(18,5) " & vbCrLf & _
                      "     Declare @FinishGoodCls  char(2) " & vbCrLf & _
                      "     Declare @tbReceipt Table ( " & vbCrLf & _
                      "         [Seq_No] [numeric](18, 0), " & vbCrLf & _
                      "         [Supplier_Code] [char](15) , " & vbCrLf & _
                      "         [PO_No] [char](35) , " & vbCrLf & _
                      "         [Warehouse_Code] [char](15), " & vbCrLf
    
    ls_sql = ls_sql + "         [Address] [char](15) , " & vbCrLf & _
                      "         [Receipt_Cls] [char](2) , " & vbCrLf & _
                      "         [MaterialConsump_Cls] [char](2) , " & vbCrLf & _
                      "         [Receipt_Date] [datetime], " & vbCrLf & _
                      "         [Item_Code] [char](25), " & vbCrLf & _
                      "         [Qty] [numeric](18,5) , " & vbCrLf & _
                      "         [SerialNoFrom] char(10), " & vbCrLf & _
                      "         [SerialNoTo] char(10), " & vbCrLf & _
                      "         [Unit_Cls] [char](2) , " & vbCrLf & _
                      "         [Currency_Code] char(2), " & vbCrLf & _
                      "         [Price] numeric(18, 5) , " & vbCrLf & _
                      "         [Price_Service] numeric(18, 5) , " & vbCrLf & _
                      "         [Amount] numeric(22, 2) , " & vbCrLf & _
                      "         [SuratJalan_No] char(25), " & vbCrLf & _
                      "         [ProductionResult_Cls] char(1), " & vbCrLf
    
    ls_sql = ls_sql + "         [DailySeq_No] numeric(18, 0), " & vbCrLf & _
                      "         [Remarks] char(35), " & vbCrLf & _
                      "         [BC40_No] char(30), " & vbCrLf & _
                      "         [Transport_Cls] char(2), " & vbCrLf & _
                      "         [Package_Cls] char(4), " & vbCrLf & _
                      "         [Package_Qty] numeric(8, 0), " & vbCrLf & _
                      "         [LOT_No] char(35), " & vbCrLf & _
                      "         [Last_Update] datetime, " & vbCrLf & _
                      "         [Last_User] char(15), " & vbCrLf & _
                      "         [Register_Date] datetime, " & vbCrLf & _
                      "         [BC40_Date] DateTime,BC_Type varchar(15)) " & vbCrLf & _
                      "      " & vbCrLf & _
                      "     Set @FinishGoodCls = (select FinishGoodPart_Cls from item_master where item_code=@ItemCode) " & vbCrLf & _
                      "  " & vbCrLf & _
                      "     IF (rtrim(@Optional_ReceiptSeqNo)='ALL')  " & vbCrLf & _
                      "     BEGIN " & vbCrLf & _
                      "         insert into @tbReceipt " & vbCrLf
    
    ls_sql = ls_sql + "             select * from Part_Receipt PR  " & vbCrLf & _
                      "             Where month(PR.receipt_date) =@Month " & vbCrLf & _
                      "             and year(PR.receipt_date) =@Year " & vbCrLf & _
                      "             and PR.Item_Code=@ItemCode " & vbCrLf & _
                      "     END " & vbCrLf & _
                      "     ELSE " & vbCrLf & _
                      "     BEGIN " & vbCrLf & _
                      "         insert into @tbReceipt " & vbCrLf & _
                      "             select * from Part_Receipt PR  " & vbCrLf & _
                      "             Where month(PR.receipt_date) =@Month " & vbCrLf & _
                      "             and year(PR.receipt_date) =@Year " & vbCrLf
    
    ls_sql = ls_sql + "             and PR.Item_Code=@ItemCode " & vbCrLf & _
                      "             and PR.Seq_No=rtrim(@Optional_ReceiptSeqNo) " & vbCrLf & _
                      "     END " & vbCrLf & _
                      "     Set @Price=( " & vbCrLf & _
                      "     Select Case When @Data ='in' then " & vbCrLf & _
                      "         (Select  " & vbCrLf & _
                      "             case when  " & vbCrLf & _
                      "             Isnull(sum ( " & vbCrLf & _
                      "                 case When PR.Receipt_Cls in ('R') then isnull(Qty,0) " & vbCrLf & _
                      "                 When PR.Receipt_Cls='R1' then isnull(-Qty,0) " & vbCrLf & _
                      "                 When PR.Receipt_Cls in ('P1')  then " & vbCrLf
    
    ls_sql = ls_sql + "                     case when @ReceiptDataOnly='1' then " & vbCrLf & _
                      "                         0 " & vbCrLf & _
                      "                     else " & vbCrLf & _
                      "                         isnull(Qty,0) " & vbCrLf & _
                      "                     end " & vbCrLf & _
                      "                 end  " & vbCrLf & _
                      "             ),0)=0 then 0 " & vbCrLf & _
                      "             Else " & vbCrLf & _
                      "             Isnull(sum ( " & vbCrLf & _
                      "                 case When PR.Receipt_Cls in ('R') then isnull(Amount,0)  * dbo.UF_GetBookExchangeRate(Year(PR.Receipt_Date),Month(PR.Receipt_Date),PR.Currency_Code)  " & vbCrLf & _
                      "                 When PR.Receipt_Cls='R1' then isnull(-Amount,0) * dbo.UF_GetBookExchangeRate(Year(PR.Receipt_Date),Month(PR.Receipt_Date),PR.Currency_Code)  " & vbCrLf
    
    ls_sql = ls_sql + "                 When PR.Receipt_Cls in ('P1')  then " & vbCrLf & _
                      "                     case when @ReceiptDataOnly='1' then " & vbCrLf & _
                      "                         0 " & vbCrLf & _
                      "                     else " & vbCrLf & _
                      "                         dbo.UF_GetAverageProductionCost(@Year ,@Month ,@ItemCode,@Optional_ReceiptSeqNo,'ALL' ) " & vbCrLf & _
                      "                     end " & vbCrLf & _
                      "                 end  " & vbCrLf & _
                      "             ),0) / " & vbCrLf & _
                      "             Isnull(sum ( " & vbCrLf & _
                      "                 case When PR.Receipt_Cls in ('R') then isnull(Qty,0) " & vbCrLf & _
                      "                 When PR.Receipt_Cls='R1' then isnull(-Qty,0) " & vbCrLf
    
    ls_sql = ls_sql + "                 When PR.Receipt_Cls in ('P1')  then " & vbCrLf & _
                      "                     case when @ReceiptDataOnly='1' then " & vbCrLf & _
                      "                         0 " & vbCrLf & _
                      "                     else " & vbCrLf & _
                      "                         isnull(Qty,0) " & vbCrLf & _
                      "                     end " & vbCrLf & _
                      "                 end  " & vbCrLf & _
                      "             ),0) " & vbCrLf & _
                      "             End " & vbCrLf & _
                      "         from @tbReceipt PR group By PR.Item_Code) " & vbCrLf & _
                      "     end " & vbCrLf
    
    ls_sql = ls_sql + "     ) " & vbCrLf & _
                      "  " & vbCrLf & _
                      "     IF (isnull(@Price,0)=0) " & vbCrLf & _
                      "     BEGIN " & vbCrLf & _
                      "         Set @Price=dbo.UF_GetPreviousValuationPrice(@Year,@Month,@itemCode ) " & vbCrLf & _
                      "     END " & vbCrLf & _
                      "     Return isnull(@Price,0) " & vbCrLf & _
                      " END " & vbCrLf & _
                      "  "
    
    Db.Execute ls_sql

End Sub
