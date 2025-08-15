Attribute VB_Name = "MdlEZRunner"
Option Explicit

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Db As New ADODB.Connection
Public ErrSt As Boolean

Public IniFile As String
Public ConnStr As String, dsn As String, userName As String, Password
Public userLogin As String, StatusAdmin As Integer, UserInitPO As String

Public sql As String
Public ls_FG As String

Public Const strExcludeCustSup = "('C997','C998','C999')"

Public gs_DBName As String

' =====================================================================================================
' Add By    : Budi Subagja for Backup Database
' Add Date  : Mei 21, 2007
Public gvDBBackupServer As String
Public gvDBBackupLocal As String
' =====================================================================================================

Public Const gs_DefaultCurrencyCode = "03"

'################################################################
'#Screen Parameter
Public Const gs_formatQty = "#,##0.00"
Public Const gs_formatQtyBOM = "#,##0.00000"
Public Const gs_formatPrice = "#,##0.00000"
Public Const gs_formatPriceIDR = "#,##0.00"
Public Const gs_formatAmount = "#,##0.00000"
Public Const gs_formatAmountIDR = "#,##0.00"
Public Const gs_formatExchangeRate = "#,##0.00"

Public Const gs_formatEW = "#,##0.00"
Public Const gs_formatSW = "#,##0.00"

Public Const gs_formatDay = "###0"
Public Const gs_formatLot = "###0"
Public Const gs_formatPitch = "###0"
Public Const gs_formatProcess = "###0"
Public Const gs_formatCoefficient = "###0"
Public Const gs_formatPercentage = "#,##0.00"

Public Const gs_formatBox = "#,##0"
Public Const gs_formatVolume = "#,##0.000"
Public Const gs_formatWidth = "#,##0"
Public Const gs_formatLength = "#,##0"
Public Const gs_formatThickness = "#,##0"
Public Const gs_formatWeight = "#,##0.00"
Public Const gs_formatDrySize = "#,##0"
Public Const gs_formatReqMRP = "#,##0"

Public Const gs_formatWorkingTime = "#,##0"
Public Const gs_formatNSample = "#,##0"
Public Const gs_formatEfficiency = "#,##0.00"

Public Const gs_formatNoAju = "&&&&&&-&&&&&&-&&&&&&&&-&&&&&&"

Public Const gd_MaxQty = "9,999,999.99"
Public Const gd_MaxPrice = "999,999,999.999"
Public Const gd_MaxExchangeRate = "99,999.99"
Public Const gd_MaxAmount = "999,999,999,999,999,999.99"

Public Const gd_MaxWorkingTime = "9,999"
Public Const gd_MaxCoefficient = "99,999.99"
Public Const gd_MaxLot = "9,999"

Public Const gd_MaxBox = "9,999"
Public Const gd_MaxVolume = "9,999"
Public Const gd_MaxPitch = "999"
Public Const gd_MaxSample = "999"
Public Const gd_MaxSW = "999"
Public Const gd_MaxEW = "999"
Public Const gd_MaxPercentage = "999"
Public Const gd_MaxWidth = "9,999"
Public Const gd_MaxLength = "9,999"
Public Const gd_MaxThickness = "9,999"
Public Const gd_MaxWeight = "9,999"
Public Const gd_MaxTime = "999"



'#################################################################

'################################################################
'#Crystal Report Parameter
Public Const gi_decimalDigitQty = 2
Public Const gi_decimalDigitQtyBOM = 5
Public Const gi_decimalDigitPrice = 5
Public Const gi_decimalDigitPriceIDR = 2
Public Const gi_decimalDigitAmount = 5
Public Const gi_decimalDigitAmountIDR = 2
Public Const gi_decimalDigitExchangeRate = 2

Public Const gi_decimalDigitThickness = 0
Public Const gi_decimalDigitWidth = 0
Public Const gi_decimalDigitLength = 0

Public Const gi_decimalDigitBox = 0
Public Const gi_decimalDigitWeight = 2
Public Const gi_decimalDigitVolume = 3
'#################################################################


'print report
Public TutupPtr As Boolean, pesaninvalid As String, reportcode As String
Public do_no As String, inv_no As String, packing_no As String
Public pajak_No As String
Public invstatus As String, F_Factory As String, F_Cust_Name As String
Public zbln(11) As String, xthn(11) As String
Public sqlprint As String, ginvno As String, Fbulan As String, Ftahun As String, xbln As String
Public dfactory As String, DLine As String, dschedule1 As Date, dschedule2 As Date
Public Kt1 As String, Kt2 As String, Kt3 As String, kt4 As String, Kt5 As String, Kt6 As String, kt7 As String, kt8 As String

Public tglAwalRptPrint As String, tglAkhirRptPrint As String

Public strUnit As String, dtMPList As String, datePiList As String
Public printorient As Byte, Rqty As Double
Public MonthPre As String, MonthReceipt As String
Public MonthSupply As String, MonthCurrent As String
Public sqlprint2 As String, sqlprint3 As String, xdays(30) As String * 2

Public PacNo As String
Public InvNN As String
Public serverPath As String

Public kirimPar As String
Public Const strAll As String = "ALL"

Public Function DisplayMsg(n As String)
    Dim StrErr
    Dim RS As New ADODB.Recordset
    Dim sql

    sql = "select * from message where MsgId=" & n
    Set RS = Db.Execute(sql)
    If Not (RS.EOF And RS.BOF) Then
        StrErr = "[" & CStr(Trim(RS("MsgId"))) & "]  " & Trim(RS("MsgDesc"))
        DisplayMsg = StrErr
        Set RS = Nothing
        ErrSt = True
    End If
End Function

Public Function hakAkses(mn As String, Optional desc As String) As Integer
    Dim sql As String
    Dim RS As New ADODB.Recordset
    
    sql = "select status from user_Privilege where App_ID = 'P01' " & _
        "and Menu_ID = (select Menu_ID from User_Menu where Menu_Name ='" & mn & "'"
        
    If desc <> "" Then
        sql = sql & " and Menu_Desc ='" & desc & "'"
    End If
    
    sql = sql & ") and userName ='" & userLogin & "'"
    
    Set RS = Db.Execute(sql)
    If Not (RS.EOF And RS.BOF) Then
        hakAkses = RS("status")
    End If
End Function

Public Function hakUpdate(mn As String, Optional desc As String) As Integer
    Dim sql As String
    Dim RS As New ADODB.Recordset
    
    sql = "select Allow_Update from user_Privilege where App_ID = 'P01' " & _
        "and Menu_ID = (select Menu_ID from User_Menu where Menu_Name ='" & mn & "'"
        
    If desc <> "" Then
        sql = sql & " and Menu_Desc ='" & desc & "'"
    End If
    
    sql = sql & ") and userName ='" & userLogin & "'"
    
    Set RS = Db.Execute(sql)
    If Not (RS.EOF And RS.BOF) Then
        hakUpdate = RS("Allow_Update")
    End If
End Function

Public Function hakPrice(mn As String, Optional desc As String) As Integer
    Dim sql As String
    Dim RS As New ADODB.Recordset
    
    sql = "select ISNULL(Allow_Price,0) Allow_Price from user_Privilege where App_ID = 'P01' " & _
        "and Menu_ID = (select Menu_ID from User_Menu where Menu_Name ='" & mn & "'"
        
    If desc <> "" Then
        sql = sql & " and Menu_Desc ='" & desc & "'"
    End If
    
    sql = sql & ") and userName ='" & userLogin & "'"
    
    Set RS = Db.Execute(sql)
    If Not (RS.EOF And RS.BOF) Then
        hakPrice = RS("Allow_Price")
    End If
End Function

Public Function frmcode(frmName$, Optional Ket As String)
    Dim sql As String
    Dim rst As Recordset
    sql = "select * from user_menu where menu_name='" & frmName & "'"
    If Ket <> "" Then
        sql = sql & " and menu_desc='" & Ket & "'"
    End If
    Set rst = New Recordset
    rst.Open sql, Db, adOpenKeyset, adLockOptimistic
    If rst.EOF = False Then frmcode = Trim(rst!menu_id)
End Function

Public Function panggilForm(nmMenu As String, Optional nmForm As String) As Integer
Dim RS As New Recordset

    
    sql = " Select A.* From  " & vbCrLf & _
                " ( " & vbCrLf & _
                " select Menu_Name,Menu_Desc from user_menu where menu_ID='" & nmMenu & "'  " & vbCrLf & _
                " )a Inner join User_Privilege b on B.Menu_ID='" & nmMenu & "' and b.UserName='" & userLogin & "' " & vbCrLf & _
                "  "
    Set RS = Db.Execute(sql)
    
    If RS.EOF And RS.BOF Then
        panggilForm = 1 'msg Error "Menu ID not Found"
        Exit Function
    Else
        If (LCase(Trim(RS(0))) = LCase(Trim(nmForm))) Then panggilForm = 2: Exit Function
        If Trim(nmForm) <> "" Then Call hideThisForm(nmForm)
        If frmMainMenu.Tree.Nodes.Count < 1 Then frmMainMenu.loadtree
        If frmLogin.Visible = True Then frmLogin.Hide
        LoadDynamicForm (Trim$(RS("menu_Name") & ""))
    End If
End Function

'***** Filter Combo Unit ******
Public Sub filterCboUnit(UnitCls As String, nmCombo, Optional textKol As Integer)
'    Dim ls_unit As String
'    ls_unit = ""
'    Dim lrs As New ADODB.Recordset
'    If lrs.State <> adStateClosed Then lrs.Close
'    lrs.Open "select * from unit_cls", Db, adOpenKeyset, adLockOptimistic
'    While lrs.EOF = False
'        If Trim(ls_unit) = "" Then
'            ls_unit = ls_unit & Trim(lrs!Description)
'        Else
'            ls_unit = ls_unit & "," & Trim(lrs!Description)
'        End If
'        lrs.MoveNext
'    Wend
'    If lrs.State <> adStateClosed Then lrs.Close
    
    Call up_FillCombo(nmCombo, "unit_cls")
    
End Sub

Function InvReport() As String
Dim application As New CRAXDDRT.application
Dim report As New CRAXDDRT.report
Dim rs1 As Recordset, rs2 As Recordset, xdo As String
Dim temp As String, do1 As String, do2 As String, Y As Integer
Dim tbank As String, tadd As String, cbol As Boolean, xy As Integer
      
'    sql = "Select '" & bteHakPrice & "' HakPrice, RTRim(ivm.Invoice_No) Invoice_No, ivm.Invoice_Date, RTRim(ivm.Remarks) Remarks,  RTRim(ivm.List_PO) List_PO, " & _
'        vbLf & "RTRim(ivd.Item_Code) Item_Code, RTrim(ivd.MakerItem_Code) MakerItem_Code, ivd.Qty, ivd.Price, ivd.Amount, RTrim(ivd.PO_No) PO_No, " & _
'        vbLf & "RTrim(im.Item_Name) Item_Name, im.Group_Cls, RTrim(cc.Description) Curr_Desc, RTrim(uc.Description) Unit_Desc, " & _
'        vbLf & "RTrim(cp.Company_Name) Company_Name, RTrim(cp.Address1) Address1, RTrim(cp.Address2) Address2, RTrim(cp.Province) Province, " & _
'        vbLf & "RTrim(cp.City) City, RTrim(cp.Postal_Code) Postal_Code, RTrim(cp.Phone1) Phone1, RTrim(cp.Phone2) Phone2, RTrim(cp.Fax) Fax, " & _
'        vbLf & "RTrim(cp.SJ_Position) Prepared_Position, RTrim(cp.SJ_Person) Prepared_Person, RTrim(cp.Invoice_Position) Checked_Position, " & _
'        vbLf & "RTrim(cp.Invoice_Person) Checked_Person, RTrim(cp.PO_Position) Approved_Position, RTrim(cp.PO_Person) Approved_Person, " & _
'        vbLf & "RTrim(tm.Trade_Name) Cust_Name, RTrim(tm.Address1) Cust_Address1, RTrim(tm.Address2) Cust_Address2, RTrim(tm.City) Cust_City, " & _
'        vbLf & "RTrim(tm.Contact_Person) Contact_Person, Isnull(pc.Description, 'FOB') Description, om.NoCommercial_Cls " & _
'        vbLf & "From Invoice_Master ivm " & _
'        vbLf & "Inner Join Invoice_Detail ivd On ivm.Invoice_No = ivd.Invoice_No " & _
'        vbLf & "Inner Join  Delivery_Order do On do.DO_NO = ivd.DO_No and do.DOSeq_No = ivd.DOSeq_No and do.PO_No = ivd.PO_No and do.Seq_No = ivd.Seq_No " & _
'        vbLf & "Inner Join DO_Master dm On dm.DO_No = do.DO_No " & _
'        vbLf & "Inner Join OrderEntry_Detail od On od.PO_No = do.PO_No and od.seq_no = do.seq_no " & _
'        vbLf & "Inner Join OrderEntry_Master om On om.PO_No = od.PO_No and om.Cust_Code = od.cust_code " & _
'        vbLf & "Inner Join Item_Master im On ivd.Item_Code = im.Item_Code " & _
'        vbLf & "Inner Join Trade_Master tm On ivm.Cust_Code = tm.Trade_Code " & _
'        vbLf & "Left Outer Join PriceCondition_cls pc on ivm.tradeterms_cls  = pc.pricecondition_cls " & _
'        vbLf & "Left Outer Join Unit_Cls uc On ivd.Unit_Cls = uc.Unit_Cls " & _
'        vbLf & "Left Outer Join Curr_Cls cc On ivd.Currency_Code = cc.Curr_Cls, Company_Profile cp " & _
'        vbLf & "Where ivm.Invoice_No In (" & inv_no & ") " & _
'        vbLf & "Order By ivm.Invoice_No, ivd.MakerItem_Code"
      
sql = "select xxx.*,xxx.invoice_to, rtrim(yyy.trade_name) as Nama1,rtrim(yyy.address1) as add1, " & _
      vbLf & "rtrim(yyy.address2) as ToAdd2 ,sf.*,rtrim(yyy.city) city1, rtrim(yyy.postal_code) postal_code1 " & _
      vbLf & ",DO_NO,Do_Date from ( " & _
      vbLf & "select tm.trade_code,rtrim(tm.trade_name) trade_name, rtrim(tm.invoice_to) invoice_to, " & _
      vbLf & "rtrim(tm.address1) address1,rtrim(tm.address2) address2, " & _
      vbLf & "rtrim(tm.city) city, rtrim(tm.postal_code) postal_code,ivd.invoice_no,ivm.invoice_date, " & _
      vbLf & "right(ivm.delivery_date,2)  bulan, left(ivm.delivery_date,4) thn, " & _
      vbLf & "ivd.item_code,rtrim(isnull(ivd.makeritem_code,'')) part_no,rtrim(im.item_name) item_name, " & _
      vbLf & "sum(ivd.qty) qty ,ivd.unit_cls,ivd.currency_code,ivd.price,sum(ivd.amount) amount, " & _
      vbLf & "ivm.amount invamount, ivm.ppn, ivm.total_amount,ivd.seq_no, " & _
      vbLf & "rtrim(ivm.list_do) list_do, rtrim(ivm.list_po) list_po, " & _
      vbLf & "rtrim(cp.company_name) company_name, rtrim(cp.address1) cpaddress1, " & _
      vbLf & "rtrim(cp.address2) cpaddress2,rtrim(cp.Province) cpProvince, rtrim(cp.City) cpcity, " & _
      vbLf & "rtrim(cp.postal_code) cpPostal_code, " & _
      vbLf & "rtrim(cp.phone1) cpphone1, rtrim(cp.phone2) cpphone2, rtrim(cp.fax) cpfax, " & _
      vbLf & "rtrim(invoice_position) invpos, rtrim(invoice_person) invperson,tm.country_cls,ivd.service ss,ivm.Remarks  "

sql = sql & _
      vbLf & ",IVD.DO_NO,dm.Do_date,IVM.Due_date from invoice_detail ivd " & _
      vbLf & "inner join invoice_master ivm " & _
      vbLf & "on ivd.invoice_no = ivm.invoice_no " & _
      vbLf & "inner join do_master dm " & _
      vbLf & "on dm.do_no = ivd.do_no " & _
      vbLf & "inner join trade_master tm " & _
      vbLf & "on dm.cust_code = tm.trade_code " & _
      vbLf & "inner join Item_Master im " & _
      vbLf & "on im.item_code = ivd.item_code, " & _
      vbLf & "company_profile cp " & _
      vbLf & "Where ivd.invoice_no in (" & inv_no & " ) " & _
      vbLf & "Group By tm.trade_code,tm.trade_name ,tm.address1 ,tm.address2,tm.city, tm.postal_code, " & _
      vbLf & "ivd.invoice_no,ivm.invoice_date,ivm.delivery_date,ivd.price, ivd.item_code, ivd.makeritem_code,im.item_name, " & _
      vbLf & "ivd.unit_cls , ivd.currency_code, ivm.amount, ivm.ppn, ivm.total_amount, ivm.list_do, ivm.list_po, " & _
      vbLf & "cp.company_name, cp.address1,cp.address2,cp.Province, cp.City , cp.Postal_code, " & _
      vbLf & "cp.phone1 , cp.phone2, cp.fax, invoice_position, invoice_person, tm.invoice_to, tm.country_cls " & _
      vbLf & ",IVD.DO_NO,dm.Do_date,ivd.service,ivm.Remarks,ivm.Due_date ,ivd.seq_no ) xxx  INNER JOIN(SELECT SUM(Price*qty)SumPrice,Sum(SErvice*Qty)SumService," & _
      vbLf & " SubTotal=SUM(Price*qty)+Sum(SErvice*Qty),seq_No,DO_NO as DO_Number," & _
      vbLf & " Invoice_No as INVNO FROM " & _
      vbLf & " Invoice_Detail Group BY INvoice_NO,Do_nO,currency_code,seq_NO) Sf ON Sf.seq_No=xxx.seq_No AND sf.seq_No=xxx.seq_No" & _
      vbLf & "  and SF.DO_Number=xxx.do_NO AND sf.invno=xxx.invoice_no" & _
      vbLf & "left join " & _
      vbLf & "(select * from trade_master ) yyy on xxx.invoice_to = yyy.trade_code " & _
      vbLf & "order by item_code "

    Set rs1 = New Recordset
    rs1.CursorLocation = adUseClient
    rs1.Open sql, Db, adOpenKeyset, adLockOptimistic
    If Not rs1.EOF Then
        Set report = application.OpenReport(App.path & "\REPORTs\invoice.rpt")
        report.Database.Tables(1).SetDataSource rs1
        Dim Rpt As New FrmRpt3
        reportcode = "invoice"
        printorient = 1
        sqlprint = sql
        
        sql = "select cb.*, cc.Description curr " & _
              vbLf & "from Company_bank cb " & _
              vbLf & "left join Curr_Cls cc " & _
              vbLf & "on cb.Currency_Code = cc.Curr_Cls " & _
              vbLf & "order by cb.bank_name, cb.address1, cb.address2, cb.currency_code "
        
        Set rs2 = New Recordset
        rs2.Open sql, Db, adOpenKeyset, adLockOptimistic
       
        If Not rs2.EOF Then
            For Y = 0 To rs2.RecordCount - 1
                If Y = 0 Then
                    cbol = True
                    tbank = Trim(rs2!bank_name) & Trim(rs2!address1) & Trim(rs2!address2)
                    report.FormulaFields.GetItemByName("Bank1").Text = "'" & Trim(rs2!bank_name) & "'"
                    report.FormulaFields.GetItemByName("Bank1Add1").Text = "'" & Trim(rs2!address1) & "'"
                    report.FormulaFields.GetItemByName("Bank1Add2").Text = "'" & Trim(rs2!address2) & "'"
                    report.FormulaFields.GetItemByName("Bank1City").Text = "'" & Trim(rs2!City) & "'"
                    report.FormulaFields.GetItemByName("Bank1KodePos").Text = "'" & Trim(rs2!postal_code) & "'"
                    report.FormulaFields.GetItemByName("Bank1Acc1").Text = "'" & Trim(rs2!Curr) & " " & Trim(rs2!Account_No) & "'"
                    xy = 1
                Else
                    If tbank = Trim(rs2!bank_name) & Trim(rs2!address1) & Trim(rs2!address2) Then
                       If xy < 7 Then
                        xy = xy + 1 'Account max s/d 7
                        If cbol Then
                            report.FormulaFields.GetItemByName("Bank1Acc" & xy).Text = "'" & Trim(rs2!Curr) & " " & Trim(rs2!Account_No) & "'"
                        Else
                            report.FormulaFields.GetItemByName("Bank2Acc" & xy).Text = "'" & Trim(rs2!Curr) & " " & Trim(rs2!Account_No) & "'"
                        End If
                      End If
                    Else
                        If cbol = False Then Exit For
                        cbol = False 'Utk pembatas jumlah (nama_bank + addr1 + addr2) yg beda sebyk 2
                        xy = 1
                        tbank = Trim(rs2!bank_name) & Trim(rs2!address1) & Trim(rs2!address2)
                        report.FormulaFields.GetItemByName("Bank2").Text = "'" & Trim(rs2!bank_name) & "'"
                        report.FormulaFields.GetItemByName("Bank2Add1").Text = "'" & Trim(rs2!address1) & "'"
                        report.FormulaFields.GetItemByName("Bank2Add2").Text = "'" & Trim(rs2!address2) & "'"
                        report.FormulaFields.GetItemByName("Bank2City").Text = "'" & Trim(rs2!City) & "'"
                        report.FormulaFields.GetItemByName("Bank2KodePos").Text = "'" & Trim(rs2!postal_code) & "'"
                        report.FormulaFields.GetItemByName("Bank2Acc1").Text = "'" & Trim(rs2!Curr) & " " & Trim(rs2!Account_No) & "'"
                    End If
                End If
                rs2.MoveNext
            Next
        End If
        
        report.FormulaFields.GetItemByName("DecimalQty").Text = gi_decimalDigitQty
        report.FormulaFields.GetItemByName("DecimalPrice").Text = gi_decimalDigitPrice
        report.FormulaFields.GetItemByName("DecimalPriceIDR").Text = gi_decimalDigitPriceIDR
        report.FormulaFields.GetItemByName("DecimalAmount").Text = gi_decimalDigitAmount
        report.FormulaFields.GetItemByName("DecimalAmountIDR").Text = gi_decimalDigitAmountIDR
        
        With Rpt.CRViewer1
            .ReportSource = report
            .ViewReport
            .Zoom 1
        End With
        With Rpt
            .WindowState = 2
            .Show 1
        End With
    End If
 
End Function

Sub InvReportExport(bteHakPrice As Byte)
    
    Dim application As New CRAXDDRT.application
    Dim report As New CRAXDDRT.report
    Dim rsMain As New Recordset
    Dim rsSub As New Recordset
    Dim RsSubDetail As New Recordset
    
'    Sql = "Select '1' HakPrice, RTRim(ivm.Invoice_No) Invoice_No, ivm.Invoice_Date, " & _
'          vbLf & "RTrim(pm.List_PO) List_PO, RTrim(pm.List_PODate) List_PODate, RTrim(ivm.Remarks) Remarks, " & _
'          vbLf & "RTRim(ivd.Item_Code) Item_Code, RTrim(ivd.MakerItem_Code) MakerItem_Code, " & _
'          vbLf & "RTrim(im.Item_Name) Item_Name, ivd.Qty, " & _
'          vbLf & "ivd.Price, ivd.Amount, RTrim(ivd.PO_No) PO_No, " & _
'          vbLf & "RTrim(ivd.Currency_Code) CurrCls, " & _
'          vbLf & "Ex_Rate = Case When ivd.Currency_Code = '03' Then 1 Else IsNull( " & _
'          vbLf & "(Select Daily_ExchangeRate From Daily_ExchangeRate Where ExchangeRate_Date = ivm.Invoice_Date And Currency_Code = ivd.Currency_Code), 0) End, " & _
'          vbLf & "RTrim(uc.Description) Unit_Desc, RTrim(cc.Description) Curr_Desc, im.Group_Cls, " & _
'          vbLf & "RTrim(pc.Description) TradeTermDesc, om.NoCommercial_Cls, " & _
'          vbLf & "pm.ETA, pm.ETD, RTrim(pm.Vessel) FeederVessel, RTrim(pm.Mother_Vessel) MotherVessel, " & _
'          vbLf & "RTrim(tm.Trade_Name) ConsigneeName, RTrim(tm.Address1) ConsigneeAddr1, " & _
'          vbLf & "RTrim(tm.Address2) ConsigneeAddr2, RTrim(tm.City) ConsgineeCity, " & _
'          vbLf & "RTrim(tm.Postal_Code) ConsigneePostalCode, " & _
'          vbLf & "RTrim(tm.Telephone) ConsigneeTelp, RTrim(tm.Fax) ConsigneeFax, " & _
'          vbLf & "RTrim(tm.Contact_Person) ConsigneePerson, "
'    Sql = Sql & _
'          vbLf & "RTrim(pm.From_Port) From_Port, RTrim(pm.To_Port) To_Port, " & _
'          vbLf & "RTrim(pm.Country_Origin) Country_Origin, RTrim(Final_Destination) Final_Destination, " & _
'          vbLf & "RTrim(pm.payment_terms) PaymentTerms, RTrim(ptc.Description) PaymentTermsDesc, " & _
'          vbLf & "RTrim(pm.POCaseMark1) POCaseMark1, RTrim(pm.POCaseMark2) POCaseMark2, " & _
'          vbLf & "RTrim(pm.POCaseMark3) POCaseMark3, RTrim(pm.POCaseMark4) POCaseMark4, RTrim(pm.POCaseMark5) POCaseMark5, " & _
'          vbLf & "pm.PackingStyle_Cls, pm.Transportation_Cls, " & _
'          vbLf & "RTrim(cp.Company_Name) Company_Name, RTrim(cp.Address1) Address1, RTrim(cp.Address2) Address2, " & _
'          vbLf & "RTrim(cp.Province) Province, RTrim(cp.City) City, RTrim(cp.Postal_Code) Postal_Code, " & _
'          vbLf & "RTrim(cp.Phone1) Phone1, RTrim(cp.Phone2) Phone2, RTrim(cp.Fax) Fax, " & _
'          vbLf & "RTrim(cp.DI_Position) Prepared_Position, RTrim(cp.DI_Person) Prepared_Person, " & _
'          vbLf & "RTrim(cp.Invoice_Person) Checked_Person, RTrim(cp.Invoice_Position) Checked_Position, " & _
'          vbLf & "RTrim(cp.PO_Position) Approved_Position, RTrim(cp.PO_Person) Approved_Person "
'    Sql = Sql & _
'          vbLf & "From Invoice_Master ivm " & _
'          vbLf & "Inner Join Invoice_Detail ivd On ivm.Invoice_No = ivd.Invoice_No " & _
'          vbLf & "Inner Join Item_Master im On ivd.Item_Code = im.Item_Code " & _
'          vbLf & "Inner Join Packing_Master pm On ivd.Packing_No = pm.Packing_No " & _
'          vbLf & "Inner Join Packing_Detail pd On ivd.Packing_No = pd.Packing_No And ivd.PackingSeq_No = pd.PackingSeq_No And ivd.PO_No = pd.Order_No And ivd.Seq_No = pd.Order_SeqNo " & _
'          vbLf & "Inner Join OrderEntry_Detail od On od.PO_No = pd.Order_No And od.Seq_No = pd.order_SeqNo " & _
'          vbLf & "Inner Join OrderEntry_Master om On om.PO_No = od.PO_No And om.cust_code = od.cust_code " & _
'          vbLf & "Inner Join Trade_Master tm On pm.Consignee = tm.Trade_Code " & _
'          vbLf & "Left Join PriceCondition_Cls pc On pc.pricecondition_cls = ivm.tradeterms_cls " & _
'          vbLf & "Left Join Unit_Cls uc On ivd.Unit_Cls = uc.Unit_Cls " & _
'          vbLf & "Left Join Curr_Cls cc On ivd.Currency_Code = cc.Curr_Cls " & _
'          vbLf & "Left Join paymentterm_cls ptc on ptc.PaymentTerm_Cls = pm.payment_terms, " & _
'          vbLf & "Company_Profile cp " & _
'          vbLf & "Where ivm.Invoice_No In (" & Trim(inv_no) & ") " & _
'          vbLf & "Order By ivm.Invoice_No, ivd.Item_Code "
    
sql = "  Select PM.Packing_No,PM.Packing_Date,PM.Stuffing_Date,PM.ETD,PM.ETA,PM.Payment_Days,PM.Payment,  " & vbCrLf & _
            "      PM.Transportation_Cls,PM.Vessel,PM.Mother_Vessel,PM.From_Port,PM.To_Port,PM.Final_Destination,  " & vbCrLf & _
            "      PM.Remarks,PM.Last_User, PM.POCaseMark1,PM.POCaseMark2, " & vbCrLf & _
            "      PM.Cust_Code,TM.Trade_Name Customer_Name,TM.Address1 CsAddress1,TM.Address2 CsAddress2,TM.City CsCity, TM.Country CsCountry,  " & vbCrLf & _
            "      PM.Consignee,PM.ConsigneeTitle,TM1.Trade_Name Consignee_Name,TM1.Address1 CgAddress1,TM1.Address2 CgAddress2,TM1.City CgCity, TM1.Country CgCountry,  " & vbCrLf & _
            "      PM.Payment_Terms,Isnull(PT.Description,'') Payment_Desription, UC.Description Units, " & vbCrLf & _
            "      PD.Order_No,PD.Container_No,PD.Container_Size,PD.SerialNoFrom,PD.SerialNoTo,PD.Qty,PD.Unit_Cls,  " & vbCrLf & _
            "      PD.Item_Code,IM.Item_Name, PD.Price, PD.Currency_Code,CC.Description Currency,PD.Amount  " & vbCrLf & _
            "      From Packing_Master PM  " & vbCrLf & _
            "      Inner Join Packing_Detail PD On PM.Packing_No=PD.Packing_No       " & vbCrLf & _
            "    Inner Join Trade_Master TM on PM.Cust_Code=TM.Trade_Code  "

sql = sql + "      Inner Join Trade_Master TM1 on PM.Consignee=TM1.Trade_Code  " & vbCrLf & _
            "      Inner Join Item_Master IM on PD.Item_Code=IM.Item_Code  " & vbCrLf & _
            "    Inner Join Curr_Cls CC on PD.Currency_Code=CC.Curr_Cls " & vbCrLf & _
            "    Left Join Unit_Cls UC on PD.Unit_Cls=UC.Unit_Cls " & vbCrLf & _
            "      Left Join PaymentTerm_Cls PT on PM.Payment_Terms=PT.PaymentTerm_Cls  " & vbCrLf & _
            "      Where PM.Packing_No=" & Trim(inv_no) & "  " & vbCrLf & _
            "      Order By PM.Packing_No, PD.PackingSeq_No, PD.Item_Code "
    
    
    sqlprint = sql
    sqlprint2 = sql
    
    rsMain.CursorLocation = adUseClient
    If rsMain.State <> adStateClosed Then rsMain.Close
    rsMain.Open sql, Db, adOpenKeyset, adLockOptimistic
    
    If Not rsMain.EOF Then
        Set report = application.OpenReport(App.path & "\Reports\invoice_KWI.rpt")
        report.Database.Tables(1).SetDataSource rsMain
        report.DiscardSavedData
        
'        report.FormulaFields.GetItemByName("DecimalQty").Text = gi_decimalDigitQty
'        report.FormulaFields.GetItemByName("DecimalPrice").Text = gi_decimalDigitPrice
'        report.FormulaFields.GetItemByName("DecimalPriceIDR").Text = gi_decimalDigitPriceIDR
'        report.FormulaFields.GetItemByName("DecimalAmount").Text = gi_decimalDigitAmountIDR
        
        rsSub.CursorLocation = adUseClient
        If rsSub.State <> adStateClosed Then rsSub.Close
        rsSub.Open sql, Db, adOpenKeyset, adLockOptimistic
              
'        With report.OpenSubreport("sub_packing_list_ex") 'Sub Invoice
'         .Database.Tables(1).SetDataSource rsSub
'         .FormulaFields.GetItemByName("SubDecimalQty").Text = gi_decimalDigitQty
'         .FormulaFields.GetItemByName("SubDecimalPrice").Text = gi_decimalDigitPrice
'         .FormulaFields.GetItemByName("subDecimalPriceIDR").Text = gi_decimalDigitPriceIDR
'         .FormulaFields.GetItemByName("subDecimalAmount").Text = gi_decimalDigitAmountIDR
'        End With
        
'        Sql = "Select '1' HakPrice, RTRim(ivm.Invoice_No) Invoice_No, ivm.Invoice_Date, " & _
'              vbLf & "RTrim(pm.List_PO) List_PO, RTrim(pm.List_PODate) List_PODate, RTrim(ivm.Remarks) Remarks, " & _
'              vbLf & "RTrim(pd.Container_No) Container_No, " & _
'              vbLf & "RTRim(ivd.Item_Code) Item_Code, RTrim(ivd.MakerItem_Code) MakerItem_Code, " & _
'              vbLf & "RTrim(im.Item_Name) Item_Name, ivd.Qty, " & _
'              vbLf & "ivd.Price, ivd.Amount, RTrim(ivd.PO_No) PO_No, " & _
'              vbLf & "RTrim(ivd.Currency_Code) CurrCls, " & _
'              vbLf & "Ex_Rate = Case When ivd.Currency_Code = '03' Then 1 Else IsNull( " & _
'              vbLf & "(Select Daily_ExchangeRate From Daily_ExchangeRate Where ExchangeRate_Date = ivm.Invoice_Date And Currency_Code = ivd.Currency_Code), 0) End, " & _
'              vbLf & "RTrim(uc.Description) Unit_Desc, RTrim(cc.Description) Curr_Desc, im.Group_Cls, " & _
'              vbLf & "RTrim(pc.Description) TradeTermDesc "
'        Sql = Sql & _
'              vbLf & "From Invoice_Master ivm " & _
'              vbLf & "Inner Join Invoice_Detail ivd On ivm.Invoice_No = ivd.Invoice_No " & _
'              vbLf & "Inner Join Item_Master im On ivd.Item_Code = im.Item_Code " & _
'              vbLf & "Inner Join Packing_Master pm On ivd.Packing_No = pm.Packing_No " & _
'              vbLf & "Inner Join Packing_Detail pd On ivd.Packing_No = pd.Packing_No And ivd.PackingSeq_No = pd.PackingSeq_No And ivd.PO_No = pd.Order_No And ivd.Seq_No = pd.Order_SeqNo " & _
'              vbLf & "Inner Join OrderEntry_Detail od On od.PO_No = pd.Order_No And od.Seq_No = pd.order_SeqNo " & _
'              vbLf & "Inner Join OrderEntry_Master om On om.PO_No = od.PO_No And om.cust_code = od.cust_code " & _
'              vbLf & "Left Join PriceCondition_Cls pc On pc.pricecondition_cls = ivm.tradeterms_cls " & _
'              vbLf & "Left Join Unit_Cls uc On ivd.Unit_Cls = uc.Unit_Cls " & _
'              vbLf & "Left Join Curr_Cls cc On ivd.Currency_Code = cc.Curr_Cls, " & _
'              vbLf & "Company_Profile cp " & _
'              vbLf & "Where ivm.Invoice_No In (" & Trim(inv_no) & ") " & _
'              vbLf & "Order By ivm.Invoice_No, ivd.Item_Code "
'
'        sqlprint3 = Sql
'        rsSubDetail.CursorLocation = adUseClient
'        If rsSubDetail.State <> adStateClosed Then rsSubDetail.Close
'        rsSubDetail.Open Sql, Db, adOpenKeyset, adLockOptimistic
'
'        If Not rsSubDetail.EOF Then
'         With report.OpenSubreport("subDetailOf_packing_list_ex")
'          .Database.Tables(1).SetDataSource rsSubDetail
'          .FormulaFields.GetItemByName("SubDetailDecimalQty").Text = gi_decimalDigitQty
'          .FormulaFields.GetItemByName("SubDetailDecimalPrice").Text = gi_decimalDigitPrice
'          .FormulaFields.GetItemByName("SubDetailDecimalPriceIDR").Text = gi_decimalDigitPriceIDR
'          .FormulaFields.GetItemByName("SubDetailDecimalAmount").Text = gi_decimalDigitAmountIDR
'         End With
'        End If
        
        Dim Rpt As New FrmRpt3
        'reportcode = "invoice_ex"
        reportcode = "invoice_kwi"
        printorient = 1
        With Rpt.CRViewer1
            .ReportSource = report
            .ViewReport
            .Zoom 1
        End With
        With Rpt
            .WindowState = 2
            .Show 1
        End With
    
    End If
    
End Sub

Function DOPrintStatus(Dono$) As Boolean

    Dim rstdo As Recordset
    sql = "select * from DO_master where DO_no in (" & Dono & ") and fix_cls = '1'"
    Set rstdo = New Recordset
    rstdo.Open sql, Db, adOpenDynamic, adLockOptimistic
    If rstdo.EOF Then
        pesaninvalid = "DO " & Dono & " is not fixed!"
    Else
        pesaninvalid = ""
    End If

End Function

Function DOReport(Dono$) As String
Dim application As New CRAXDDRT.application
Dim report As New CRAXDDRT.report
Dim rs1 As Recordset
   
'    Sql = "select d.trade_code,rtrim(d.trade_name) trade_name,rtrim(d.address1) address1,rtrim(d.address2) address2, " & _
'          vbLf & "rtrim(d.city) city, rtrim(d.postal_code) postal_code,a.do_no,b.do_date,rtrim(a.po_no) po_no,a.delivery_date, " & _
'          vbLf & "a.item_code, rtrim(c.item_name) item_name,rtrim(a.makeritem_code) makeritem_code, " & _
'          vbLf & "rtrim(a.lot_no) lot_no,C.number_entering , A.qty, A.unit_cls, rtrim(A.SerialNoFrom) SFrom,rtrim(A.SerialNoTo) STo,rtrim(isnull(b.list_po,'')) list_po " & _
'          vbLf & ",rtrim(f.company_name) company_name, rtrim(f.address1) cpaddress1, " & _
'          vbLf & "rtrim(f.address2) cpaddress2,rtrim(f.Province) cpProvince, rtrim(f.City) cpcity, rtrim(f.postal_code) cpPostal_code, " & _
'          vbLf & "rtrim(f.phone1) cpphone1, rtrim(f.phone2) cpphone2, rtrim(f.fax) cpfax,rtrim(sj_position) sjpos, rtrim(sj_person) sjperson " & _
'          vbLf & "from delivery_order a,do_master b, item_master c,trade_master d ,company_profile f " & _
'          vbLf & "Where B.do_no = A.do_no and b.cust_code = d.trade_code and a.item_code = c.item_code " & _
'          vbLf & "and b.do_no in (" & Dono & ") order by Delivery_Date,a.Item_Code,Lot_No,a.PO_NO,Seq_No,DOSeq_No "
          
' Update report for KAWAI with change Customer Code to Location Code ( Information From OrderEntry_Master ) -- 20090420
    
sql = " select d.trade_code,rtrim(d.trade_name) trade_name,rtrim(d.address1) address1,rtrim(d.address2) address2,  " & vbCrLf & _
            " rtrim(d.city) city, rtrim(d.postal_code) postal_code,a.do_no,b.do_date,rtrim(a.po_no) po_no,a.delivery_date,  " & vbCrLf & _
            " a.item_code, rtrim(c.item_name) item_name,rtrim(a.makeritem_code) makeritem_code,  " & vbCrLf & _
            " rtrim(a.lot_no) lot_no,C.number_entering , A.qty, A.unit_cls, rtrim(A.SerialNoFrom) SFrom,rtrim(A.SerialNoTo) STo,rtrim(isnull(b.list_po,'')) list_po  " & vbCrLf & _
            " ,rtrim(f.company_name) company_name, rtrim(f.address1) cpaddress1,  " & vbCrLf & _
            " rtrim(f.address2) cpaddress2,rtrim(f.Province) cpProvince, rtrim(f.City) cpcity, rtrim(f.postal_code) cpPostal_code,  " & vbCrLf & _
            " rtrim(f.phone1) cpphone1, rtrim(f.phone2) cpphone2, rtrim(f.fax) cpfax,rtrim(sj_position) sjpos, rtrim(sj_person) sjperson  " & vbCrLf & _
            " from delivery_order a,do_master b, item_master c,trade_master d ,OrderEntry_Master P,company_profile f  " & vbCrLf & _
            " Where B.do_no = A.do_no and a.PO_No=p.Po_No and p.location_code = d.trade_code and a.item_code = c.item_code  " & vbCrLf & _
            " and b.do_no in (" & Dono & ") order by Delivery_Date,a.Item_Code,Lot_No,a.PO_NO,Seq_No,DOSeq_No  "
    
    
    Set rs1 = New Recordset

    rs1.CursorLocation = adUseClient
    rs1.Open sql, Db, adOpenKeyset, adLockOptimistic
    If Not rs1.EOF Then
        
        Set report = application.OpenReport(App.path & "\REPORTs\Delivery_Order.rpt")
        report.Database.Tables(1).SetDataSource rs1
        
        '#####################################################################
        '# Qty Digit and decimal
        report.FormulaFields.GetItemByName("DecimalBox").Text = "" & gi_decimalDigitBox & ""
        report.FormulaFields.GetItemByName("DecimalQty").Text = "" & gi_decimalDigitQty & ""
        '#####################################################################
        
        Dim Rpt As New FrmRpt3
        sqlprint = sql
        reportcode = "DO"
        printorient = 1
        do_no = Dono
        With Rpt.CRViewer1
            .ReportSource = report
            .ViewReport
            .Zoom 1
        End With
        With Rpt
            .WindowState = 2
            .Show 1
        End With
    End If

End Function
Function DIReport(Dono$) As String
Dim application As New CRAXDDRT.application
Dim report As New CRAXDDRT.report
Dim rs1 As Recordset
   
    sql = "select d.trade_code,rtrim(d.trade_name) trade_name,rtrim(d.address1) address1,rtrim(d.address2) address2, " & vbCrLf & _
            "rtrim(d.city) city, rtrim(d.postal_code) postal_code,a.do_no,b.do_date,rtrim(a.po_no) po_no,a.delivery_date, a.item_code,rtrim(c.item_name) item_name, " & vbCrLf & _
            "c.number_entering , A.qty, A.unit_cls,rtrim(a.SerialNoFrom) SFrom, rtrim(A.SerialNoto) STo, e.wh_code, rtrim(b.remarks) note,rtrim(isnull(b.list_po,'')) list_po " & vbCrLf & _
            ",rtrim(f.company_name) company_name, rtrim(f.address1) cpaddress1, " & vbCrLf & _
            "rtrim(f.address2) cpaddress2,rtrim(f.Province) cpProvince, rtrim(f.City) cpcity, rtrim(f.postal_code) cpPostal_code, " & vbCrLf & _
            "rtrim(f.phone1) cpphone1, rtrim(f.phone2) cpphone2, rtrim(f.fax) cpfax, rtrim(DI_position) DIpos, rtrim(DI_person) DIperson, Lot_No " & _
            "from delivery_order a,do_master b, item_master c,trade_master d, warehouse_master e, company_profile f  " & vbCrLf & _
            "Where B.do_no = A.do_no and b.cust_code = d.trade_code and a.item_code = c.item_code and c.wh_code= e.wh_code " & vbCrLf & _
            " and b.do_no in (" & Dono & ") order by Delivery_Date,a.Item_Code,Lot_No,a.PO_NO,Seq_No,DOSeq_No"
        
    Set rs1 = New Recordset

    rs1.CursorLocation = adUseClient
    rs1.Open sql, Db, adOpenKeyset, adLockOptimistic
    If Not rs1.EOF Then
        
        Set report = application.OpenReport(App.path & "\Reports\Delivery_Instruction.rpt")
        report.Database.Tables(1).SetDataSource rs1
        
        '#####################################################################
        '# Qty Digit and decimal
        report.FormulaFields.GetItemByName("DecimalQty").Text = "" & gi_decimalDigitQty & ""
        report.FormulaFields.GetItemByName("DecimalBox").Text = "" & gi_decimalDigitBox & ""
        '#####################################################################
        
        Dim Rpt As New FrmRpt3
        pesaninvalid = ""
        reportcode = "DI"
        printorient = 1
        sqlprint = sql
        do_no = Dono
        With Rpt.CRViewer1
            .ReportSource = report
            .ViewReport
            .Zoom 1
        End With
        With Rpt
            .WindowState = 2
            .Show 1
        End With
    End If
End Function

Function PackingPrintStatus(PackingNo$) As Boolean
    
    Dim rstdo As Recordset
    sql = "select * from Packing_Master where Packing_No in (" & PackingNo & ") and fix_cls = '1'"
    Set rstdo = New Recordset
    rstdo.Open sql, Db, adOpenDynamic, adLockOptimistic
    If rstdo.EOF Then
        pesaninvalid = "Packing No " & PackingNo$ & " is not fixed!"
    Else
        pesaninvalid = ""
    End If

End Function

Function PackingPrint(PackingNo$) As String
    
    Dim application As New CRAXDDRT.application
    Dim report As New CRAXDDRT.report
    Dim rsMain As New Recordset
    Dim rsSub As New Recordset
    Dim RsSubDetail As New Recordset
   
    sql = " Select PM.Packing_No,PM.Packing_Date,PM.Stuffing_Date,PM.ETD,PM.ETA,PM.Payment_Days,PM.Payment, " & vbCrLf & _
                      "     PM.Transportation_Cls,PM.Vessel,PM.Mother_Vessel,PM.From_Port,PM.To_Port,PM.Final_Destination,PM.PackingStyle_Cls, " & vbCrLf & _
                      "     PM.Remarks,PM.Last_User, PM.POCaseMark1,PM.POCaseMark2," & vbCrLf & _
                      "     PM.Cust_Code,TM.Trade_Name Customer_Name,TM.Address1 CsAddress1,TM.Address2 CsAddress2,TM.City CsCity, TM.Country CsCountry, " & vbCrLf & _
                      "     PM.Consignee,PM.ConsigneeTitle,TM1.Trade_Name Consignee_Name,TM1.Address1 CgAddress1,TM1.Address2 CgAddress2,TM1.City CgCity, TM1.Country CgCountry, " & vbCrLf & _
                      "     PM.Payment_Terms,Isnull(PT.Description,'') Payment_Desription, " & vbCrLf & _
                      "     PD.Order_No,PD.Container_No,PD.Container_Size,PD.SerialNoFrom,PD.SerialNoTo,PD.Qty,PD.Unit_Cls, " & vbCrLf & _
                      "     PD.QtyWeight_Netto,Pd.QtyWeight_Gross,Rtrim(PD.Ctn_No) Ctn_No,case when PD.Qty_Ctn > 0 then pd.Qty_Ctn else 1 end Qty_Ctn,PD.Length,PD.Width,PD.Thickness, " & vbCrLf & _
                      "     PD.Item_Code,IM.Item_Name " & vbCrLf & _
                      "     From Packing_Master PM " & vbCrLf & _
                      "     Inner Join Packing_Detail PD On PM.Packing_No=PD.Packing_No "
    
    sql = sql + "     Inner Join Trade_Master TM on PM.Cust_Code=TM.Trade_Code " & vbCrLf & _
                      "     Inner Join Trade_Master TM1 on PM.Consignee=TM1.Trade_Code " & vbCrLf & _
                      "     Inner Join Item_Master IM on PD.Item_Code=IM.Item_Code " & vbCrLf & _
                      "     Left Join PaymentTerm_Cls PT on PM.Payment_Terms=PT.PaymentTerm_Cls " & vbCrLf & _
                      "     Where PM.Packing_No=" & Trim(PackingNo$) & " " & vbCrLf & _
                      "     Order By PM.Packing_No, PD.PackingSeq_No, PD.Item_Code"


    sqlprint = sql
    sqlprint2 = sql
    rsMain.CursorLocation = adUseClient
    If rsMain.State <> adStateClosed Then rsMain.Close
    rsMain.Open sql, Db, adOpenKeyset, adLockOptimistic
    
    If Not rsMain.EOF Then
            
        Set report = application.OpenReport(App.path & "\Reports\packinglist_KWI.rpt")
        report.Database.Tables(1).SetDataSource rsMain
        report.DiscardSavedData
'        report.FormulaFields.GetItemByName("DecimalQty").Text = gi_decimalDigitQty
'        report.FormulaFields.GetItemByName("DecimalWeight").Text = gi_decimalDigitWeight
'
'        rsSub.CursorLocation = adUseClient
'        If rsSub.State <> adStateClosed Then rsSub.Close
'        rsSub.Open Sql, Db, adOpenKeyset, adLockOptimistic
'
'        Sql = "Select RTRim(pm.Packing_No) Packing_No, pm.Packing_Date, pm.PackingStyle_Cls, " & _
'              vbLf & "RTrim(pc.Description) PackingStyleDesc, " & _
'              vbLf & "RTrim(pd.Container_No) Container_No, RTrim(Ctn_No) Ctn_No, " & _
'              vbLf & "Rtrim(pd.Order_No) OrderNo, pm.List_PO, pm.List_PODate, " & _
'              vbLf & "RTrim(pm.payment_terms) PaymentTerms, " & _
'              vbLf & "RTrim(ptc.Description) PaymentTermsDesc, " & _
'              vbLf & "RTRim(pd.Item_Code) Item_Code, RTrim(pd.MakerItem_Code) MakerItem_Code, Isnull(pd.Qty, 0) Qty, " & _
'              vbLf & "Isnull(pd.Qty_Ctn, 0) Qty_Ctn,Isnull(pd.QtyWeight_Netto, 0) QtyWeight_Netto, " & _
'              vbLf & "Isnull(pd.QtyWeight_Gross, 0) QtyWeight_Gross, RTrim(im.Item_Name) Item_Name, im.Group_Cls, " & _
'              vbLf & "RTrim(uc.Description) Unit_Desc, " & _
'              vbLf & "RTrim(tm.Trade_Name) Cust_Name, RTrim(tm.Address1) Cust_Address1, " & _
'              vbLf & "RTrim(tm.Address2) Cust_Address2, " & _
'              vbLf & "RTrim(tm.City) Cust_City " & _
'              vbLf & "From Packing_Master pm " & _
'              vbLf & "Inner Join Packing_Detail pd On pm.Packing_No = pd.Packing_No " & _
'              vbLf & "Inner Join Item_Master im On pd.Item_Code = im.Item_Code " & _
'              vbLf & "Inner Join Trade_Master tm On pm.Cust_Code = tm.Trade_Code " & _
'              vbLf & "Left Outer Join Unit_Cls uc On pd.Unit_Cls = uc.Unit_Cls " & _
'              vbLf & "Left Join Trade_Master tm2 On pm.consignee = tm2.trade_Code " & _
'              vbLf & "Left Join PackingStyle_cls pc On pc.packingStyle_cls = pm.PackingStyle_Cls " & _
'              vbLf & "Left Join paymentterm_cls ptc on ptc.PaymentTerm_Cls = pm.payment_terms " & _
'              vbLf & vbLf & "Where pm.Packing_No In (" & PackingNo$ & ") " & _
'              vbLf & "Order By pm.Packing_No, pd.Container_No, pd.Item_Code "
'
'        sqlprint3 = Sql
'        rsSubDetail.CursorLocation = adUseClient
'        If rsSubDetail.State <> adStateClosed Then rsSubDetail.Close
'        rsSubDetail.Open Sql, Db, adOpenKeyset, adLockOptimistic
'
'        If Not rsSubDetail.EOF Then
'        End If
'
        Dim Rpt As New FrmRpt3
        'reportcode = "packing_list_ex"
        reportcode = "packinglist_Kwi"
        printorient = 1
        packing_no = PackingNo
        With Rpt.CRViewer1
            .ReportSource = report
            .ViewReport
            .Zoom 1
        End With
        With Rpt
            .WindowState = 2
            .Show 1
        End With
    End If

End Function

Public Sub POLocal(strPONo As String, bteHakPrice As Byte, PoCls As Byte, tcode As String, DelDate As Date)
    
    Dim application As New CRAXDDRT.application
    Dim report As New CRAXDDRT.report
    Dim rs1 As Recordset
    
    If PoCls = 1 Then
                sql = "select '" & bteHakPrice & "' AS HakPrice, rtrim(a.po_no) po_no, a.po_date, a.delivery_date, a.amount as tamount, a.ppn, a.total_amount, rtrim(e.trade_name) trade_name, rtrim(e.address1) taddress1, rtrim(e.address2) taddress2, rtrim(e.city) tcity, rtrim(e.postal_code) tpostal_code, " & _
                         " (SELECT Description FROM Curr_Cls  WHERE (Curr_Cls = b.Currency_Code)) AS Curr_desc, (SELECT     Description FROM Unit_Cls AS uc    WHERE      (Unit_Cls = b.Unit_Cls)) AS Unit_Desc," & _
                        "rtrim(b.item_code) item_code, rtrim(c.item_name) item_name, c.finishgoodpart_cls, c.number_entering, c.number_box, b.unit_cls, b.qty, b.currency_code, b.price, b.amount, " & _
                        "rtrim(f.company_name) company_name, rtrim(f.address1) caddress1, rtrim(f.address2) caddress2, rtrim(f.Province) cprovince, rtrim(f.City) ccity, rtrim(f.postal_code) cpostal_code, rtrim(f.phone1) cphone1, rtrim(f.phone2) cphone2, rtrim(f.fax) cfax, " & _
                        "rtrim(PO_position) po_position, rtrim(PO_person) po_person, (select description from paymentterm_cls where paymentterm_cls = a.paymentterm_cls) payment_terms, e.country_cls, e.trade_cls,(select description from priceCondition_cls where priceCondition_cls = a.priceCondition_cls) as Price_cls,(select description from packingstyle_cls where packingstyle_cls = a.POPacking_cls) as Packing_cls,(select description from Insurance_cls where insurance_cls = a.insurance_cls) as insurance_cls,(select description from Transportation_cls where Transportation_cls = a.Transportation_cls) as Transportation_cls, a.remarks " & _
                        " from purchaseorder_master a, purchaseorder_detail b, item_master c, trade_master e, company_profile f " & _
                        "where a.po_no=b.po_no and b.item_code=c.item_code and A.supplier_code = E.trade_code and a.po_no='" & Trim(strPONo) & "' " & _
                        "Group By e.trade_code,e.trade_name ,e.address1 ,e.address2,e.city,e.postal_code, a.paymentterm_cls, e.country_cls, e.trade_cls, a.po_no,a.po_date,a.delivery_date, a.amount, a.ppn, a.total_amount, b.item_code,c.item_name, " & _
                        "C.finishgoodpart_cls , C.number_entering, C.number_box, B.qty, B.unit_cls, B.currency_code, B.price, B.amount, F.company_name, F.address1, F.address2, F.Province, F.City, F.Postal_code, F.phone1, F.phone2, F.fax, po_position, po_person, a.pricecondition_cls, a.popacking_cls,a.insurance_cls, a.Transportation_cls, a.remarks Union all " & _
                        "select '" & bteHakPrice & "' AS HakPrice, rtrim(a.po_no) po_no, a.po_date, a.delivery_date, a.amount as tamount, a.ppn, a.total_amount, rtrim(e.trade_name) trade_name, rtrim(e.address1) taddress1, rtrim(e.address2) taddress2, rtrim(e.city) tcity, rtrim(e.postal_code) tpostal_code, (SELECT Description FROM Curr_Cls   WHERE (Curr_Cls = p.Currency_Code)) AS Curr_desc,(SELECT     Description FROM Unit_Cls AS uc    WHERE      (Unit_Cls = p.Unit_Cls)) AS Unit_Desc," & _
                        "rtrim(p.item_code) item_code, rtrim(q.item_name) item_name , q.finishgoodpart_cls, q.number_entering, q.number_box, p.unit_cls, 0 as qty, p.currency_code, p.price, 0 as amount, " & _
                        "rtrim(f.company_name) company_name, rtrim(f.address1) caddress1, rtrim(f.address2) caddress2, rtrim(f.Province) cprovince, rtrim(f.City) ccity, rtrim(f.postal_code) cpostal_code, rtrim(f.phone1) cphone1, rtrim(f.phone2) cphone2, rtrim(f.fax) cfax, " & _
                        "rtrim(PO_position) po_position, rtrim(PO_person) po_person, (select description from paymentterm_cls where paymentterm_cls = a.paymentterm_cls) payment_terms, e.country_cls, e.trade_cls,(select description from priceCondition_cls where priceCondition_cls = a.priceCondition_cls) as Price_cls,(select description from packingstyle_cls where packingstyle_cls = a.POPacking_cls) as Packing_cls,(select description from Insurance_cls where insurance_cls = a.insurance_cls) as insurance_cls,(select description from Transportation_cls where Transportation_cls = a.Transportation_cls) as Transportation_cls, a.remarks " & _
                        " from purchaseorder_master a, trade_master e, company_profile f, price_master p, item_master q " & _
                        "where A.supplier_code = E.trade_code and a.po_no='" & Trim(strPONo) & "' and " & _
                        "p.item_code=q.item_code and (p.trade_code='" & Trim(tcode) & "' or p.trade_code='000000') and start_date<='" & Format(DelDate, "yyyymmdd") & "' and end_date>='" & Format(DelDate, "yyyymmdd") & "' " & _
                        "and price_cls='01' and (rtrim(q.sheetcoil_cls) is null or rtrim(q.sheetcoil_cls)='') and p.item_code not in (select item_code from purchaseorder_detail where po_no='" & Trim(strPONo) & "') Union all " & _
                        "select '" & bteHakPrice & "' AS HakPrice, rtrim(a.po_no) po_no, a.po_date, a.delivery_date, a.amount as tamount, a.ppn, a.total_amount, rtrim(e.trade_name) trade_name, rtrim(e.address1) taddress1, rtrim(e.address2) taddress2, rtrim(e.city) tcity, rtrim(e.postal_code) tpostal_code,'' AS Curr_desc,'' AS Unit_Desc, " & _
                        "rtrim(item_code) item_code, rtrim(item_name) item_name, finishgoodpart_cls, number_entering, number_box, unit_cls, 0 as qty, (select top 1 currency_code from purchaseorder_detail where po_no='" & Trim(strPONo) & "') as currency_code, 0 as price, 0 as amount, " & _
                        "rtrim(f.company_name) company_name, rtrim(f.address1) caddress1, rtrim(f.address2) caddress2, rtrim(f.Province) cprovince, rtrim(f.City) ccity, rtrim(f.postal_code) cpostal_code, rtrim(f.phone1) cphone1, rtrim(f.phone2) cphone2, rtrim(f.fax) cfax, " & _
                        "rtrim(PO_position) po_position, rtrim(PO_person) po_person, (select description from paymentterm_cls where paymentterm_cls = a.paymentterm_cls) payment_terms, e.country_cls, e.trade_cls,(select description from priceCondition_cls where priceCondition_cls = a.priceCondition_cls) as Price_cls,(select description from packingstyle_cls where packingstyle_cls = a.POPacking_cls) as Packing_cls,(select description from Insurance_cls where insurance_cls = a.insurance_cls) as insurance_cls,(select description from Transportation_cls where Transportation_cls = a.Transportation_cls) as Transportation_cls, a.remarks " & _
                        " from purchaseorder_master a, item_master c, trade_master e, company_profile f " & _
                        "where A.supplier_code = E.trade_code and a.po_no='" & Trim(strPONo) & "' and " & _
                        "c.supplier_code='" & Trim(tcode) & "' and item_code not in (select item_code from price_master where (trade_code='" & Trim(tcode) & "' or trade_code='000000') and start_date<='" & Format(DelDate, "yyyymmdd") & "' and end_date>='" & Format(DelDate, "yyyymmdd") & "' and price_cls='01') " & _
                        "and (rtrim(c.sheetcoil_cls) is null or rtrim(c.sheetcoil_cls)='') and item_code not in(select item_code from purchaseorder_detail where po_no='" & Trim(strPONo) & "') "
    Else
                sql = "select '" & bteHakPrice & "' AS HakPrice, rtrim(a.po_no) po_no, a.po_date, a.delivery_date, a.amount as tamount, a.ppn, a.total_amount, rtrim(e.trade_name) trade_name, rtrim(e.address1) taddress1, rtrim(e.address2) taddress2, rtrim(e.city) tcity, rtrim(e.postal_code) tpostal_code, " & _
                         " (SELECT Description FROM Curr_Cls  WHERE (Curr_Cls = b.Currency_Code)) AS Curr_desc, (SELECT     Description FROM Unit_Cls AS uc    WHERE      (Unit_Cls = b.Unit_Cls)) AS Unit_Desc," & _
                        "rtrim(b.item_code) item_code, rtrim(c.item_name) item_name, c.finishgoodpart_cls, c.number_entering, c.number_box, b.unit_cls, b.qty, b.currency_code, b.price, b.amount, " & _
                        "rtrim(f.company_name) company_name, rtrim(f.address1) caddress1, rtrim(f.address2) caddress2, rtrim(f.Province) cprovince, rtrim(f.City) ccity, rtrim(f.postal_code) cpostal_code, rtrim(f.phone1) cphone1, rtrim(f.phone2) cphone2, rtrim(f.fax) cfax, " & _
                        "rtrim(PO_position) po_position, rtrim(PO_person) po_person, (select description from paymentterm_cls where paymentterm_cls = a.paymentterm_cls) payment_terms, e.country_cls, e.trade_cls,(select description from priceCondition_cls where priceCondition_cls = a.priceCondition_cls) as Price_cls,(select description from packingstyle_cls where packingstyle_cls = a.POPacking_cls) as Packing_cls,(select description from Insurance_cls where insurance_cls = a.insurance_cls) as insurance_cls,(select description from Transportation_cls where Transportation_cls = a.Transportation_cls) as Transportation_cls , a.remarks" & _
                        " from purchaseorder_master a, purchaseorder_detail b, item_master c, trade_master e, company_profile f " & _
                        "where a.po_no=b.po_no and b.item_code=c.item_code and A.supplier_code = E.trade_code and a.po_no='" & Trim(strPONo) & "' " & _
                        "Group By e.trade_code,e.trade_name ,e.address1 ,e.address2,e.city,e.postal_code, a.paymentterm_cls, e.country_cls, e.trade_cls, a.po_no,a.po_date,a.delivery_date, a.amount, a.ppn, a.total_amount, b.item_code,c.item_name, " & _
                        "C.finishgoodpart_cls , C.number_entering, C.number_box, B.qty, B.unit_cls, B.currency_code, B.price, B.amount, F.company_name, F.address1, F.address2, F.Province, F.City, F.Postal_code, F.phone1, F.phone2, F.fax, po_position, po_person, a.pricecondition_cls, a.popacking_cls,a.insurance_cls, a.Transportation_cls, a.remarks"
    End If
   
    Set rs1 = New Recordset
    rs1.CursorLocation = adUseClient
    rs1.Open sql, Db, adOpenKeyset, adLockOptimistic
    If Not rs1.EOF Then
        
        Set report = application.OpenReport(App.path & "\Reports\rptPOParts_Lokal.rpt")
        report.Database.Tables(1).SetDataSource rs1
        
        report.FormulaFields(1).Text = "" & gi_decimalDigitQty & ""
        report.FormulaFields(2).Text = "" & gi_decimalDigitPrice & ""
        report.FormulaFields(3).Text = "" & gi_decimalDigitAmount & ""
        report.FormulaFields(6).Text = "" & gi_decimalDigitPriceIDR & ""
        report.FormulaFields(7).Text = "" & gi_decimalDigitAmountIDR & ""
        
        Dim Rpt As New FrmRpt3
        sqlprint = sql
        reportcode = "polocal"
        printorient = 1
        With Rpt.CRViewer1
            .ReportSource = report
            .ViewReport
            .Zoom 1
        End With
        With Rpt
            .WindowState = 2
            .Show 1
        End With
    End If
    
End Sub

Public Sub POImport(strPONo As String, bteHakPrice As Byte, PoCls As Byte, tcode As String, DelDate As Date)
    
    Dim application As New CRAXDDRT.application
    Dim report As New CRAXDDRT.report
    Dim rs1 As Recordset
    Dim SqlRpt As String


SqlRpt = " Select  PD.Po_No,PM.Revise_No,PM.PO_Date,PM.PO_LOT, PM.Total_amount, TM.Trade_Name, RTrim(Tm.Address1) Address1, RTrim(isnull(Tm.Address2,'.')) Address2, RTrim(TM.City) City, RTrim(TM.Country) Country, isnull(TM.Telephone,'')Telephone, isnull(TM.Fax,'')fax, " & vbCrLf & _
                  "     PT.Description TermOfPayment, PC.Description DeliveryTerm, TM.Epte_Cls, " & vbCrLf & _
                  "     DT.Trade_Name DeliverTo,RTrim(DT.Address1) DTAddress1, RTrim(DT.Address2) DTAddress2,RTrim(DT.City) DTCity,RTrim(DT.Country) DTCountry, " & vbCrLf & _
                  "     PM.Last_User,PD.Item_Code,IM.Item_Name,PD.Qty, UC.Description Unit, PD.Price, CC.Description Currency, PD.Amount,PD.Delivery_Date " & vbCrLf & _
                  "  , rtrim(CP.PPC_Person)PPC_Person,rtrim(CP.PPC_Position)PPC_Position," & _
                  " rtrim(cp.Tax_Position)PL_Postion,rtrim(cp.tax_person)PL_person"
    SqlRpt = SqlRpt & ",rtrim(Up.Name) Name " & _
                  "     From PurchaseOrder_Detail PD " & vbCrLf & _
                  "         Inner Join PurchaseOrder_Master PM on PD.PO_No=PM.Po_No " & vbCrLf & _
                  "         Inner Join Trade_Master TM on PM.Supplier_Code=TM.Trade_Code " & vbCrLf & _
                  "         Inner Join WareHouse_Master WM on PM.WHTO=WM.Wh_Code " & vbCrLf & _
                  "         Inner Join Trade_Master DT on WM.Adm_Group=DT.Trade_Code " & vbCrLf & _
                  "         Inner Join Item_Master AS IM ON IM.Item_Code = PD.Item_Code " & vbCrLf & _
                  "         Inner Join Curr_Cls AS CC ON CC.Curr_Cls = PD.Currency_Code  "

SqlRpt = SqlRpt + "         Inner Join Unit_Cls AS UC ON UC.Unit_Cls=PD.Unit_Cls " & vbCrLf & _
                  "         Left Join PaymentTerm_Cls PT on PT.PaymentTerm_Cls = PM.PaymentTerm_Cls " & vbCrLf & _
                  "         Left Join PriceCondition_Cls PC on Pc.PriceCondition_Cls = PM.PriceCondition_Cls,Company_Profile as cp,User_Setup UP" & vbCrLf & _
                  " Where PD.Po_No='" & Trim(strPONo) & "' and userName='" & userLogin & "' " & vbCrLf & _
                  "     Order By PM.PO_Date, PD.PO_No, PD.Item_Code   "
' --------

    Set rs1 = New Recordset
    rs1.CursorLocation = adUseClient
    rs1.Open SqlRpt, Db, adOpenKeyset, adLockOptimistic
    If Not rs1.EOF Then
        
        Set report = application.OpenReport(App.path & "\Reports\reportPo.rpt")
        report.Database.Tables(1).SetDataSource rs1
        report.DiscardSavedData
'        report.FormulaFields(1).Text = "" & gi_decimalDigitQty & ""
'        report.FormulaFields(2).Text = "" & gi_decimalDigitPrice & ""
'        report.FormulaFields(3).Text = "" & gi_decimalDigitAmount & ""
'        report.FormulaFields(6).Text = "" & gi_decimalDigitPriceIDR & ""
'        report.FormulaFields(7).Text = "" & gi_decimalDigitAmountIDR & ""
        Dim rsSubPO As New ADODB.Recordset
        If rsSubPO.State <> adStateClosed Then rsSubPO.Close
        rsSubPO.Open SqlRpt, Db, adOpenKeyset, adLockOptimistic
        If Not rsSubPO.EOF Then
         With report.OpenSubreport("SubReportPO.rpt")
          .Database.Tables(1).SetDataSource rsSubPO
          '.FormulaFields.GetItemByName("SubDetailDecimalQty").Text = gi_decimalDigitQty
          '.FormulaFields.GetItemByName("SubDetailDecimalPrice").Text = gi_decimalDigitPrice
          '.FormulaFields.GetItemByName("SubDetailDecimalPriceIDR").Text = gi_decimalDigitPriceIDR
          '.FormulaFields.GetItemByName("SubDetailDecimalAmount").Text = gi_decimalDigitAmountIDR
         End With
        End If
        
        Dim Rpt As New FrmRpt3
        sqlprint = SqlRpt
        reportcode = "poimport"
        printorient = 1
        With Rpt.CRViewer1
            .ReportSource = report
            .ViewReport
            .Zoom 1
        End With
        With Rpt
            .WindowState = 2
            .Show 1
        End With
    End If
    
End Sub

Public Sub POSubcon(strPONo As String, bteHakPrice As Byte)
    
    Dim application As New CRAXDDRT.application
    Dim report As New CRAXDDRT.report
    Dim rs1 As Recordset
    Dim SqlRpt As String

    
SqlRpt = " Select  PD.Po_No,PM.Revise_No,PM.PO_Date,PM.PO_LOT,PM.Total_amount, TM.Trade_Name, RTrim(Tm.Address1) Address1, RTrim(isnull(Tm.Address2,'.')) Address2, RTrim(TM.City) City, RTrim(TM.Country) Country, isnull(TM.Telephone,'')Telephone, isnull(TM.Fax,'')fax, " & vbCrLf & _
                  "     PT.Description TermOfPayment, PC.Description DeliveryTerm, TM.Epte_Cls, " & vbCrLf & _
                  "     DT.Trade_Name DeliverTo,RTrim(DT.Address1) DTAddress1, RTrim(DT.Address2) DTAddress2,RTrim(DT.City) DTCity,RTrim(DT.Country) DTCountry, " & vbCrLf & _
                  "     PM.Last_User,PD.Item_Code,IM.Item_Name,PD.Qty, UC.Description Unit, PD.Price, CC.Description Currency, PD.Amount,PD.Delivery_Date " & vbCrLf & _
                  "  , rtrim(CP.PPC_Person)PPC_Person,rtrim(CP.PPC_Position)PPC_Position," & _
                  " rtrim(cp.Tax_Position)PL_Postion,rtrim(cp.tax_person)PL_person"
                  SqlRpt = SqlRpt & ",rtrim(Up.Name) Name " & _
                  "     From PurchaseOrder_Detail PD " & vbCrLf & _
                  "         Inner Join PurchaseOrder_Master PM on PD.PO_No=PM.Po_No " & vbCrLf & _
                  "         Inner Join Trade_Master TM on PM.Supplier_Code=TM.Trade_Code " & vbCrLf & _
                  "         Inner Join WareHouse_Master WM on PM.WHTO=WM.Wh_Code " & vbCrLf & _
                  "         Inner Join Trade_Master DT on WM.Adm_Group=DT.Trade_Code " & vbCrLf & _
                  "         Inner Join Item_Master AS IM ON IM.Item_Code = PD.Item_Code " & vbCrLf & _
                  "         Inner Join Curr_Cls AS CC ON CC.Curr_Cls = PD.Currency_Code "

SqlRpt = SqlRpt + "         Inner Join Unit_Cls AS UC ON UC.Unit_Cls=PD.Unit_Cls " & vbCrLf & _
                  "         Left Join PaymentTerm_Cls PT on PT.PaymentTerm_Cls = PM.PaymentTerm_Cls " & vbCrLf & _
                  "         Left Join PriceCondition_Cls PC on Pc.PriceCondition_Cls = PM.PriceCondition_Cls,Company_Profile as cp,User_Setup UP " & vbCrLf & _
                  "     Where PD.Po_No='" & Trim(strPONo) & "' and userName='" & userLogin & "' " & vbCrLf & _
                  "     Order By PM.PO_Date, PD.PO_No, PD.Item_Code   "

    Set rs1 = New Recordset
    rs1.CursorLocation = adUseClient
    rs1.Open SqlRpt, Db, adOpenKeyset, adLockOptimistic
    If Not rs1.EOF Then
        
        Set report = application.OpenReport(App.path & "\Reports\ReportPO.rpt")
        report.Database.Tables(1).SetDataSource rs1
        report.DiscardSavedData
'        report.FormulaFields(1).Text = "" & gi_decimalDigitQty & ""
'        report.FormulaFields(2).Text = "" & gi_decimalDigitPrice & ""
'        report.FormulaFields(3).Text = "" & gi_decimalDigitAmount & ""
'        report.FormulaFields(11).Text = "" & gi_decimalDigitPriceIDR & ""
'        report.FormulaFields(12).Text = "" & gi_decimalDigitAmountIDR & ""
        Dim rsSubPO As New ADODB.Recordset
        If rsSubPO.State <> adStateClosed Then rsSubPO.Close
        rsSubPO.Open SqlRpt, Db, adOpenKeyset, adLockOptimistic
        If Not rsSubPO.EOF Then
         With report.OpenSubreport("SubReportPO.rpt")
          .Database.Tables(1).SetDataSource rsSubPO
          '.FormulaFields.GetItemByName("SubDetailDecimalQty").Text = gi_decimalDigitQty
          '.FormulaFields.GetItemByName("SubDetailDecimalPrice").Text = gi_decimalDigitPrice
          '.FormulaFields.GetItemByName("SubDetailDecimalPriceIDR").Text = gi_decimalDigitPriceIDR
          '.FormulaFields.GetItemByName("SubDetailDecimalAmount").Text = gi_decimalDigitAmountIDR
         End With
        End If
        Dim Rpt As New FrmRpt3
        sqlprint = SqlRpt
        reportcode = "posubcon"
        printorient = 1
        With Rpt.CRViewer1
            .ReportSource = report
            .ViewReport
            .Zoom 1
        End With
        With Rpt
            .WindowState = 2
            .Show 1
        End With
    End If
    
End Sub

Public Function GetLastMonthStock() As String 'YYYYMM
Dim sql As String, RS As New Recordset

sql = "Select * from Inventory_Control Order By Inventory_Year desc,Inventory_Month desc"
RS.Open sql, Db
If Not RS.EOF Then
    GetLastMonthStock = Format(RS!Inventory_Year, "0000") & Format(RS!Inventory_Month, "00")
Else
    GetLastMonthStock = ""
End If
End Function

Function reportrequestauto(ByVal requestno As String) As String
    Dim application As New CRAXDDRT.application
    Dim report As New CRAXDDRT.report
    Dim rsRpt As New ADODB.Recordset
    Dim SqlRpt As String
    Dim Rpt As New FrmRpt3
    
'    SqlRpt = "select rtrim(pd.supplyrec_no) supplyrec_no, isnull(rtrim(pd.parentitem_code),'') parentitem_code, " & vbCrLf & _
'        "rtrim(pim.makeritem_code) pmakercode, rtrim(pim.item_name) item_name, pd.dailyseq_no, dp.qty, pd.childlot_no, rtrim(pd.childitem_code) childitem_code, " & vbCrLf & _
'        "rtrim(cim.item_name) childitem_name, pd.childrequirement_qty, isnull(cim.number_box,0) as cnumber_entering, " & vbCrLf & _
'        "pm.childsupply_date, rtrim(company_name) as company_name, rtrim(pm.fromwarehouse_code) from_code, " & vbCrLf & _
'        "case when (Select rtrim(trade_name) from trade_master where trade_code = pm.fromwarehouse_code) is null " & vbCrLf & _
'        "then (Select rtrim(wh_name) from warehouse_master where wh_code = pm.fromwarehouse_code) " & vbCrLf & _
'        "else (Select rtrim(trade_name) from trade_master where trade_code = pm.fromwarehouse_code) end From_name, " & vbCrLf & _
'        "rtrim(pm.towarehouse_code) t_code, bm.Qty qtybom, " & vbCrLf & _
'        "case when (Select rtrim(trade_name) from trade_master where trade_code = pm.towarehouse_code) is null " & vbCrLf & _
'        "then (Select rtrim(wh_name) from warehouse_master where wh_code = pm.towarehouse_code) " & vbCrLf & _
'        "else (Select rtrim(trade_name) from trade_master where trade_code = pm.towarehouse_code) end to_name, " & vbCrLf & _
'        "rtrim(supply_position) supply_position, rtrim(supply_person) supply_person, rtrim(receipt_position) receipt_position, " & vbCrLf & _
'        "rtrim(receipt_person) receipt_person " & vbCrLf & _
'        "from partsupplyrequest_master pm " & vbCrLf & _
'        "inner join  partsupplyrequest_detail pd on pm.supplyrec_no = pd.supplyrec_no " & vbCrLf & _
'        "inner join item_master cim on pd.childitem_code = cim.item_code " & vbCrLf & _
'        "inner join item_master pim on pd.parentitem_code = pim.item_code " & vbCrLf & _
'        "inner join daily_production dp on dp.seq_no = pd.dailyseq_no " & vbCrLf & _
'        "inner join bom_master bm on pd.parentitem_code = bm.parent_itemcode and pd.childitem_code = bm.item_code, company_profile " & vbCrLf & _
'        "where pm.supplyRec_No in (" & requestno & ")"
        
    SqlRpt = " select rtrim(pd.supplyrec_no) supplyrec_no, isnull(rtrim(pd.parentitem_code),'') parentitem_code,  " & vbCrLf & _
                  " rtrim(pim.makeritem_code) pmakercode, rtrim(pim.item_name) item_name, pd.dailyseq_no, dp.qty, pd.childlot_no,  " & vbCrLf & _
                  " CASE WHEN pd.ReplacementItem_Code IS NULL THEN RTRIM(pd.childitem_code) ELSE RTRIM(pd.ReplacementItem_Code) END childitem_code, " & vbCrLf & _
                  " CASE WHEN pd.ReplacementItem_Code IS NULL THEN RTRIM(cim.item_name)  " & vbCrLf & _
                  "      ELSE (SELECT RTRIM(Item_Name) FROM dbo.Item_Master WHERE Item_Code = pd.ReplacementItem_Code)  " & vbCrLf & _
                  " END childitem_name, " & vbCrLf & _
                  " pd.childrequirement_qty, isnull(cim.number_box,0) as cnumber_entering,  " & vbCrLf & _
                  " pm.childsupply_date, rtrim(company_name) as company_name, rtrim(pm.fromwarehouse_code) from_code,  " & vbCrLf & _
                  " case when (Select rtrim(trade_name) from trade_master where trade_code = pm.fromwarehouse_code) is null  " & vbCrLf & _
                  " then (Select rtrim(wh_name) from warehouse_master where wh_code = pm.fromwarehouse_code)  " & vbCrLf & _
                  " else (Select rtrim(trade_name) from trade_master where trade_code = pm.fromwarehouse_code) end From_name,  "

    SqlRpt = SqlRpt + " rtrim(pm.towarehouse_code) t_code,  " & vbCrLf & _
                      " (Select top 1 Qty From BOM_Master TR where TR.Item_Code=pd.childItem_Code) QTYBOM,  " & vbCrLf & _
                      " case when (Select rtrim(trade_name) from trade_master where trade_code = pm.towarehouse_code) is null  " & vbCrLf & _
                      " then (Select rtrim(wh_name) from warehouse_master where wh_code = pm.towarehouse_code)  " & vbCrLf & _
                      " else (Select rtrim(trade_name) from trade_master where trade_code = pm.towarehouse_code) end to_name,  " & vbCrLf & _
                      " rtrim(supply_position) supply_position, rtrim(supply_person) supply_person, rtrim(receipt_position) receipt_position,  " & vbCrLf & _
                      " rtrim(receipt_person) receipt_person  " & vbCrLf & _
                      " from partsupplyrequest_master pm  " & vbCrLf & _
                      " inner join  partsupplyrequest_detail pd on pm.supplyrec_no = pd.supplyrec_no  " & vbCrLf & _
                      " inner join item_master cim on pd.childitem_code = cim.item_code  " & vbCrLf & _
                      " inner join item_master pim on pd.parentitem_code = pim.item_code  "
    
    SqlRpt = SqlRpt + " inner join daily_production dp on dp.seq_no = pd.dailyseq_no  " & vbCrLf & _
                      " , company_profile  " & vbCrLf & _
                      " where pm.supplyRec_No in (" & requestno & ") And Cim.Material_Cls <> '02' "


    If rsRpt.State <> adStateClosed Then rsRpt.Close
    rsRpt.Open SqlRpt, Db, adOpenDynamic, adLockOptimistic
    
    sqlprint = SqlRpt
    
    If rsRpt.EOF Then Exit Function
    
    Set report = application.OpenReport(App.path & "\Reports\requestauto.rpt")
    reportcode = "rptrequestauto"
    printorient = "2"
    report.Database.Tables(1).SetDataSource rsRpt
    report.FormulaFields(1).Text = "" & gi_decimalDigitQtyBOM & ""
    
    Rpt.CRViewer1.ReportSource = report
    Rpt.CRViewer1.ViewReport
    Rpt.CRViewer1.Zoom 1
    Rpt.WindowState = 2
    Rpt.Show 1
    
End Function

Function CekClsStokDanWarehouse(idstock As String, idwarehouse As String) As Boolean
Dim rsSt As New ADODB.Recordset
Dim rsWr As New ADODB.Recordset

CekClsStokDanWarehouse = False

If rsSt.State = 1 Then rsSt.Close
rsSt.Open "select stockcontrol_cls from item_master where item_code = '" & idstock & "'", Db, adOpenKeyset, adLockOptimistic
If Not rsSt.EOF And Not rsSt.BOF Then
    If rsSt.Fields("stockcontrol_cls") = "01" Then
    
        If rsWr.State = 1 Then rsWr.Close
        rsWr.Open "select stockcontrol_cls from warehouse_master where wh_code = '" & idwarehouse & "'", Db, adOpenKeyset, adLockOptimistic
        If Not rsWr.EOF And Not rsWr.BOF Then
            If rsWr.Fields("stockcontrol_cls") = "01" Then
                CekClsStokDanWarehouse = True
            Else
                CekClsStokDanWarehouse = False
            End If
        End If
        
    Else
        CekClsStokDanWarehouse = False
    End If
End If
End Function

Public Sub BC40(strBCNo As String)

    Dim application As New CRAXDDRT.application
    Dim report As New CRAXDDRT.report
    Dim rs1 As Recordset
    Dim rs2 As Recordset
    Dim rs3 As Recordset
    
    sql = "Select RTrim(pr.SuratJalan_No) SuratJalan_No, pr.Receipt_Date, RTrim(pr.Item_Code) Item_Code, pr.Qty, " & _
        "RTrim(pr.Remarks) Remarks, pr.Amount, pr.bc40_no, RTrim(pr.PO_No) PO_No, pm.PO_Date, RTrim(im.Item_Name) Item_Name, " & _
        "RTrim(im.MakerItem_Code) MakerItem_Code, RTrim(uc.Description) Unit_Desc, RTrim(cc.Description) Curr_Desc, " & _
        "pr.Package_Qty, RTRim(pr.Package_Cls) Package_Cls, RTrim(pc.Description) Package_Desc, RTrim(tc.Description) Transport_Desc, " & _
        "RTRim(tm.NPWP_No) CustNPWP_No, RTrim(tm.NPWP_Name) CustNPWP_Name, RTrim(tm.NPWP_Address) CustNPWP_Address, " & _
        "RTrim(tm.NPWP_City) CustNPWP_City, RTrim(tm.NPPKP_No) CustNPPKP_No, RTrim(tm.Trade_Code)Cust_Code, " & _
        "RTrim(tm.Trade_Name)Cust_Name, RTrim(tm.Address1)Cust_Address1, RTrim(tm.Address2)Cust_Address2, " & _
        "RTrim(tm.City)Cust_City, RTRim(cp.NPWP_No) NPWP_No, RTrim(cp.NPWP_Name) NPWP_Name, RTrim(cp.NPWP_Address) NPWP_Address, " & _
        "RTrim(cp.NPWP_City) NPWP_City, RTrim(cp.Company_Name) Comp_Name, RTrim(cp.Address1) Comp_Address1, " & _
        "RTrim(cp.Address2) Comp_Address2, RTrim(cp.City) Comp_City, RTrim(fc.Contact_Person) Contact_Person, " & _
        "RTrim(cp.BC_Person1) BC_Person1, RTrim(cp.BC_NIP1) BC_NIP1, RTrim(cp.BC_Person2) BC_Person2, RTrim(cp.BC_NIP2) BC_NIP2, " & _
        "Rate = Isnull((Select Tax_ExchangeRate From Tax_exchangerate where currency_code = pr.Currency_Code And Start_Date <= pr.Receipt_Date And End_Date >= pr.Receipt_Date), 0) " & _
        "From Part_Receipt pr " & _
        "Left Outer Join PurchaseOrder_Master pm On pr.Supplier_Code = pm.Supplier_Code And pr.PO_No = pm.PO_No " & _
        "Inner Join Trade_Master tm On pr.Supplier_Code = tm.Trade_Code " & _
        "Inner Join Item_Master im On pr.Item_Code = im.Item_Code " & _
        "Left Outer Join Unit_Cls uc On pr.Unit_Cls = uc.Unit_Cls " & _
        "Left Outer Join Curr_Cls cc On pr.Currency_Code = cc.Curr_Cls " & _
        "Left Outer Join Warehouse_Master wh on pr.Warehouse_Code = wh.WH_Code " & _
        "Left Outer Join Package_Cls pc On pr.Package_Cls = pc.Package_Cls " & _
        "Left Outer Join Transport_Cls tc On pr.Transport_Cls = tc.Transport_Cls " & _
        "Left Outer Join Trade_Master fc on wh.Adm_Group = fc.Trade_Code, Company_Profile cp " & _
        "Where pr.BC40_No = '" & strBCNo & "' " & _
        "And Year(pr.Receipt_Date) = '" & FrmPart_Rec.TglReceipt.Year & "' " & _
        "Order By pr.MakerItem_Code"

    Set rs1 = New Recordset
    rs1.CursorLocation = adUseClient
    rs1.Open sql, Db, adOpenKeyset, adLockOptimistic
    
    If Not rs1.EOF Then
        
        Set report = application.OpenReport(App.path & "\Reports\bc_40.rpt")
        report.Database.Tables(1).SetDataSource rs1
        
        report.FormulaFields(1).Text = "" & gi_decimalDigitQty & ""
        report.FormulaFields(2).Text = "" & gi_decimalDigitAmount & ""
        report.FormulaFields(3).Text = "" & gi_decimalDigitAmountIDR & ""
        report.FormulaFields(6).Text = "" & gi_decimalDigitBox & ""
        report.FormulaFields(7).Text = "" & rs1.RecordCount & ""
        
        sqlprint2 = "select distinct pr.bc40_no, case when pr.po_no = '0' then pr.remarks else pr.po_no end po_no, pm.po_date " & _
            "From Part_Receipt pr " & _
            "Left Outer Join PurchaseOrder_Master pm On pr.Supplier_Code = pm.Supplier_Code And pr.PO_No = pm.PO_No " & _
            "Where pr.bc40_no = '" & strBCNo & "' " & _
            "And Year(pr.Receipt_Date) = '" & FrmPart_Rec.TglReceipt.Year & "' "

        Set rs2 = New Recordset
        rs2.CursorLocation = adUseClient
        rs2.Open sqlprint2, Db, adOpenKeyset, adLockOptimistic
        report.OpenSubreport("po").Database.Tables(1).SetDataSource rs2
        
        sqlprint3 = "Select pr.bc40_no, rtrim(cc.description) curr_desc, sum(amount) amount, " & _
            "Rate = Isnull((Select Tax_ExchangeRate From Tax_exchangerate where currency_code = pr.Currency_Code And Start_Date <= pr.Receipt_Date And End_Date >= pr.Receipt_Date), 0) " & _
            "From Part_Receipt pr " & _
            "Left Outer Join Curr_Cls cc On pr.Currency_Code = cc.Curr_Cls " & _
            "Where pr.bc40_no = '" & strBCNo & "' " & _
            "And Year(pr.Receipt_Date) = '" & FrmPart_Rec.TglReceipt.Year & "' " & _
            "Group By pr.bc40_no, cc.description, pr.Currency_Code, pr.Receipt_Date"
        
        Set rs3 = New Recordset
        rs3.CursorLocation = adUseClient
        rs3.Open sqlprint3, Db, adOpenKeyset, adLockOptimistic
        report.OpenSubreport("amount").Database.Tables(1).SetDataSource rs3
        report.OpenSubreport("amount").FormulaFields(3).Text = "" & gi_decimalDigitAmount & ""
        report.OpenSubreport("amount").FormulaFields(4).Text = "" & gi_decimalDigitAmountIDR & ""
                        
        Dim Rpt As New FrmRpt3
        sqlprint = sql
        reportcode = "bc40"
        printorient = 1
        With Rpt.CRViewer1
            .ReportSource = report
            .ViewReport
            .Zoom 1
        End With
        With Rpt
            .WindowState = 2
            .Show 1
        End With
    End If
    
End Sub

Private Sub LoadDynamicForm(newFormName As String)
    Dim DynamicForm As Form
    Set DynamicForm = Forms.Add(newFormName)
    Load DynamicForm
    DynamicForm.Show
End Sub

Private Sub hideThisForm(strFormName As String)
Dim i As Integer
    For i = 0 To Forms.Count - 1
        If (Forms(i).Name = strFormName) Then
            Forms(i).Hide
            Exit For
        End If
    Next
End Sub

Function Get_Record(sql)
Dim Cnul As New ADODB.Recordset
Set Cnul = Nothing
Cnul.Open sql, Db, 1, 3
If Cnul.EOF Then
Get_Record = ""
Else
Get_Record = Cnul.Fields(0)
End If
End Function

'added by do 26 Aug 2016 , untuk report Loading Form
Function LoadForm(po_no$) As String
Dim application As New CRAXDDRT.application
Dim report As New CRAXDDRT.report
Dim rs1 As Recordset
             
    sql = " SELECT a.Cust_Code, a.item_code, d.item_name, Delivery_date, " & _
          " a.Po_No, Qty, SerialNoFrom, SerialNoTo, Remarks ,b.location_code,trade_name " & _
          " from orderentry_detail a " & _
          " INNER JOIN orderentry_master b ON a.po_no=b.po_no " & _
          " INNER JOIN trade_master c ON b.location_code=c.trade_code " & _
          " INNER JOIN item_master d ON a.item_code = d.item_code " & _
          " where a.po_no= " & po_no & " " & _
          " order by a.delivery_date, a.makeritem_code, a.seq_no "
   
    Set rs1 = New Recordset

    rs1.CursorLocation = adUseClient
    rs1.Open sql, Db, adOpenKeyset, adLockOptimistic
    If Not rs1.EOF Then
        
        Set report = application.OpenReport(App.path & "\Reports\LoadingForm.rpt")
        report.Database.Tables(1).SetDataSource rs1
        
''        '#####################################################################
''        '# Qty Digit and decimal
''        report.FormulaFields.GetItemByName("DecimalQty").Text = "" & gi_decimalDigitQty & ""
''        report.FormulaFields.GetItemByName("DecimalBox").Text = "" & gi_decimalDigitBox & ""
''        '#####################################################################
        
'        Dim Rpt As New FrmRpt3
'        pesaninvalid = ""
'        reportcode = "LoadingForm"
''        printorient = 1
'        sqlprint = sql
'        po_no = po_no
'        With Rpt.CRViewer1
'            .ReportSource = report
'            .ViewReport
'            .Zoom 1
'        End With
'        With Rpt
'            .WindowState = 2
'            .Show 1
'        End With

        Dim Rpt As New FrmRpt3
        sqlprint = sql
        reportcode = "LoadingForm"
        printorient = 1
        With Rpt.CRViewer1
            .ReportSource = report
            .ViewReport
            .Zoom 1
        End With
        With Rpt
            .WindowState = 2
            .Show 1
        End With
              


    End If
 
End Function

