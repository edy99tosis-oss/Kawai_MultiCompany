Attribute VB_Name = "MdlNIC"
Option Explicit

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Db As New ADODB.Connection
Public ErrSt As Boolean

Public IniFile As String
Public ConnStr As String, DSN As String, userName As String, Password
Public userLogin As String, StatusAdmin As Integer

Public Sql As String
'report
Public TutupPtr As Boolean, pesaninvalid As String, reportcode As String
Public do_no As String, inv_no As String
'----
Public Const isiunit = "pcs,g,Kg,t,cm,m,cc,Ltr,Kl"
Public Const isiCurr = "YEN,US$,IDR,EUR,S$"
Public Const bulan = "Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec"

Public kirimPar As String

Public Function DisplayMsg(n As String)
    Dim StrErr
    Dim rs As New ADODB.Recordset
    Dim Sql

    Sql = "select * from message where MsgId=" & n
    Set rs = Db.Execute(Sql)
    If Not (rs.EOF And rs.BOF) Then
        StrErr = "[" & CStr(Trim(rs("MsgId"))) & "]  " & Trim(rs("MsgDesc"))
        DisplayMsg = StrErr
        Set rs = Nothing
        ErrSt = True
    End If
End Function

Public Function hakAkses(mn As String, Optional desc As String) As Integer
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    
    Sql = "select status from user_Privilege where App_ID = 'P01' " & _
        "and Menu_ID = (select Menu_ID from User_Menu where Menu_Name ='" & mn & "'"
        
    If desc <> "" Then
        Sql = Sql & " and Menu_Desc ='" & desc & "'"
    End If
    
    Sql = Sql & ") and userName ='" & userLogin & "'"
    
    Set rs = Db.Execute(Sql)
    If Not (rs.EOF And rs.BOF) Then
        hakAkses = rs("status")
    End If
End Function

Public Function hakUpdate(mn As String, Optional desc As String) As Integer
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    
    Sql = "select Allow_Update from user_Privilege where App_ID = 'P01' " & _
        "and Menu_ID = (select Menu_ID from User_Menu where Menu_Name ='" & mn & "'"
        
    If desc <> "" Then
        Sql = Sql & " and Menu_Desc ='" & desc & "'"
    End If
    
    Sql = Sql & ") and userName ='" & userLogin & "'"
    
    Set rs = Db.Execute(Sql)
    If Not (rs.EOF And rs.BOF) Then
        hakUpdate = rs("Allow_Update")
    End If
End Function

Public Function frmcode(frmName$, Optional Ket As String)
    Dim Sql As String
    Dim rst As Recordset
    Sql = "select * from user_menu where menu_name='" & frmName & "'"
    If Ket <> "" Then
        Sql = Sql & " and menu_desc='" & Ket & "'"
    End If
    Set rst = New Recordset
    rst.Open Sql, Db, adOpenKeyset, adLockOptimistic
    If rst.EOF = False Then frmcode = Trim(rst!menu_id)
End Function

Public Function panggilForm(nmMenu As String, Optional nmForm As String) As Integer
Dim rs As New Recordset
Dim cItem As cExplorerBarItem

    Sql = "select Menu_Name,Menu_Desc from user_menu where menu_ID='" & nmMenu & "'"
    Set rs = Db.Execute(Sql)
    
    If rs.EOF And rs.BOF Then
        panggilForm = 1 'msg Error "Menu ID not Found"
        Exit Function
    Else
        If (LCase(Trim(rs(0))) = LCase(Trim(nmForm))) Then panggilForm = 2: Exit Function
            Dim i As Integer, j As Integer
            'For i = 1 To frmMainMenu.tree.Nodes.Count
            For i = 1 To frmMainMenu.vbalExplorerBarCtl1.Bars.Count
                For j = 1 To frmMainMenu.vbalExplorerBarCtl1.Bars(i).Items.Count
                    If Trim(frmMainMenu.vbalExplorerBarCtl1.Bars(i).Items(j).Key) = "n" & Trim(rs(1)) Then
                        Set cItem = frmMainMenu.vbalExplorerBarCtl1.Bars(i).Items(j)
                        frmMainMenu.vbalExplorerBarCtl1_ItemClick cItem
                        Exit Function
                    End If
                Next j
            Next i
    End If
End Function

'***** Filter Combo Unit ******
Public Sub filterCboUnit(unitCls As String, nmCombo, Optional textKol As Integer)
    If unitCls = "00" Then
        Call isiCboUnitCurr(nmCombo, isiunit, 0, 8, textKol)
    ElseIf unitCls = "01" Then
        Call isiCboUnitCurr(nmCombo, isiunit, 0, 8, textKol)
    ElseIf unitCls = "02" Or unitCls = "03" Or unitCls = "04" Then
        Call isiCboUnitCurr(nmCombo, isiunit, 0, 3, textKol)
    ElseIf unitCls = "05" Or unitCls = "06" Then
        Call isiCboUnitCurr(nmCombo, isiunit, 4, 5, textKol)
    ElseIf unitCls = "07" Or unitCls = "08" Or unitCls = "09" Then
        Call isiCboUnitCurr(nmCombo, isiunit, 6, 8, textKol)
    End If
End Sub

'******** Isi Combo Unit atau Currency ******
Public Sub isiCboUnitCurr(nmCombo, nmField, mulai As Integer, akhir As Integer, Optional textKol As Integer)    'Isi Combo Unit
Dim j As Integer, i As Integer

With nmCombo
    .clear
    .ColumnCount = 2
    
    If textKol = 0 Then .TextColumn = 1 Else .TextColumn = 2
    
    If mulai = 4 Or mulai = 6 Then
        j = 1
        .AddItem ""
        .List(0, 0) = "01"
        .List(0, 1) = "pcs"
    Else
        j = 0
    End If
    
    For i = mulai To akhir
        .AddItem ""
        .List(j, 0) = Format(i + 1, "0#")
        .List(j, 1) = Split(nmField, ",")(i)
        j = j + 1
    Next i
    .ListWidth = 60
    .ColumnWidths = "20 pt;40 pt"
End With
End Sub

Function Doreport(Dono$) As String
Dim application As New CRAXDDRT.application
Dim report As New CRAXDDRT.report
Dim Rs1 As Recordset
   
    Sql = "select d.trade_code,rtrim(d.trade_name) trade_name,rtrim(d.address1) address1,rtrim(d.address2) address2," & _
            "rtrim(d.city) city, rtrim(d.postal_code) postal_code,a.do_no,b.do_date,rtrim(a.po_no) po_no,a.delivery_date, a.item_code, rtrim(c.item_name) item_name,rtrim(a.makeritem_code) makeritem_code, " & _
            "C.number_entering , A.qty, A.unit_cls, rtrim(isnull(b.list_po,'')) list_po " & _
            ",rtrim(f.company_name) company_name, rtrim(f.address1) cpaddress1, " & _
            "rtrim(f.address2) cpaddress2,rtrim(f.Province) cpProvince, rtrim(f.City) cpcity, rtrim(f.postal_code) cpPostal_code, " & _
            "rtrim(f.phone1) cpphone1, rtrim(f.phone2) cpphone2, rtrim(f.fax) cpfax,rtrim(sj_position) sjpos, rtrim(sj_person) sjperson " & _
            "from delivery_order a,do_master b, item_master c,trade_master d ,company_profile f " & _
            "Where B.do_no = A.do_no and b.cust_code = d.trade_code and a.item_code = c.item_code " & _
            " and b.do_no in (" & Dono & ")"
        
    Set Rs1 = New Recordset

    Rs1.CursorLocation = adUseClient
    Rs1.Open Sql, Db, adOpenKeyset, adLockOptimistic
    If Not Rs1.EOF Then
        
        Set report = application.OpenReport(App.Path & "\REPORTs\Delivery_Order.rpt")
        report.Database.Tables(1).SetDataSource Rs1
        
        Dim Rpt As New FrmRpt2
        reportcode = 1
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


Function invreport(invno$) As String
Dim application As New CRAXDDRT.application
Dim report As New CRAXDDRT.report
Dim Rs1 As Recordset, rs2 As Recordset, xdo As String
Dim temp As String, do1 As String, do2 As String, Y As Integer
Dim tbank As String, tadd As String, cbol As Boolean, xy As Integer
   
   
   Sql = "select xxx.*,xxx.invoice_to, rtrim(yyy.trade_name) as Nama1,rtrim(yyy.address1) as add1, rtrim(yyy.address2) as ToAdd2 ,rtrim(yyy.city) city1, rtrim(yyy.postal_code) postal_code1 from ( " & _
            "select d.trade_code,rtrim(d.trade_name) trade_name, rtrim(d.invoice_to) invoice_to,rtrim(d.address1) address1,rtrim(d.address2) address2, " & _
            "rtrim(d.city) city, rtrim(d.postal_code) postal_code,a.invoice_no,e.invoice_date,right(e.delivery_date,2)  bulan, left(e.delivery_date,4) thn, " & _
            " a.item_code,rtrim(isnull(a.makeritem_code,'')) part_no,rtrim(c.item_name) item_name, " & _
            "sum(a.qty) qty ,a.unit_cls,a.currency_code,a.price,sum(a.amount) amount,e.amount invamount,e.ppn,e.total_amount " & _
            ", rtrim(e.list_do) list_do, rtrim(e.list_po) list_po, " & _
            "rtrim(f.company_name) company_name, rtrim(f.address1) cpaddress1, " & _
            "rtrim(f.address2) cpaddress2,rtrim(f.Province) cpProvince, rtrim(f.City) cpcity, rtrim(f.postal_code) cpPostal_code, " & _
            "rtrim(f.phone1) cpphone1, rtrim(f.phone2) cpphone2, rtrim(f.fax) cpfax,rtrim(invoice_position) invpos, rtrim(invoice_person) invperson,d.country_cls " & _
        "from invoice_detail a,do_master b, item_master c,trade_master d, invoice_master e, company_profile f " & _
        "Where A.invoice_no = E.invoice_no and b.do_no= a.do_no and b.cust_code = d.trade_code " & _
            "and a.item_code = c.item_code and a.invoice_no in (" & invno & " ) " & _
        "Group By d.trade_code,d.trade_name ,d.address1 ,d.address2,d.city, d.postal_code," & _
            "a.invoice_no,e.invoice_date,e.delivery_date,a.price, a.item_code, a.makeritem_code,c.item_name, " & _
            "A.unit_cls,a.currency_code, E.amount, E.ppn, E.total_amount,e.list_do,e.list_po " & _
            ",f.company_name, f.address1,f.address2,f.Province, f.City , f.Postal_code, " & _
            "F.phone1 , F.phone2, F.fax,invoice_position,invoice_person,d.invoice_to,d.country_cls ) xxx left join " & _
            "(select * from trade_master ) yyy on xxx.invoice_to = yyy.trade_code " & _
            " order by item_code "
         
    Set Rs1 = New Recordset
    Rs1.CursorLocation = adUseClient
    Rs1.Open Sql, Db, adOpenKeyset, adLockOptimistic
    If Not Rs1.EOF Then
        Set report = application.OpenReport(App.Path & "\REPORTs\invoice.rpt")
        report.Database.Tables(1).SetDataSource Rs1
        Dim Rpt As New FrmRpt2
        reportcode = 2
        Sql = "select *, case currency_code " & _
                "when '01' then 'YEN' " & _
                "when '02' then 'US$' " & _
                "when '03' then 'IDR' " & _
                "when '04' then 'EUR' " & _
                "Else " & _
                    "'S$' " & _
                " end Curr from Company_bank order by bank_name, address1, address2, currency_code"
        Set rs2 = New Recordset
        rs2.Open Sql, Db, adOpenKeyset, adLockOptimistic
        If Not rs2.EOF Then
            For Y = 0 To rs2.RecordCount - 1
                If Y = 0 Then
                    cbol = True
                    tbank = Trim(rs2!bank_name) & Trim(rs2!address1) & Trim(rs2!address2)
                    report.FormulaFields(10).Text = "'" & Trim(rs2!bank_name) & "'"
                    report.FormulaFields(11).Text = "'" & Trim(rs2!address1) & "'"
                    report.FormulaFields(12).Text = "'" & Trim(rs2!address2) & "'"
                    report.FormulaFields(13).Text = "'" & Trim(rs2!city) & "'"
                    report.FormulaFields(14).Text = "'" & Trim(rs2!Postal_code) & "'"
                    report.FormulaFields(15).Text = "'" & Trim(rs2!curr) & " " & Trim(rs2!Account_No) & "'"
                    xy = 1
                Else
                    If tbank = Trim(rs2!bank_name) & Trim(rs2!address1) & Trim(rs2!address2) Then
                        If cbol Then
                            report.FormulaFields(xy + 15).Text = "'" & Trim(rs2!curr) & " " & Trim(rs2!Account_No) & "'"
                        Else
                            report.FormulaFields(xy + 27).Text = "'" & Trim(rs2!curr) & " " & Trim(rs2!Account_No) & "'"
                        End If
                        xy = xy + 1
                    Else
                        If cbol = False Then Exit For
                        cbol = False
                        xy = 1
                        tbank = Trim(rs2!bank_name) & Trim(rs2!address1) & Trim(rs2!address2)
                        report.FormulaFields(22).Text = "'" & Trim(rs2!bank_name) & "'"
                        report.FormulaFields(23).Text = "'" & Trim(rs2!address1) & "'"
                        report.FormulaFields(24).Text = "'" & Trim(rs2!address2) & "'"
                        report.FormulaFields(25).Text = "'" & Trim(rs2!city) & "'"
                        report.FormulaFields(26).Text = "'" & Trim(rs2!Postal_code) & "'"
                        report.FormulaFields(27).Text = "'" & Trim(rs2!curr) & " " & Trim(rs2!Account_No) & "'"
                    End If
                End If
                rs2.MoveNext
            Next
        End If
        
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

Function DOprintStatus(Dono$) As Boolean
Dim rstdo As Recordset

Sql = "select * from DO_master where DO_no in (" & Dono & ") and fix_cls = '1'"
Set rstdo = New Recordset
rstdo.Open Sql, Db, adOpenDynamic, adLockOptimistic
If rstdo.EOF Then
    pesaninvalid = "DO " & Dono & " is not fixed!"
Else
    pesaninvalid = ""
End If
End Function



Function DIreport(Dono$) As String
Dim application As New CRAXDDRT.application
Dim report As New CRAXDDRT.report
Dim Rs1 As Recordset
   
    Sql = "select d.trade_code,rtrim(d.trade_name) trade_name,rtrim(d.address1) address1,rtrim(d.address2) address2, " & _
            "rtrim(d.city) city, rtrim(d.postal_code) postal_code,a.do_no,b.do_date,rtrim(a.po_no) po_no,a.delivery_date, a.item_code,rtrim(c.item_name) item_name, " & _
            "c.number_entering , A.qty, A.unit_cls,e.wh_code, rtrim(b.remarks) note,rtrim(isnull(b.list_po,'')) list_po " & _
            ",rtrim(f.company_name) company_name, rtrim(f.address1) cpaddress1, " & _
            "rtrim(f.address2) cpaddress2,rtrim(f.Province) cpProvince, rtrim(f.City) cpcity, rtrim(f.postal_code) cpPostal_code, " & _
            "rtrim(f.phone1) cpphone1, rtrim(f.phone2) cpphone2, rtrim(f.fax) cpfax, rtrim(DI_position) DIpos, rtrim(DI_person) DIperson " & _
            "from delivery_order a,do_master b, item_master c,trade_master d, warehouse_master e, company_profile f  " & _
            "Where B.do_no = A.do_no and b.cust_code = d.trade_code and a.item_code = c.item_code and c.wh_code= e.wh_code " & _
            " and b.do_no in (" & Dono & ")"
        
    Set Rs1 = New Recordset

    Rs1.CursorLocation = adUseClient
    Rs1.Open Sql, Db, adOpenKeyset, adLockOptimistic
    If Not Rs1.EOF Then
        
        Set report = application.OpenReport(App.Path & "\REPORTs\Delivery_Instruction.rpt")
        report.Database.Tables(1).SetDataSource Rs1
        
        Dim Rpt As New FrmRpt2
        reportcode = 4
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

Public Function GetLastMonthStock() As String 'YYYYMM
Dim Sql As String, rs As New Recordset

Sql = "Select * from Inventory_Control Order By Inventory_Year desc,Inventory_Month desc"
rs.Open Sql, Db
If Not rs.EOF Then
    GetLastMonthStock = Format(rs!inventory_year, "0000") & Format(rs!inventory_month, "00")
Else
    GetLastMonthStock = ""
End If
End Function
