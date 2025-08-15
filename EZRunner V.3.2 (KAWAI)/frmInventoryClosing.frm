VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInventoryClosing 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Inventory Stock Closing"
   ClientHeight    =   4365
   ClientLeft      =   120
   ClientTop       =   345
   ClientWidth     =   8355
   Icon            =   "frmInventoryClosing.frx":0000
   LinkMode        =   1  'Source
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   8355
   StartUpPosition =   2  'CenterScreen
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
      Left            =   435
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3645
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   585
      Left            =   435
      TabIndex        =   8
      Top             =   2940
      Width           =   7335
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
         Height          =   270
         Left            =   105
         TabIndex        =   10
         Top             =   195
         Width           =   7110
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Process"
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
      Left            =   6630
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3645
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   885
      Left            =   435
      TabIndex        =   6
      Top             =   1350
      Width           =   7335
      Begin MSComCtl2.DTPicker period 
         Height          =   345
         Left            =   3765
         TabIndex        =   0
         Top             =   315
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   609
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
         Format          =   92274691
         UpDown          =   -1  'True
         CurrentDate     =   37831
      End
      Begin VB.Label LblMonth 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Month Period :"
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
         Left            =   2385
         TabIndex        =   7
         Top             =   390
         Width           =   1245
      End
   End
   Begin VB.PictureBox PicProg 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   435
      ScaleHeight     =   495
      ScaleWidth      =   7305
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2370
      Width           =   7335
      Begin MSComctlLib.ProgressBar Prg1 
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   90
         Width           =   7065
         _ExtentX        =   12462
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   5925
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   330
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   714
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Inventory Stock Closing"
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
      Height          =   390
      Left            =   435
      TabIndex        =   5
      Top             =   615
      Width           =   7335
   End
End
Attribute VB_Name = "frmInventoryClosing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sql As String, sqlstok As String
Dim RS As New ADODB.Recordset
Dim rsstok As New ADODB.Recordset
Dim i As Double, temptgl As Byte

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    LblErrMsg.Caption = ErrMsg
End If
End Sub

Private Sub Command1_Click()

    Dim temp1 As String, temp2 As String
    Dim s(20) As Double, j As Double, first As Boolean
    Dim kodeitem As String, sqlitem As String, sqlDelete As String
    Dim RsItem As New Recordset
    Dim db1 As New Connection
    Dim sqlhist As String
    Dim lmp, lmi, lmr, lms, lml, lmc
    Dim m As Double
    Dim tanya, a As Double
    Dim reason As String
    Dim cmd As ADODB.Command
    Dim rsSp As ADODB.Recordset

    kodeitem = ""
    j = 1
    m = 0
    first = False
    
    LblErrMsg.Caption = "Start"
    
    If Not (RS.BOF And RS.EOF) Then
        RS.MoveLast
        temp1 = RS("inventory_year") & "-" & RS("inventory_month") & "-01"
        temp2 = Format(period, "yyyy-mm-dd")
        If CDate(temp2) < CDate(temp1) Then
            LblErrMsg.Caption = DisplayMsg("0061")
            period.SetFocus
            Exit Sub
        ElseIf CDate(temp2) = CDate(temp1) Then
            LblErrMsg.Caption = DisplayMsg("0061")
            period.SetFocus
            Exit Sub
        Else
            If DateDiff("m", temp1, temp2) > 1 Then
                LblErrMsg.Caption = DisplayMsg(4058) 'Month Period is not valid !
                period.SetFocus
                Exit Sub
            End If
        End If
    End If
  
    If IsEmpty(tanya) Then tanya = MsgBox("Do You Really Want To Close This Inventory?", vbQuestion & vbYesNo, "Confirmation")
    If tanya = vbYes Then
        
        DoEvents
        
        db1.ConnectionString = Db.ConnectionString
        db1.ConnectionTimeout = 0
        db1.CommandTimeout = 0
        db1.CursorLocation = adUseClient
        db1.Open
        db1.BeginTrans
        
        
        
        sqlstok = "select * from stock_master"
        LblErrMsg.Caption = sqlstok
        If rsstok.State <> adStateClosed Then rsstok.Close
        rsstok.CursorLocation = adUseClient
        rsstok.Open sqlstok, db1, adOpenDynamic, adLockOptimistic
                
        Prg1.Value = 0
                
        If Not (rsstok.BOF And rsstok.EOF) Then
          m = rsstok.RecordCount
            End If
        
        If m = 0 Then
            sqlitem = "select item_code, item_master.wh_code, item_master.stockcontrol_cls, " & _
                      "warehouse_master.stockcontrol_cls from item_master inner join " & _
                      "warehouse_master on item_master.wh_code=warehouse_master.wh_code " & _
                      "where item_master.stockcontrol_cls='01' and warehouse_master.stockcontrol_cls='01' "
                      
            LblErrMsg.Caption = sqlitem
            
            If RsItem.State <> adStateClosed Then RsItem.Close
            RsItem.Open sqlitem, db1, adOpenKeyset, adLockOptimistic
            If Not (RsItem.BOF And RsItem.EOF) Then _
                m = RsItem.RecordCount: first = True
        End If
        
        If m = 0 Then
            Prg1.Max = 1
        Else
            Prg1.Max = m
            
            LblErrMsg.Caption = m
        End If
        
        j = 1

        If first Then
            RsItem.filter = ""
            
            LblErrMsg.Caption = RsItem.Source
            
            RsItem.Requery
            
            If Not (RsItem.BOF And RsItem.EOF) Then
                Do While Not RsItem.EOF
                
                    LblErrMsg.Caption = rsstok.Source
                    
                    rsstok.AddNew
                    rsstok("Warehouse_code") = RsItem("wh_code")
                    rsstok("item_code") = RsItem("item_code")
                    For i = 2 To 20
                        If i <> 8 Or i <> 15 Then
                            rsstok(i) = 0
                        End If
                    Next i
                    rsstok!lm_inventory = 0
                    rsstok!tm_inventory = 0
                    rsstok!nm_inventory = 0
                    
                    rsstok!Last_Update = Now
                    rsstok!last_user = userLogin
                    rsstok.update
    
                    Prg1.Value = j
                    LblErrMsg.Caption = m & j
                    j = j + 1
                     
                    RsItem.MoveNext
                Loop
            End If
        
        Else

            rsstok.filter = ""
            rsstok.Requery
            
            a = DateAdd("m", -1, CDate(period.Value))
            
            If Not (rsstok.BOF And rsstok.EOF) Then
    
              Do While Not rsstok.EOF
                
                lmp = IIf(IsNull(rsstok("lm_premonth")), 0, rsstok("lm_premonth"))
                lmr = IIf(IsNull(rsstok("lm_receipt")), 0, rsstok("lm_receipt"))
                lms = IIf(IsNull(rsstok("lm_supply")), 0, rsstok("lm_supply"))
                lml = IIf(IsNull(rsstok("lm_lossreject")), 0, rsstok("lm_lossreject"))
                lmc = IIf(IsNull(rsstok("lm_current")), 0, rsstok("lm_current"))
                lmi = IIf(IsNull(rsstok("lm_inventory")), 0, rsstok("lm_inventory"))
                reason = IIf(IsNull(rsstok("lm_reason")), 0, rsstok("lm_reason"))
                
                
                LblErrMsg.Caption = "stock_history"
                
                sqlhist = "insert into stock_history (stock_year, stock_month, warehouse_code, item_code, " & _
                          "premonth, Receipt, Supply, LossReject, [Current], inventory, reason) values ('" & Year(a) & "','" & Month(a) & _
                          "','" & Trim(rsstok("Warehouse_code")) & "','" & Trim(rsstok("item_code")) & _
                          "'," & lmp & "," & lmr & "," & lms & "," & lml & "," & lmc & "," & lmi & ",'" & reason & "' ) "
                db1.Execute sqlhist
                
                For i = 9 To 20
                    If i = 15 Then
                        s(i) = 0
                    ElseIf i = 14 Then
                        s(i) = IIf(IsNull(rsstok(i)), rsstok(i - 1), rsstok(i))
                    Else
                        s(i) = IIf(IsNull(rsstok(i)), 0, rsstok(i))
                    End If
                Next i
                             
                For i = 2 To 7
                    rsstok(i) = s(i + 7)
                Next i
                             
                rsstok(9) = s(14)
                rsstok(10) = s(17)
                rsstok(11) = s(18)
                rsstok(12) = s(19)
                rsstok(13) = s(14) + s(17) - s(18) - s(19)
                rsstok(14) = Null
                             
                rsstok(16) = s(14) + s(17) - s(18) - s(19)
                rsstok(17) = 0
                rsstok(18) = 0
                rsstok(19) = 0
                rsstok(20) = s(14) + s(17) - s(18) - s(19)
                rsstok(21) = Null
                
                Prg1.Value = j
                LblErrMsg.Caption = m & j
                j = j + 1
                rsstok.MoveNext
              Loop
              
              LblErrMsg.Caption = "Delete Stock Master"
              
'              Set cmd = New ADODB.Command
'              cmd.CommandType = adCmdStoredProc
'              cmd.CommandTimeout = 0
'              cmd.ActiveConnection = Db
'              cmd.CommandText = "SP_Delete_StockMaster"
'
'              cmd.Parameters.append cmd.CreateParameter("Period", adDBDate, adParamInput, , Format(Period, "yyyy-mm-01"))
'
'              Set rsSp = cmd.Execute
              
              sqlDelete = "delete from stock_master " & _
                          "where (lm_premonth is null or lm_premonth=0) and (lm_receipt is null or lm_receipt=0) and " & _
                          "(lm_supply is null or lm_supply=0) and (lm_lossreject is null or lm_lossreject=0) and " & _
                          "(lm_current is null or lm_current=0) and (lm_inventory is null or lm_inventory=0) and " & _
                          "(tm_premonth is null or tm_premonth=0) and (tm_receipt is null or tm_receipt=0) and " & _
                          "(tm_supply is null or tm_supply=0) and (tm_lossreject is null or tm_lossreject=0) and " & _
                          "(tm_current is null or tm_current=0) and (tm_inventory is null or tm_inventory=0) and " & _
                          "(nm_premonth is null or nm_premonth=0) and (nm_receipt is null or nm_receipt=0) and " & _
                          "(nm_supply is null or nm_supply=0) and (nm_lossreject is null or nm_lossreject=0) and " & _
                          "(nm_current is null or nm_current=0) and (nm_inventory is null or nm_inventory=0) " & _
                          "and isnull((select datediff(m,max(receipt_date),'" & Format(period, "yyyy-mm-01") & "') from part_receipt where item_code=stock_master.item_code and " & _
                          "warehouse_code=stock_master.warehouse_code),4) > 3 " & _
                          "and isnull((select datediff(m,max(childsupply_date),'" & Format(period, "yyyy-mm-01") & "') from part_supply where childitem_code=stock_master.item_code and " & _
                          "fromwarehouse_code=stock_master.warehouse_code),4) > 3 "

              db1.Execute sqlDelete
              
            End If
            
            
        End If
        
        If m = 0 Then Prg1.Value = 1
        
        If err.number = 0 Then
            LblErrMsg.Caption = "CommiTrans"
            db1.CommitTrans
            
            'Add to Inventory Control
            RS.AddNew
            RS("inventory_year") = Year(period)
            RS("inventory_month") = Month(period)
            RS("fix_cls") = 1
            RS("ClosingDate") = Now
            LblErrMsg.Caption = "inventory Control"
            RS.update
            
            'Clear MRP Setting
            sql = "Delete From MRP_Setting Where MRP_Year + MRP_Month < " & Format(period, "YYYYMM")
            Db.Execute sql
            
            LblErrMsg.Caption = DisplayMsg(4059)    'Inventory Closing Success !
            
        Else
        
            db1.RollbackTrans
            LblErrMsg.Caption = "[" & err.number & "] " & err.Description
            err.clear
        
        End If
            
        Set rsstok = Nothing
        Set RsItem = Nothing
        Set db1 = Nothing
        
    End If

End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
Dim temp As Date

LblErrMsg.Caption = ""
Prg1.Value = 0
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"

    sql = "select * from inventory_control where fix_cls='1'"
    If RS.State <> adStateClosed Then RS.Close
    RS.Open sql, Db, adOpenKeyset, adLockOptimistic
    
    If Not (RS.EOF And RS.BOF) Then
        RS.MoveLast
        temp = RS("inventory_year") & "-" & RS("inventory_month") & "-01"
        temp = DateAdd("m", 1, temp)
        period.Value = Format(temp, "MMM yyyy")
    Else
        period.Value = Format(Now, "MMM yyyy")
    End If
    temptgl = period.Month
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub CmdSubMenu_Click()
  frmMainMenu.Show
  Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If RS.State <> adStateClosed Then RS.Close
    If rsstok.State <> adStateClosed Then rsstok.Close
End Sub

Private Sub period_Change()
LblErrMsg.Caption = ""
Prg1.Value = 0
Call period_Click
temptgl = period.Month
End Sub

Private Sub period_Click()
    If period.Month = 1 And Val(temptgl) = 12 Then period.Year = period.Year + 1
    If period.Month = 12 And Val(temptgl) = 1 Then period.Year = period.Year - 1
End Sub

