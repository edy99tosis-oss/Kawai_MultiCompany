VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMRPCalculation 
   BackColor       =   &H00FDDFE3&
   Caption         =   "MRP Calculation"
   ClientHeight    =   4380
   ClientLeft      =   690
   ClientTop       =   3630
   ClientWidth     =   8070
   Icon            =   "frmMRPCalculation.frx":0000
   LinkMode        =   1  'Source
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   8070
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
      Left            =   394
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3480
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   394
      TabIndex        =   8
      Top             =   2820
      Width           =   7275
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
         Height          =   240
         Left            =   105
         TabIndex        =   9
         Top             =   195
         Width           =   7050
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
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
      Index           =   0
      Left            =   6529
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3495
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Ca&ncel"
      Enabled         =   0   'False
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
      Left            =   5299
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3495
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   855
      Left            =   394
      TabIndex        =   7
      Top             =   1020
      Width           =   7275
      Begin MSComCtl2.DTPicker dt 
         Height          =   315
         Left            =   3540
         TabIndex        =   0
         Top             =   315
         Width           =   1260
         _ExtentX        =   2223
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
         CustomFormat    =   "MMM yyyy"
         Format          =   151453699
         UpDown          =   -1  'True
         CurrentDate     =   37831
      End
      Begin VB.Label LblMonth 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Calculate up to :"
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
         Left            =   1965
         TabIndex        =   12
         Top             =   375
         Width           =   1425
      End
   End
   Begin VB.PictureBox PicProg 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   394
      ScaleHeight     =   495
      ScaleWidth      =   7245
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2010
      Width           =   7275
      Begin MSComctlLib.ProgressBar Prg1 
         Height          =   315
         Left            =   105
         TabIndex        =   11
         Top             =   90
         Width           =   7020
         _ExtentX        =   12383
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   5820
      TabIndex        =   4
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
      Height          =   195
      Left            =   405
      TabIndex        =   10
      Top             =   2610
      Width           =   60
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MRP Calculation"
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
      Left            =   405
      TabIndex        =   6
      Top             =   435
      Width           =   7275
   End
End
Attribute VB_Name = "frmMRPCalculation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim adoCon As New ADODB.Connection

Dim i As Integer, sql As String
Dim blnCancel As Boolean
Dim inputParent As String, lotno As String
Dim tglProd As String, qtyParent As Double
Dim factoryCD As String, lineCD As String
Dim TampungDt As Byte

Dim totProd As Double, totOffQty As Double
Dim startDaily As String

Private Function fc_CheckSQL(strTemp As String) As String
    
    Dim strResult As String
    
    strResult = Replace$(strTemp, Chr(34), Chr(34) & Chr(34))
    strResult = Replace$(strResult, Chr(39), Chr(39) & Chr(39))
        
    fc_CheckSQL = strResult
    
End Function

Private Sub SetControl(booStatus As Boolean)
    
    Prg1.Value = 0
    cmdsubmenu.Enabled = booStatus
    command1(0).Enabled = booStatus
    command1(1).Enabled = Not booStatus
    dt.Enabled = booStatus
    
End Sub

Private Sub MRPCalculation()
    
    Dim adoRsItem As New ADODB.Recordset
    Dim adoRsMRPSetting As New ADODB.Recordset
    Dim adoRsActProd As New ADODB.Recordset
    Dim adoCmd As New Command
    
    On Error GoTo errHandler
    Me.MousePointer = vbHourglass
        
    LblErrMsg.Caption = ""
    SetControl False
    blnCancel = False
    
    Dim strItem_Code As String
    Dim strFactory_Code As String
    Dim strLine_Code As String
    Dim strUnit_Cls As String
    Dim strLotNo As String
    Dim strActual_Cls As String
    Dim strComplete_Cls As String
    Dim dblQty As Double
    Dim dblOffQty As Double
    Dim dblQtyProd As Double
    Dim dtActual_Date As Date
    
    Dim booForecastBig As Boolean
    
    If MsgBox("Do you really want to Process MRP Calculation?", vbQuestion & vbYesNo, "Confirmation") = vbNo Then GoTo ErrExit
        
    'Open New Connection
    adoCon.ConnectionString = Db.ConnectionString
    adoCon.CursorLocation = adUseClient
    adoCon.Open
    
    adoCon.BeginTrans
    
    'Clean Actual Production
    adoCmd.ActiveConnection = adoCon
    adoCmd.CommandType = adCmdStoredProc
    adoCmd.CommandText = "mrp_actual_production_clean"
    adoCmd.Parameters(1) = Year(dt)
    adoCmd.Parameters(2) = Month(dt)
    adoCmd.Execute
    While adoCmd.State = adStateExecuting
    Wend
    Set adoCmd = Nothing
        
    'Get item(s)
    adoRsItem.Open "mrp_item_view", adoCon, adOpenDynamic, adLockReadOnly, adCmdStoredProc '
    If Not adoRsItem.EOF Then
            
        Prg1.Max = adoRsItem.RecordCount
        
        While adoRsItem.EOF = False
                        
            Prg1.Value = adoRsItem.AbsolutePosition
            lblKet.Caption = "MRP Calculation Progress : Insert Actual Production... " & Round((adoRsItem.AbsolutePosition / Prg1.Max) * 100, 0) & " %"
            
            'Check if process canceled by user
            DoEvents
            If blnCancel Then
                adoCon.RollbackTrans
                lblKet.Caption = "MRP Calculation Progress : Canceled by user."
                LblErrMsg.Caption = DisplayMsg(8097)
                GoTo ErrExit
            End If
                        
            'Get MRP Setting
            sql = "Select MRP_Year, MRP_Month From MRP_Setting Where MRP_Year + MRP_Month <= " & Format(dt.Value, "YYYYMM") & "Order By MRP_Year + MRP_Month"
            adoRsMRPSetting.Open sql, adoCon, adOpenForwardOnly, adLockReadOnly, adCmdText
            While adoRsMRPSetting.EOF = False
                
                If Trim$(adoRsItem.Fields("Use_EndDay")) > Format(DateSerial(adoRsMRPSetting.Fields("MRP_Year"), adoRsMRPSetting.Fields("MRP_Month"), 1), "yyyyMMdd") Then
                
                    'Get and Insert Actual Production
                    adoRsActProd.Open "mrp_actual_production (" & adoRsMRPSetting.Fields("MRP_Year") & ", " & adoRsMRPSetting.Fields("MRP_Month") & ", '" & fc_CheckSQL(Trim$(adoRsItem.Fields("Item_Code"))) & "')", adoCon, adOpenDynamic, adLockReadOnly, adCmdStoredProc
                    While adoRsActProd.EOF = False
                        adoCmd.ActiveConnection = adoCon
                        adoCmd.CommandType = adCmdStoredProc
                        adoCmd.CommandText = "mrp_actual_production_save"
                        adoCmd.Parameters(1) = Trim$(adoRsActProd.Fields("Item_Code"))
                        adoCmd.Parameters(2) = adoRsActProd.Fields("Lot_No")
                        adoCmd.Parameters(3) = Format(adoRsActProd.Fields("Actual_Date"), "yyyy-MM-dd")
                        adoCmd.Parameters(4) = adoRsActProd.Fields("Actual_Cls")
                        adoCmd.Parameters(5) = adoRsActProd.Fields("off_qty")
                        adoCmd.Parameters(6) = adoRsActProd.Fields("Qty")
                        adoCmd.Parameters(7) = adoRsActProd.Fields("QtyProd")
                        adoCmd.Execute
                        While adoCmd.State = adStateExecuting
                        Wend
                        Set adoCmd = Nothing
                        adoRsActProd.MoveNext
                    Wend
                    
                    adoRsActProd.Close
                    Set adoRsActProd = Nothing
                                    
                End If
                
                adoRsMRPSetting.MoveNext
                
            Wend
            
            adoRsMRPSetting.Close
            Set adoRsMRPSetting = Nothing
            adoRsItem.MoveNext
        
        Wend
        
        adoRsItem.Close
        Set adoRsItem = Nothing
    
    End If
    
    'Get MRP Setting
    sql = "Select MRP_Year, MRP_Month From MRP_Setting Where MRP_Year + MRP_Month <= " & Format(dt.Value, "YYYYMM") & "Order By MRP_Year + MRP_Month"
    adoRsMRPSetting.Open sql, adoCon, adOpenForwardOnly, adLockReadOnly, adCmdText
    While adoRsMRPSetting.EOF = False

        'Get Actual Production and insert actual production WIP
        adoRsActProd.Open "mrp_actual_production_view (" & adoRsMRPSetting.Fields("MRP_Year") & ", " & adoRsMRPSetting.Fields("MRP_Month") & ", Null, 'W')", adoCon, adOpenDynamic, adLockReadOnly, adCmdStoredProc
        If Not adoRsActProd.EOF Then

            Prg1.Max = adoRsActProd.RecordCount
            While adoRsActProd.EOF = False

                Prg1.Value = adoRsActProd.AbsolutePosition
                lblKet.Caption = "MRP Calculation Progress : Insert Actual Production WIP (" & Format(DateSerial(adoRsMRPSetting.Fields("MRP_Year"), adoRsMRPSetting.Fields("MRP_Month"), 1), "MMMM yyyy") & ") ... " & Round((adoRsActProd.AbsolutePosition / Prg1.Max) * 100, 0) & " %"

                'Check if process canceled by user
                DoEvents
                If blnCancel Then
                    adoCon.RollbackTrans
                    lblKet.Caption = "MRP Calculation Progress : Canceled by user."
                    LblErrMsg.Caption = DisplayMsg(8097)
                    GoTo ErrExit
                End If
                
                adoCmd.ActiveConnection = adoCon
                adoCmd.CommandType = adCmdStoredProc
                adoCmd.CommandText = "mrp_actual_production_save"
                adoCmd.Parameters(1) = Trim(adoRsActProd.Fields("Item_Code"))
                adoCmd.Parameters(2) = ""
                adoCmd.Parameters(3) = Format(adoRsActProd.Fields("Actual_Date"), "yyyy-MM-dd")
                adoCmd.Parameters(4) = "W"
                adoCmd.Parameters(5) = 0
                adoCmd.Parameters(6) = adoRsActProd.Fields("Qty") - adoRsActProd.Fields("Qty_Daily")
                adoCmd.Parameters(7) = 0
                adoCmd.Execute
                While adoCmd.State = adStateExecuting
                Wend
                Set adoCmd = Nothing

                If adoRsActProd.Fields("Qty") > adoRsActProd.Fields("Qty_Daily") Then
                    If Not RecursifActualProductionWIP(Trim(adoRsActProd.Fields("Item_Code")), adoRsActProd.Fields("Actual_Date"), 1, adoRsActProd.Fields("Qty"), adoRsActProd.Fields("Qty_Daily")) Then
                        adoCon.RollbackTrans
                        lblKet.Caption = "MRP Calculation Progress : Calculation error."
                        GoTo ErrExit
                    End If
                Else
                    If Not RecursifActualProductionWIP(Trim(adoRsActProd.Fields("Item_Code")), adoRsActProd.Fields("Actual_Date"), 1, adoRsActProd.Fields("Qty_Daily"), adoRsActProd.Fields("Qty_Daily")) Then
                        adoCon.RollbackTrans
                        lblKet.Caption = "MRP Calculation Progress : Calculation error."
                        GoTo ErrExit
                    End If
                End If

                adoRsActProd.MoveNext

            Wend

        End If

        adoRsActProd.Close
        Set adoRsActProd = Nothing
        adoRsMRPSetting.MoveNext

    Wend
    adoRsMRPSetting.Close

    'Clean Requirement Master & Requirement
    adoCmd.ActiveConnection = adoCon
    adoCmd.CommandType = adCmdStoredProc
    adoCmd.CommandText = "mrp_requirement_clean"
    adoCmd.Parameters(1) = Year(dt)
    adoCmd.Parameters(2) = Month(dt)
    adoCmd.Execute
    While adoCmd.State = adStateExecuting
    Wend
    Set adoCmd = Nothing

    'Get MRP Setting
    sql = "Select MRP_Year, MRP_Month From MRP_Setting Where MRP_Year + MRP_Month <= " & Format(dt.Value, "YYYYMM") & "Order By MRP_Year + MRP_Month"
    adoRsMRPSetting.Open sql, adoCon, adOpenForwardOnly, adLockReadOnly, adCmdText
    While adoRsMRPSetting.EOF = False

        'Get Actual Production and calculating requirement
        adoRsActProd.Open "mrp_actual_production_view (" & adoRsMRPSetting.Fields("MRP_Year") & ", " & adoRsMRPSetting.Fields("MRP_Month") & ", Null, Null)", adoCon, adOpenDynamic, adLockReadOnly, adCmdStoredProc
        If Not adoRsActProd.EOF Then

            Prg1.Max = adoRsActProd.RecordCount
            While adoRsActProd.EOF = False

                Prg1.Value = adoRsActProd.AbsolutePosition
                lblKet.Caption = "MRP Calculation Progress : Calculating requirement (" & Format(DateSerial(adoRsMRPSetting.Fields("MRP_Year"), adoRsMRPSetting.Fields("MRP_Month"), 1), "MMMM yyyy") & ") ... " & Round((adoRsActProd.AbsolutePosition / Prg1.Max) * 100, 0) & " %"

                'Check if process canceled by user
                DoEvents
                If blnCancel Then
                    adoCon.RollbackTrans
                    lblKet.Caption = "MRP Calculation Progress : Canceled by user."
                    LblErrMsg.Caption = DisplayMsg(8097)
                    GoTo ErrExit
                End If

                strItem_Code = adoRsActProd.Fields("Item_Code")
                strFactory_Code = adoRsActProd.Fields("Manufacture_Code") & ""
                strLine_Code = adoRsActProd.Fields("Line_Code") & ""
                strUnit_Cls = adoRsActProd.Fields("Unit_Cls")
                strLotNo = adoRsActProd.Fields("Lot_No")
                strActual_Cls = adoRsActProd.Fields("Actual_Cls")
                strComplete_Cls = "0" 'adoRs.Fields("Complete_Cls")
                dtActual_Date = adoRsActProd.Fields("Actual_Date")
                dblQty = adoRsActProd.Fields("Qty")
                dblOffQty = adoRsActProd.Fields("Off_Qty")
                dblQtyProd = adoRsActProd.Fields("QtyProd")

                booForecastBig = (adoRsActProd.Fields("TotForecast") > adoRsActProd.Fields("TotDaily"))

                If Not RecursiveInsertRequirement(strItem_Code, strItem_Code, strFactory_Code, strLine_Code, strUnit_Cls, _
                    strLotNo, strActual_Cls, strComplete_Cls, dtActual_Date, dblQty, dblOffQty, dblQtyProd, 1, booForecastBig) Then
                    adoCon.RollbackTrans
                    lblKet.Caption = "MRP Calculation Progress : Calculation error."
                    GoTo ErrExit
                End If

                adoRsActProd.MoveNext

            Wend

        End If

        adoRsActProd.Close
        Set adoRsActProd = Nothing
        adoRsMRPSetting.MoveNext

    Wend
    adoRsMRPSetting.Close
    
    adoCon.CommitTrans
    lblKet.Caption = "MRP Calculation Progress : Calculation completed."
    LblErrMsg.Caption = DisplayMsg("0062")
    
    adoCon.Close
    
ErrExit:
    SetControl True
    Set adoRsItem = Nothing
    Set adoRsMRPSetting = Nothing
    Set adoRsActProd = Nothing
    Set adoCmd = Nothing
    Set adoCon = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
errHandler:
    adoCon.RollbackTrans
    LblErrMsg.Caption = "[" & err.number & "] " & err.Description
    err.clear
    Resume ErrExit
    
End Sub

Private Function RecursiveInsertRequirement(strItem_Code As String, StrParent As String, strFactory_Code As String, strLine_Code As String, _
     strUnit_Cls As String, strLotNo As String, strActual_Cls As String, strComplete_Cls As String, dtActual_Date As Date, _
     dblQty As Double, dblOffQty As Double, dblQtyProd As Double, dblQtyParent As Double, booForecastBig As Boolean) As Boolean
        
    Dim adoRs As New ADODB.Recordset
    Dim adoCmd As New ADODB.Command
    Dim adosubcon As New ADODB.Recordset
    Dim sql As String
    
    On Error GoTo errHandler
    
    adoRs.Open "mrp_actual_production_view (" & Year(dtActual_Date) & ", " & Month(dtActual_Date) & ", '" & Trim(fc_CheckSQL(strItem_Code)) & "', Null)", adoCon, adOpenDynamic, adLockReadOnly, adCmdStoredProc
    While adoRs.EOF = False
    
        If Format(dtActual_Date, "YYYYMMDD") >= adoRs.Fields("Start_Date") And Format(dtActual_Date, "YYYYMMDD") <= adoRs.Fields("End_Date") Then
        
            If adoRs.Fields("Production_Cls") = "01" Then
            
                adoCmd.ActiveConnection = adoCon
                adoCmd.CommandType = adCmdStoredProc
                adoCmd.CommandText = "mrp_requirement_save"
                adoCmd.Parameters(1) = StrParent
                adoCmd.Parameters(2) = Trim$(strLotNo)
                adoCmd.Parameters(3) = Trim$(strFactory_Code)
                adoCmd.Parameters(4) = Trim$(strLine_Code)
                adoCmd.Parameters(5) = dtActual_Date
                adoCmd.Parameters(6) = dblQty
                adoCmd.Parameters(7) = dblOffQty
                adoCmd.Parameters(8) = dblQtyProd
                adoCmd.Parameters(9) = Trim$(strComplete_Cls)
                adoCmd.Parameters(10) = Trim$(adoRs.Fields("Item_Code"))
                adoCmd.Parameters(11) = dblQtyParent * adoRs.Fields("Qty")
                adoCmd.Parameters(12) = Trim$(adoRs.Fields("Unit_Cls"))
                adoCmd.Parameters(13) = userLogin
                adoCmd.Execute
                While adoCmd.State = adStateExecuting
                Wend
                Set adoCmd = Nothing
            
            ' Option for SubConItem Material Request ' For KAWAI
                
'                Sql = " Select Trade_Cls From Item_Master im  " & vbCrLf & _
'                            "   Inner Join Trade_Master TM On IM.Supplier_Code=TM.Trade_Code " & vbCrLf & _
'                            "       Where Item_Code='" & adoRs.Fields("Item_Code") & "' "
'
'                If adosubcon.State <> adStateClosed Then adosubcon.Close
'                Set adosubcon = adoCon.Execute(Sql)
'
'                If Not adosubcon.EOF And adosubcon.Fields(0) = "3" Then
'                    RecursiveInsertRequirement Trim$(adoRs.Fields("Item_Code")), StrParent, strFactory_Code, strLine_Code, strUnit_Cls, _
'                        strLotNo, strActual_Cls, strComplete_Cls, dtActual_Date, dblQty, dblOffQty, dblQtyProd, dblQtyParent * adoRs.Fields("Qty"), booForecastBig
'                End If
            ' ------------------------
            
            Else
                
                If IsNull(adoRs.Fields("Item_Child")) Then
                
                    adoCmd.ActiveConnection = adoCon
                    adoCmd.CommandType = adCmdStoredProc
                    adoCmd.CommandText = "mrp_requirement_save"
                    adoCmd.Parameters(1) = StrParent
                    adoCmd.Parameters(2) = Trim$(strLotNo)
                    adoCmd.Parameters(3) = Trim$(strFactory_Code)
                    adoCmd.Parameters(4) = Trim$(strLine_Code)
                    adoCmd.Parameters(5) = dtActual_Date
                    adoCmd.Parameters(6) = dblQty
                    adoCmd.Parameters(7) = dblOffQty
                    adoCmd.Parameters(8) = dblQtyProd
                    adoCmd.Parameters(9) = Trim$(strComplete_Cls)
                    adoCmd.Parameters(10) = Trim$(adoRs.Fields("Item_Code"))
                    adoCmd.Parameters(11) = dblQtyParent * adoRs.Fields("Qty")
                    adoCmd.Parameters(12) = Trim$(adoRs.Fields("Unit_Cls"))
                    adoCmd.Parameters(13) = userLogin
                    adoCmd.Execute
                    While adoCmd.State = adStateExecuting
                    Wend
                    Set adoCmd = Nothing
                    
                Else
                    
                    RecursiveInsertRequirement Trim$(adoRs.Fields("Item_Code")), StrParent, strFactory_Code, strLine_Code, strUnit_Cls, _
                        strLotNo, strActual_Cls, strComplete_Cls, dtActual_Date, dblQty, dblOffQty, dblQtyProd, dblQtyParent * adoRs.Fields("Qty"), booForecastBig
                    
                End If
                
            End If
                        
        End If
        
        adoRs.MoveNext
                
    Wend
    adoRs.Close
    
    RecursiveInsertRequirement = True
    
ErrExit:
    Set adoRs = Nothing
    Set adoCmd = Nothing
    Exit Function
errHandler:
    LblErrMsg.Caption = "[" & err.number & "] " & err.Description
    err.clear
    RecursiveInsertRequirement = False
    Resume ErrExit
    
End Function

Private Function RecursifActualProductionWIP(strItem_Code As String, dtActual_Date As Date, dblQtyParent As Double, dblQty As Double, dblQtyDaily As Double) As Boolean
    
    Dim adoRs As New ADODB.Recordset
    Dim adoCmd As New ADODB.Command
        
    On Error GoTo errHandler
        
    adoRs.Open "mrp_actual_production_view (" & Year(dtActual_Date) & ", " & Month(dtActual_Date) & ", '" & Trim(fc_CheckSQL(strItem_Code)) & "', Null)", adoCon, adOpenDynamic, adLockReadOnly, adCmdStoredProc
    While adoRs.EOF = False
    
        If Format(dtActual_Date, "YYYYMMDD") >= adoRs.Fields("Start_Date") And Format(dtActual_Date, "YYYYMMDD") <= adoRs.Fields("End_Date") Then
        
            If adoRs.Fields("Production_Cls") = "01" Then
                
                adoCmd.ActiveConnection = adoCon
                adoCmd.CommandType = adCmdStoredProc
                adoCmd.CommandText = "mrp_actual_production_save"
                adoCmd.Parameters(1) = Trim(adoRs.Fields("Item_Code"))
                adoCmd.Parameters(2) = ""
                adoCmd.Parameters(3) = Format(dtActual_Date, "yyyy-MM-dd")
                adoCmd.Parameters(4) = "W"
                adoCmd.Parameters(5) = 0
                adoCmd.Parameters(6) = adoRs.Fields("Qty") * (dblQty - adoRs.Fields("Qty_Daily")) + adoRs.Fields("Qty_Order")
                adoCmd.Parameters(7) = 0
                adoCmd.Execute
                While adoCmd.State = adStateExecuting
                Wend
                Set adoCmd = Nothing
                
                If dblQty > adoRs.Fields("Qty_Daily") Then
                    RecursifActualProductionWIP Trim(adoRs.Fields("Item_Code")), dtActual_Date, adoRs.Fields("Qty"), (adoRs.Fields("Qty") * dblQty) + adoRs.Fields("Qty_Order"), adoRs.Fields("Qty_Daily")
                Else
                    RecursifActualProductionWIP Trim(adoRs.Fields("Item_Code")), dtActual_Date, adoRs.Fields("Qty"), (adoRs.Fields("Qty") * adoRs.Fields("Qty_Daily")) + adoRs.Fields("Qty_Order"), adoRs.Fields("Qty_Daily")
                End If
                
            ElseIf Not IsNull(adoRs.Fields("Item_Child")) Then
                
                If dblQty > adoRs.Fields("Qty_Daily") Then
                    RecursifActualProductionWIP Trim(adoRs.Fields("Item_Code")), dtActual_Date, adoRs.Fields("Qty"), (adoRs.Fields("Qty") * dblQty) + adoRs.Fields("Qty_Order"), adoRs.Fields("Qty_Daily")
                Else
                    RecursifActualProductionWIP Trim(adoRs.Fields("Item_Code")), dtActual_Date, adoRs.Fields("Qty"), (adoRs.Fields("Qty") * adoRs.Fields("Qty_Daily")) + adoRs.Fields("Qty_Order"), adoRs.Fields("Qty_Daily")
                End If
                
            End If
            
        End If
        
        adoRs.MoveNext
    
    Wend
    adoRs.Close
    
    RecursifActualProductionWIP = True
    
ErrExit:
    Set adoRs = Nothing
    Set adoCmd = Nothing
    Exit Function
errHandler:
    LblErrMsg.Caption = "[" & err.number & "] " & err.Description
    err.clear
    RecursifActualProductionWIP = False
    Resume ErrExit
    
End Function

Private Sub Command1_Click(Index As Integer)
    
    If hakUpdate(Me.Name) = 0 Then LblErrMsg = DisplayMsg(3008): Exit Sub

    Me.MousePointer = vbHourglass
    
    Select Case Index
    Case 0: MRPCalculation
    Case 1: blnCancel = (MsgBox("Are you sure want to cancel transfer proccess?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation") = vbYes)
    End Select
    
    Me.MousePointer = vbDefault

End Sub

Private Sub dt_change()
    Call dt_Click
    TampungDt = dt.Month
    LblErrMsg.Caption = ""
    lblKet.Caption = ""
End Sub

Private Sub dt_Click()
    If dt.Month = 1 And Val(TampungDt) = 12 Then dt.Year = dt.Year + 1
    If dt.Month = 12 And Val(TampungDt) = 1 Then dt.Year = dt.Year - 1
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
    Dim Ret As String, NC As Long, TempPWD As String

    dt = Format(Now, "MMM yyyy")
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
    
    Ret = String(255, 0)
    NC = GetPrivateProfileString("StartDaily", "Date", "", Ret, 255, IniFile)
    If NC <> 0 Then Ret = Left$(Ret, NC)
    startDaily = Ret
End Sub

Private Sub CmdSubMenu_Click()
    DoEvents
    frmMainMenu.Show
    DoEvents
    Unload Me
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
