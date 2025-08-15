VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmValuationPriceCalculation 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Valuation Price Calculation"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8415
   Icon            =   "FrmValuationPriceCalculation.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   8415
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCalculation 
      BackColor       =   &H0080FFFF&
      Caption         =   "Calculation Preparation"
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
      Index           =   2
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4200
      Width           =   2340
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H0080FFFF&
      Caption         =   "Cancel"
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
      Left            =   5820
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4200
      Width           =   1140
   End
   Begin VB.PictureBox PicProg 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   300
      ScaleHeight     =   555
      ScaleWidth      =   7860
      TabIndex        =   6
      Top             =   2670
      Width           =   7890
      Begin MSComctlLib.ProgressBar Prg1 
         Height          =   375
         Left            =   90
         TabIndex        =   7
         Top             =   90
         Width           =   7665
         _ExtentX        =   13520
         _ExtentY        =   661
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Max             =   20
         Scrolling       =   1
      End
   End
   Begin VB.CommandButton cmd 
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
      Index           =   0
      Left            =   7050
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4200
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   300
      TabIndex        =   4
      Top             =   3510
      Width           =   7890
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
         Left            =   90
         TabIndex        =   5
         Top             =   180
         Width           =   7650
      End
   End
   Begin VB.CommandButton cmdsub 
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
      Left            =   270
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4170
      Width           =   1140
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   6390
      TabIndex        =   11
      Top             =   480
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   714
   End
   Begin MSComCtl2.DTPicker dt 
      Height          =   345
      Left            =   1860
      TabIndex        =   0
      Top             =   2160
      Width           =   1755
      _ExtentX        =   3096
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
      Format          =   151781379
      UpDown          =   -1  'True
      CurrentDate     =   37831
   End
   Begin VB.Label lblProgress 
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
      Left            =   300
      TabIndex        =   13
      Top             =   2835
      Width           =   60
   End
   Begin VB.Label LblMonth 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Period "
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
      Left            =   300
      TabIndex        =   12
      Top             =   2235
      Width           =   600
   End
   Begin VB.Label lblBase 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
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
      Height          =   255
      Left            =   1890
      TabIndex        =   10
      Top             =   1335
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Base Currency :"
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
      Height          =   255
      Left            =   300
      TabIndex        =   9
      Top             =   1335
      Width           =   1665
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Valuation Price Calculation"
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
      Left            =   330
      TabIndex        =   8
      Top             =   480
      Width           =   7860
   End
End
Attribute VB_Name = "FrmValuationPriceCalculation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dateUp As Date
Dim blnCancel As Boolean

Private Sub cmd_Click(Index As Integer)
    
 If hakUpdate(Me.Name) = 0 Then LblErrMsg = DisplayMsg(3008): Exit Sub

    Me.MousePointer = vbHourglass
    
    Select Case Index
    Case 0: ValuationPriceCalculation
    Case 1: blnCancel = (MsgBox("Are you sure want to cancel transfer proccess?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation") = vbYes)
    End Select
    
    cmd(0).Enabled = False
    cmdCalculation(2).Enabled = True
    
    Me.MousePointer = vbDefault
End Sub

Private Sub ValuationPriceCalculation()

 Dim adoCon As New ADODB.Connection
 Dim adoCmd As New Command
 Dim RS As New ADODB.Recordset
 Dim spwhat As String
 Dim cek As String
 Dim rsCek As New ADODB.Recordset
    
 LblErrMsg.Caption = ""
 '   di remarks dulu karena mau calculate ulang 20221123
 ' LblErrMsg.Caption = up_ValidateDateRange(Format(dt.Value, "yyyy-MM-dd"), True)
 ' If LblErrMsg.Caption <> "" Then GoTo ErrExit
 '   di remarks dulu karena mau calculate ulang 20221123
    
 SetControl False
       
 'Check whether the chosen period has already been calculated or not
 If RS.State <> adStateClosed Then RS.Close
 RS.Open "select top 1 * from inventory_price where " & _
         "inventory_year = '" & Format(dt.Value, "yyyy") & "' " & _
         "and inventory_month = '" & Format(dt.Value, "mm") & "' ", Db, adOpenDynamic, adLockOptimistic
 If RS.EOF = False Then
  If MsgBox("Calculation of this period has been stored in Database." & vbCrLf & "Do you want to recalculate data of this period ?", vbExclamation + vbYesNo, "Confirmation") = vbNo Then
   GoTo ErrExit
  End If
 End If
    
 If RS.State <> adStateClosed Then RS.Close
        
    Prg1.Value = 0
    blnCancel = False
          
    'New Connection
    adoCon.ConnectionString = Db.ConnectionString
    adoCon.CursorLocation = adUseClient
    adoCon.CommandTimeout = 12000
    adoCon.Open
    adoCon.BeginTrans
    '*****************************
      
    'Material calculation first, then FG / WIP by the lowest level until the highest level
    RS.CursorLocation = adUseClient
    If RS.State <> adStateClosed Then RS.Close
    
    '20231017 Request Pak Toha, yg dikalkulasi item yg ada di stock master
    
    'rs.Open "WhatFirst ('" & DateAdd("d", -1, DateAdd("m", 1, Format(dt, "yyyy-MM-01"))) & "')", adoCon, adOpenDynamic, adLockReadOnly, adCmdStoredProc
    RS.Open " SELECT CASE WHEN FinishGoodPart_Cls='02' THEN 0 ELSE 1 END IDX, IM.Item_Code, Use_EndDay FROM Item_Master IM " & _
            " INNER JOIN (SELECT DISTINCT Item_Code FROM dbo.Stock_Master) SM ON IM.Item_Code = SM.Item_Code WHERE FinishGoodPart_Cls<>'2' ORDER BY FinishGoodPart_Cls DESC", adoCon, adOpenDynamic, adLockReadOnly
    
    'rs.Open "WhatFirstForOneItem ('" & DateAdd("d", -1, DateAdd("m", 1, Format(dt, "yyyy-MM-01"))) & "', 'E-MEG-100')", adoCon, adOpenDynamic, adLockReadOnly, adCmdStoredProc
    
    If Not RS.EOF Then
        
        Prg1.Min = 0
        Prg1.Max = RS.RecordCount
        lblProgress.Caption = ""
                                          
        While RS.EOF = False
                    
            Prg1.Value = RS.AbsolutePosition
            lblProgress.Caption = "Calculating for item : " & Trim(RS!Item_Code)
                       
            'Check if process canceled by user
            DoEvents
            If blnCancel Then
                adoCon.RollbackTrans
                lblProgress.Caption = "Valuation Calculation Progress : Canceled by user."
                LblErrMsg.Caption = DisplayMsg(8097)
                GoTo ErrExit
            End If
            
            If Trim(RS!Idx) = "0" Then
                spwhat = "AvgPriceMaterial"
                adoCmd.ActiveConnection = adoCon
                adoCmd.CommandTimeout = 12000
                adoCmd.CommandType = adCmdStoredProc
                adoCmd.CommandText = spwhat
                adoCmd.Parameters(1) = Trim(RS!Item_Code) '
                adoCmd.Parameters(2) = Format(dt, "yyyy-MM-dd")
                adoCmd.Parameters(3) = gi_decimalDigitPrice '5
                adoCmd.Parameters(4) = gi_decimalDigitPriceIDR '2
                adoCmd.Parameters(5) = gi_decimalDigitAmountIDR '2
                adoCmd.Parameters(6) = gi_decimalDigitQtyBOM '5
                adoCmd.Parameters(7) = gi_decimalDigitExchangeRate '2
                adoCmd.Parameters(8) = Trim(userLogin)
                
                adoCmd.Execute
            ElseIf Trim(RS!Idx) = "1" Then
                cek = "select FinishGoodPart_Cls from item_master where item_code='" & RS!Item_Code & "'"
                'cek = "select FinishGoodPart_Cls from item_master where item_code='KI3-230698'"
                Set rsCek = Db.Execute(cek)
                If rsCek!finishgoodpart_cls = "01" Then
                    spwhat = "AvgPriceFG"
                    adoCmd.ActiveConnection = adoCon
                    adoCmd.CommandTimeout = 12000
                    adoCmd.CommandType = adCmdStoredProc
                    adoCmd.CommandText = spwhat
                    adoCmd.Parameters(1) = Trim(RS!Item_Code)
                    adoCmd.Parameters(2) = Format(dt, "yyyy-MM-dd")
                    adoCmd.Parameters(3) = gi_decimalDigitPrice
                    adoCmd.Parameters(4) = gi_decimalDigitPriceIDR
                    adoCmd.Parameters(5) = gi_decimalDigitAmountIDR
                    adoCmd.Parameters(6) = gi_decimalDigitQtyBOM
                    adoCmd.Parameters(7) = gi_decimalDigitExchangeRate
                    adoCmd.Parameters(8) = Trim(userLogin)
                    
                    adoCmd.Execute
                Else
                    spwhat = "AvgPriceWIP_ExcludeDuty"
                    adoCmd.ActiveConnection = adoCon
                    adoCmd.CommandTimeout = 12000
                    adoCmd.CommandType = adCmdStoredProc
                    adoCmd.CommandText = spwhat
                    adoCmd.Parameters(1) = Trim(RS!Item_Code)
                    adoCmd.Parameters(2) = Format(dt, "yyyy-MM-dd")
                    adoCmd.Parameters(3) = gi_decimalDigitPrice
                    adoCmd.Parameters(4) = gi_decimalDigitPriceIDR
                    adoCmd.Parameters(5) = gi_decimalDigitAmountIDR
                    adoCmd.Parameters(6) = gi_decimalDigitQtyBOM
                    adoCmd.Parameters(7) = gi_decimalDigitExchangeRate
                    adoCmd.Parameters(8) = Trim(userLogin)
                    
                    adoCmd.Execute
                    'While adoCmd.State = adStateExecuting dimatiin looping nya
                    'Wend
                    spwhat = "AvgPriceWIP_IncludeDuty"
                    adoCmd.ActiveConnection = adoCon
                    adoCmd.CommandTimeout = 12000
                    adoCmd.CommandType = adCmdStoredProc
                    adoCmd.CommandText = spwhat
                    adoCmd.Parameters(1) = Trim(RS!Item_Code)
                    adoCmd.Parameters(2) = Format(dt, "yyyy-MM-dd")
                    adoCmd.Parameters(3) = gi_decimalDigitPrice
                    adoCmd.Parameters(4) = gi_decimalDigitPriceIDR
                    adoCmd.Parameters(5) = gi_decimalDigitAmountIDR
                    adoCmd.Parameters(6) = gi_decimalDigitQtyBOM
                    adoCmd.Parameters(7) = gi_decimalDigitExchangeRate
                    adoCmd.Parameters(8) = Trim(userLogin)
                    
                    adoCmd.Execute
               
                End If
            End If
                        
'            adoCmd.ActiveConnection = adoCon
'            adoCmd.CommandTimeout = 120
'            adoCmd.CommandType = adCmdStoredProc
'            adoCmd.CommandText = spwhat
'            adoCmd.Parameters(1) = Trim(rs!item_code)
'            adoCmd.Parameters(2) = Format(dt, "yyyy-MM-dd")
'            adoCmd.Parameters(3) = gi_decimalDigitPrice
'            adoCmd.Parameters(4) = gi_decimalDigitPriceIDR
'            adoCmd.Parameters(5) = gi_decimalDigitAmountIDR
'            adoCmd.Parameters(6) = gi_decimalDigitQtyBOM
'            adoCmd.Parameters(7) = gi_decimalDigitExchangeRate
'            adoCmd.Parameters(8) = Trim(userLogin)
'            adoCmd.Execute
            
            'While adoCmd.State = adStateExecuting dimatiin looping nya
            'Wend
                                  
            RS.MoveNext
        Wend
        
        adoCon.CommitTrans
             
        lblProgress.Caption = "Completed..."
        LblErrMsg = DisplayMsg(8058) '"Valuation Price Calculation Success!"
    End If
    If RS.State <> adStateClosed Then RS.Close
            
ErrExit:
    Set RS = Nothing
    Set adoCmd = Nothing
    Set adoCon = Nothing
    SetControl True
    Me.MousePointer = vbDefault
    Exit Sub
errHandler:
    adoCon.RollbackTrans
    LblErrMsg.Caption = "[" & err.number & "] " & err.Description
    err.clear
    Resume ErrExit
End Sub

Private Sub SetControl(booStatus As Boolean)
    Prg1.Value = 0
    cmdSub.Enabled = booStatus
    cmd(0).Enabled = booStatus
    cmd(1).Enabled = Not booStatus
    dt.Enabled = booStatus
    If cmd(0).Enabled = True Then
        cmdCalculation(2).Enabled = False
    Else
        cmdCalculation(2).Enabled = True
    End If
End Sub

Private Sub cmdCalculation_Click(Index As Integer)
Dim RS As ADODB.Recordset
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
    
    Me.MousePointer = vbHourglass
'   di remarks dulu karena mau calculate ulang 20221123
'    LblErrMsg.Caption = up_ValidateDateRange(Format(dt.Value, "yyyy-MM-dd"), True)
'    If LblErrMsg.Caption <> "" Then GoTo ErrExit
'   di remarks dulu karena mau calculate ulang 20221123
      
    'SetControl False
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "SP_PreparationCalculation"
    
    cmd.Parameters.append cmd.CreateParameter("StartDate", adDate, adParamInput, , Format(dt, "yyyy-MM-dd"))
    
    Set RS = cmd.Execute
    
    Me.MousePointer = vbDefault
    
    LblErrMsg.Caption = DisplayMsg(9012)
    
    cmdCalculation(2).Enabled = False
    
    SetControl True
   
   
ErrExit:
    Set RS = Nothing
    Set cmd = Nothing
    'SetControl False
    Me.MousePointer = vbDefault
    Exit Sub
    
End Sub

Private Sub cmdsub_Click()
    DoEvents
    frmMainMenu.Show
    DoEvents
    Unload Me
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
    If ErrMsg = "" Then
        Unload Me
    Else
        LblErrMsg.Caption = ErrMsg
    End If
End Sub

Private Sub dt_change()
    lblProgress.Caption = ""
    LblErrMsg.Caption = ""

    If Format(dt.Value, "MM") < Format(dateUp, "MM") And Val(Format(dt.Value, "MM")) = 1 And Val(Format(dateUp, "MM")) = 12 Then dt.Year = dt.Year + 1: GoTo pass
    If Format(dt.Value, "MM") > Format(dateUp, "MM") And Val(Format(dt.Value, "MM")) = 12 And Val(Format(dateUp, "MM")) = 1 Then dt.Year = dt.Year - 1
pass:
    dateUp = Format(dt.Value, "dd MMM yyyy")
       
End Sub

Private Sub Form_Load()
    Dim RS As New ADODB.Recordset
    If gb_Simulation = True Then Call up_InitSimulation(Me)
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
    dt.Value = Format(Now, "MMM yyyy")
    dateUp = dt.Value
    cmd(0).Enabled = False
    
    
    RS.Open "select valuationprice_BaseCurrency from Company_Profile", Db, adOpenForwardOnly, adLockReadOnly
    lblBase.Caption = uf_GetCurrencyDescription(Trim$(RS(0) & ""))
    RS.Close
    Set RS = Nothing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = 1
End Sub

