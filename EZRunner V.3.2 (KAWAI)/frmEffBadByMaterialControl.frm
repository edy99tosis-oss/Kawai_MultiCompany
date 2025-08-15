VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEffBadByMaterialControl 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Defective Type By Material Pareto Diagram"
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8640
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEffBadByMaterialControl.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4350
   ScaleWidth      =   8640
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBrowser 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4635
      TabIndex        =   17
      Top             =   2107
      Width           =   300
   End
   Begin VB.CommandButton cmdSubMenu 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Sub &Menu"
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3690
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   360
      TabIndex        =   10
      Top             =   2970
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
         Left            =   120
         TabIndex        =   11
         Top             =   210
         Width           =   7215
      End
   End
   Begin VB.CommandButton cmdPreview 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Preview"
      Height          =   405
      Left            =   7065
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3690
      Width           =   1185
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   6405
      TabIndex        =   7
      Top             =   240
      Width           =   1845
      _extentx        =   3254
      _extenty        =   714
   End
   Begin MSComCtl2.DTPicker Period 
      Height          =   330
      Left            =   1800
      TabIndex        =   3
      Top             =   2490
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   582
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
      Format          =   293797891
      UpDown          =   -1  'True
      CurrentDate     =   37860
   End
   Begin VB.Label lblItemNm 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Left            =   5040
      TabIndex        =   16
      Top             =   2160
      Width           =   2265
   End
   Begin VB.Line Line8 
      Index           =   2
      X1              =   4995
      X2              =   8225
      Y1              =   2400
      Y2              =   2400
   End
   Begin MSForms.ComboBox cboItem 
      Height          =   315
      Left            =   1800
      TabIndex        =   2
      Top             =   2100
      Width           =   2760
      VariousPropertyBits=   746604571
      MaxLength       =   25
      DisplayStyle    =   3
      Size            =   "4868;556"
      ListRows        =   15
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      Value           =   "AAAAAAAAAAAAAAA"
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Material Code"
      Height          =   195
      Index           =   3
      Left            =   360
      TabIndex        =   15
      Top             =   2160
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Machine No"
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   14
      Top             =   1755
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Factory CD"
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   13
      Top             =   1350
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Period"
      Height          =   195
      Index           =   2
      Left            =   360
      TabIndex        =   12
      Top             =   2565
      Width           =   540
   End
   Begin VB.Line Line8 
      Index           =   0
      X1              =   3255
      X2              =   8250
      Y1              =   1590
      Y2              =   1590
   End
   Begin VB.Line Line8 
      Index           =   1
      X1              =   3270
      X2              =   5520
      Y1              =   1995
      Y2              =   1995
   End
   Begin VB.Label lblLine 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Left            =   3255
      TabIndex        =   9
      Top             =   1755
      Width           =   2265
   End
   Begin MSForms.ComboBox cboLine 
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Top             =   1695
      Width           =   1335
      VariousPropertyBits=   746604571
      MaxLength       =   10
      DisplayStyle    =   3
      Size            =   "2355;556"
      ListRows        =   15
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      Value           =   "AAA"
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cboFactory 
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Top             =   1290
      Width           =   1335
      VariousPropertyBits=   746604571
      MaxLength       =   10
      DisplayStyle    =   3
      Size            =   "2355;556"
      ListRows        =   15
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      Value           =   "AAAAAA"
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblFactory 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Left            =   3255
      TabIndex        =   8
      Top             =   1350
      Width           =   5025
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Defective Type By Material Pareto Diagram"
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
      Left            =   1875
      TabIndex        =   6
      Top             =   675
      Width           =   4875
   End
End
Attribute VB_Name = "frmEffBadByMaterialControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim temptgl As Byte

Public rsRecord As New ADODB.Recordset
Dim sqlLine As String, sqlitem As String, sqlItem2 As String

Private Sub cmdBrowser_Click()
 Me.MousePointer = vbHourglass
 frm_BrowseItem.getItemCode = CboItem.Text
 frm_BrowseItem.Show 1
 CboItem.Text = frm_BrowseItem.getItemCode
 Me.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"

    Call isiCbo(cboFactory, "Trade_Master", "Trade_Code", "Trade_Name", 80, 200, "Trade_Code", , , " trade_code in (select distinct manufacture_code from manufacture_line) ")
    Call isiCbo(CboItem, "Item_Master", "Item_Code", "Item_Name", 150, 250, "Item_Code", , , , 1)
    
    cboFactory.Text = "": lblFactory.Caption = ""
    CboLine.clear
    CboLine.Text = "": lblLine.Caption = ""
    CboItem = "": lblItemNm = ""
    Period = Format(Now, "yyyy-mm-01")
    temptgl = Period.Month
End Sub

Private Sub cboFactory_Change()
    If cboFactory.ListIndex <> -1 Then
        lblFactory.Caption = cboFactory.Column(1)
        Call isiCbo(CboLine, "Manufacture_Line", "Line_Code", "Line_Name", 80, 140, "Line_Code", , , " Manufacture_Code = '" & cboFactory & "' ", 1)
    Else
        If cboFactory.MatchFound = False Then lblFactory.Caption = "": CboLine.clear
    End If
End Sub

Private Sub cboFactory_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call cboFactory_Change
End Sub

Private Sub cboLine_Change()
    If CboLine.ListIndex <> -1 Then
        lblLine.Caption = CboLine.Column(1)
    Else
        If CboLine.MatchFound = False Then lblLine.Caption = ""
    End If
End Sub

Private Sub cboLine_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call cboLine_Change
End Sub

Private Sub CboItem_Change()
    If CboItem.MatchFound Then lblItemNm = CboItem.Column(1) Else lblItemNm = ""
End Sub

Private Sub cboitem_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call CboItem_Change
End Sub

Private Sub period_Change()
    LblErrMsg.Caption = ""
    Call period_Click
    temptgl = Period.Month
End Sub

Private Sub period_Click()
    If Period.Month = 1 And Val(temptgl) = 12 Then Period.Year = Period.Year + 1
    If Period.Month = 12 And Val(temptgl) = 1 Then Period.Year = Period.Year - 1
End Sub

Private Sub CmdPreview_Click()
    LblErrMsg = ""
    
'    If hakUpdate(Me.Name) = 0 Then _
'    LblErrMsg = DisplayMsg(1040): Me.MousePointer = vbDefault: Exit Sub

    'VALIDATION
    If cboFactory.Text = "" Then
        LblErrMsg = DisplayMsg(1060)    '"Please Select Factory Code"
        cboFactory.SetFocus
        Exit Sub
    ElseIf cboFactory.Text <> "" Then
        If cboFactory.MatchFound = False Then
            LblErrMsg = DisplayMsg(4016)    '"Record with this Factory Code not found"
            cboFactory.SetFocus
            Exit Sub
        End If
    End If
    If CboLine.Text = "" Then
        LblErrMsg = DisplayMsg(1041)    '"Please Input Line Code"
        CboLine.SetFocus
        Exit Sub
    ElseIf CboLine.Text <> "" Then
        If CboLine.MatchFound = False Then
            LblErrMsg = DisplayMsg(4017)    '"Record with this Line Code not found"
            CboLine.SetFocus
            Exit Sub
        End If
    End If
    
    If CboItem = "" Then
        LblErrMsg = DisplayMsg(1009) 'Please Input Product Code
        CboItem.SetFocus: Exit Sub
    ElseIf CboItem.MatchFound = False Then
        LblErrMsg = DisplayMsg(4061) 'Product Code not found!
        CboItem.SetFocus: Exit Sub
    End If
    Call toGraph
End Sub

Sub toGraph()
    sqlLine = "": sqlitem = ""
    If CboLine.ListIndex <> 0 Then sqlLine = "And PO_No = '" & CboLine & "' "
    If CboItem.ListIndex <> 0 Then sqlitem = "And ChildItem_Code = '" & CboItem & "' "
    
    sql = "Select FactoryCD = R.Supplier_Code, S.ChildItem_Code, I.Item_Name " & _
            "from Part_Supply S, Part_Receipt R, Item_Master I " & _
            "where S.DO_No = Convert(Char,R.Seq_No) And S.ChildItem_Code = I.Item_Code " & _
                "And R.ProductionResult_Cls = '1' And S.Supply_Cls = 'S' and isnull(rtrim(S.DO_No),'') <> '' " & _
                "And S.MaterialConsump_Cls is not null " & _
                "And year(S.ChildSupply_Date) = '" & Year(Period) & "' and month(S.ChildSupply_Date) = '" & Month(Period) & "' " & _
                "And R.Supplier_Code = '" & cboFactory & "' " & sqlLine & sqlitem & _
            "Group By R.Supplier_Code, S.ChildItem_Code, I.Item_Name " & _
        "Order by S.ChildItem_Code"
    If rsRecord.State = adStateOpen Then rsRecord.Close
    rsRecord.CursorLocation = adUseClient
    rsRecord.Open sql, Db, adOpenStatic, adLockOptimistic
    
    If rsRecord.EOF Then LblErrMsg = DisplayMsg(4006): Exit Sub
    Call graphDetail
    If rsRecord.State = adStateOpen Then rsRecord.Close
End Sub

Sub graphDetail(Optional stShow As Byte)
Dim rsGraph As New ADODB.Recordset
Dim dist As Double
    
    sqlItem2 = "And ChildItem_Code = '" & rsRecord!childitem_code & "' "
    frmEffGraph8.LblRecord = "Record " & rsRecord.AbsolutePosition & " Of " & rsRecord.RecordCount
    
    sql = "Select dt.* " & _
        "From " & _
            "( " & _
                "Select FactoryCD = R.Supplier_Code, S.ChildItem_Code, I.Item_Name, " & _
                    "S.MaterialConsump_Cls, Mcc.Description, " & _
                    "qty = IsNull(Sum(s.ChildRequirement_Qty), 0) " & _
                "from Part_Supply S inner join Part_Receipt R on S.DO_No = Convert(Char,R.Seq_No)  " & _
                " Inner Join Item_Master I S.ChildItem_Code = I.Item_Code " & _
                " Left Join MaterialConsump_Cls Mcc on S.MaterialConsump_Cls = Mcc.MaterialConsump_Cls " & _
                "where 'A'='A' " & _
                    "And R.ProductionResult_Cls = '1' And S.Supply_Cls = 'S' and isnull(rtrim(S.DO_No),'') <> '' " & _
                    "And S.MaterialConsump_Cls in " & _
                        "(Select distinct top 7 MaterialConsump_Cls From Part_Supply S2, Part_Receipt R2 " & _
                        "where S2.DO_No = Convert(Char,R2.Seq_No) And R2.ProductionResult_Cls = 1 " & _
                            "And year(S2.ChildSupply_Date) = '" & Year(Period) & "' and month(S2.ChildSupply_Date) = '" & Month(Period) & "' " & _
                            "And S2.Supply_Cls = 'S' and isnull(rtrim(S2.DO_No),'') <> '' " & _
                            "And S2.ChildItem_Code = S.ChildItem_Code And S2.MaterialConsump_Cls is not null " & _
                            "And R2.Supplier_Code = '" & cboFactory & "' " & sqlLine & sqlItem2 & _
                        ") " & _
                    "And year(S.ChildSupply_Date) = '" & Year(Period) & "' and month(S.ChildSupply_Date) = '" & Month(Period) & "' " & _
                    "And R.Supplier_Code = '" & cboFactory & "' " & sqlLine & sqlItem2 & _
                "Group By R.Supplier_Code, S.ChildItem_Code, I.Item_Name, S.MaterialConsump_Cls,  Mcc.Description "
    sql = sql & _
                "Union " & _
                "Select FactoryCD = R.Supplier_Code, S.ChildItem_Code, I.Item_Name, " & _
                    "MaterialConsump_Cls = 'zz', Description = 'Others', " & _
                    "qty = IsNull(Sum(s.ChildRequirement_Qty), 0) " & _
                "from Part_Supply S, Part_Receipt R, Item_Master I " & _
                "where S.DO_No = Convert(Char,R.Seq_No) And S.ChildItem_Code = I.Item_Code " & _
                    "And R.ProductionResult_Cls = '1' And S.Supply_Cls = 'S' and isnull(rtrim(S.DO_No),'') <> '' " & _
                    "And S.MaterialConsump_Cls not in " & _
                        "(Select distinct top 7 MaterialConsump_Cls From Part_Supply S2, Part_Receipt R2 " & _
                        "where S2.DO_No = Convert(Char,R2.Seq_No) And R2.ProductionResult_Cls = 1 " & _
                            "And year(S2.ChildSupply_Date) = '" & Year(Period) & "' and month(S2.ChildSupply_Date) = '" & Month(Period) & "' " & _
                            "And S2.Supply_Cls = 'S' and isnull(rtrim(S2.DO_No),'') <> '' " & _
                            "And S2.ChildItem_Code = S.ChildItem_Code And S2.MaterialConsump_Cls is not null " & _
                            "And R2.Supplier_Code = '" & cboFactory & "' " & sqlLine & sqlItem2 & _
                        ") " & _
                    "And S.MaterialConsump_Cls is not null " & _
                    "And year(S.ChildSupply_Date) = '" & Year(Period) & "' and month(S.ChildSupply_Date) = '" & Month(Period) & "' " & _
                    "And R.Supplier_Code = '" & cboFactory & "' " & sqlLine & sqlItem2 & _
                "Group By R.Supplier_Code, S.ChildItem_Code, I.Item_Name " & _
            ")dt "
    
    If CboItem.ListIndex <> 0 Then sql = sql & "Where dt.ChildItem_Code = '" & CboItem & "'"
    sql = sql & "Order By dt.FactoryCD, (Case When dt.MaterialConsump_Cls = 'zz' Then -1 Else Qty End) desc"
    
    rsGraph.CursorLocation = adUseClient
    rsGraph.Open sql, Db, adOpenDynamic, adLockOptimistic
    If rsGraph.EOF Then LblErrMsg = DisplayMsg(4006): Exit Sub
    
    With frmEffGraph8
        .lblJudul(0) = lblJudul
        .lblJudul(1) = "Machine No : " & CboLine & " (" & CboLine.Column(1) & ")"
        .lblJudul(3) = "Period : " & Format(Period, "MMM yyyy")
        .visibleSign = True
        .jmlBar = rsGraph.RecordCount
        
        i = 0
        Do While Not rsGraph.EOF
            .lblJudul(2) = "Material : " & Trim(rsGraph!childitem_code) & " (" & Trim(rsGraph!item_name) & ")"
            .lblBarVal(i) = Format(rsGraph!Qty, gs_formatQty)
            If i = 0 Then
                .lblBarVal(i).Tag = rsGraph!Qty
            Else
                .lblBarVal(i).Tag = CDbl(.lblBarVal(i - 1).Tag) + rsGraph!Qty
            End If
            .lblX(i * 4).Tag = Trim(rsGraph!Description)
            .MaxQty = .lblBarVal(i).Tag
            i = i + 1
            rsGraph.MoveNext
        Loop

        .lblN = "N = " & Format(.MaxQty, gs_formatNSample)
        .viewGraph
        If stShow = 0 Then .Show 1
    End With
    
    If rsGraph.State = adStateOpen Then rsGraph.Close
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
    If ErrMsg = "" Then
        Unload Me
    Else
        LblErrMsg.Caption = ErrMsg
    End If
End Sub

Private Sub CmdSubMenu_Click()
    frmMainMenu.Show
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = 1
End Sub

