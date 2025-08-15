VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FRM_CLS 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Classification Master"
   ClientHeight    =   10500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14910
   Icon            =   "Frm_CLS.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10500
   ScaleWidth      =   14910
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUnit 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8040
      MaxLength       =   25
      TabIndex        =   15
      Top             =   8520
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   300
      TabIndex        =   11
      Top             =   9030
      Width           =   14310
      Begin VB.Label LblErrMsg 
         Alignment       =   2  'Center
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
         TabIndex        =   12
         Top             =   195
         Width           =   14085
      End
   End
   Begin VB.TextBox TxtCode 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   480
      MaxLength       =   2
      TabIndex        =   0
      Top             =   8520
      Width           =   975
   End
   Begin VB.TextBox TxtDescription 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1605
      MaxLength       =   25
      TabIndex        =   1
      Top             =   8520
      Width           =   6270
   End
   Begin VB.CommandButton CmdSubmit 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Submit"
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
      Left            =   13485
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9735
      Width           =   1125
   End
   Begin VB.CommandButton CmdSubMenu 
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
      Left            =   300
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9720
      Width           =   1125
   End
   Begin VB.CommandButton CmdSearch 
      BackColor       =   &H0080FFFF&
      Caption         =   "Sea&rch"
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
      Left            =   11045
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9735
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.CommandButton CmdClear 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Clear"
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
      Left            =   12265
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9735
      Width           =   1125
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   5175
      Left            =   480
      TabIndex        =   10
      Top             =   2535
      Width           =   13575
      _cx             =   23945
      _cy             =   9128
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483640
      BackColorFixed  =   10932991
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483624
      BackColorAlternate=   -2147483624
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   275
      RowHeightMax    =   275
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   1
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   6720
      Left            =   300
      TabIndex        =   9
      Top             =   1155
      Width           =   14205
      _ExtentX        =   25056
      _ExtentY        =   11853
      _Version        =   393216
      Tabs            =   29
      Tab             =   27
      TabsPerRow      =   10
      TabHeight       =   706
      BackColor       =   16637923
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Control"
      TabPicture(0)   =   "Frm_CLS.frx":0E42
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Currency"
      TabPicture(1)   =   "Frm_CLS.frx":0E5E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Unit"
      TabPicture(2)   =   "Frm_CLS.frx":0E7A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Group"
      TabPicture(3)   =   "Frm_CLS.frx":0E96
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "Material"
      TabPicture(4)   =   "Frm_CLS.frx":0EB2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "Sheet Coil"
      TabPicture(5)   =   "Frm_CLS.frx":0ECE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      TabCaption(6)   =   "Drawing Material"
      TabPicture(6)   =   "Frm_CLS.frx":0EEA
      Tab(6).ControlEnabled=   0   'False
      Tab(6).ControlCount=   0
      TabCaption(7)   =   "Surface Treatment"
      TabPicture(7)   =   "Frm_CLS.frx":0F06
      Tab(7).ControlEnabled=   0   'False
      Tab(7).ControlCount=   0
      TabCaption(8)   =   "Heat Treatment"
      TabPicture(8)   =   "Frm_CLS.frx":0F22
      Tab(8).ControlEnabled=   0   'False
      Tab(8).ControlCount=   0
      TabCaption(9)   =   "Mat. Consump"
      TabPicture(9)   =   "Frm_CLS.frx":0F3E
      Tab(9).ControlEnabled=   0   'False
      Tab(9).ControlCount=   0
      TabCaption(10)  =   "Loss Time"
      TabPicture(10)  =   "Frm_CLS.frx":0F5A
      Tab(10).ControlEnabled=   0   'False
      Tab(10).ControlCount=   0
      TabCaption(11)  =   "Process"
      TabPicture(11)  =   "Frm_CLS.frx":0F76
      Tab(11).ControlEnabled=   0   'False
      Tab(11).ControlCount=   0
      TabCaption(12)  =   "Reason"
      TabPicture(12)  =   "Frm_CLS.frx":0F92
      Tab(12).ControlEnabled=   0   'False
      Tab(12).ControlCount=   0
      TabCaption(13)  =   "Payment Term"
      TabPicture(13)  =   "Frm_CLS.frx":0FAE
      Tab(13).ControlEnabled=   0   'False
      Tab(13).ControlCount=   0
      TabCaption(14)  =   "Price Condition"
      TabPicture(14)  =   "Frm_CLS.frx":0FCA
      Tab(14).ControlEnabled=   0   'False
      Tab(14).ControlCount=   0
      TabCaption(15)  =   "Person In Charge"
      TabPicture(15)  =   "Frm_CLS.frx":0FE6
      Tab(15).ControlEnabled=   0   'False
      Tab(15).ControlCount=   0
      TabCaption(16)  =   "Transportation"
      TabPicture(16)  =   "Frm_CLS.frx":1002
      Tab(16).ControlEnabled=   0   'False
      Tab(16).ControlCount=   0
      TabCaption(17)  =   "Packing Style"
      TabPicture(17)  =   "Frm_CLS.frx":101E
      Tab(17).ControlEnabled=   0   'False
      Tab(17).ControlCount=   0
      TabCaption(18)  =   "Insurance"
      TabPicture(18)  =   "Frm_CLS.frx":103A
      Tab(18).ControlEnabled=   0   'False
      Tab(18).ControlCount=   0
      TabCaption(19)  =   "Region"
      TabPicture(19)  =   "Frm_CLS.frx":1056
      Tab(19).ControlEnabled=   0   'False
      Tab(19).ControlCount=   0
      TabCaption(20)  =   "PO Class"
      TabPicture(20)  =   "Frm_CLS.frx":1072
      Tab(20).ControlEnabled=   0   'False
      Tab(20).ControlCount=   0
      TabCaption(21)  =   "Department"
      TabPicture(21)  =   "Frm_CLS.frx":108E
      Tab(21).ControlEnabled=   0   'False
      Tab(21).ControlCount=   0
      TabCaption(22)  =   "Section"
      TabPicture(22)  =   "Frm_CLS.frx":10AA
      Tab(22).ControlEnabled=   0   'False
      Tab(22).ControlCount=   0
      TabCaption(23)  =   "Transport"
      TabPicture(23)  =   "Frm_CLS.frx":10C6
      Tab(23).ControlEnabled=   0   'False
      Tab(23).ControlCount=   0
      TabCaption(24)  =   "Remarks"
      TabPicture(24)  =   "Frm_CLS.frx":10E2
      Tab(24).ControlEnabled=   0   'False
      Tab(24).ControlCount=   0
      TabCaption(25)  =   "Model"
      TabPicture(25)  =   "Frm_CLS.frx":10FE
      Tab(25).ControlEnabled=   0   'False
      Tab(25).ControlCount=   0
      TabCaption(26)  =   "Clasification Part"
      TabPicture(26)  =   "Frm_CLS.frx":111A
      Tab(26).ControlEnabled=   0   'False
      Tab(26).ControlCount=   0
      TabCaption(27)  =   "Destination"
      TabPicture(27)  =   "Frm_CLS.frx":1136
      Tab(27).ControlEnabled=   -1  'True
      Tab(27).ControlCount=   0
      TabCaption(28)  =   "Color"
      TabPicture(28)  =   "Frm_CLS.frx":1152
      Tab(28).ControlEnabled=   0   'False
      Tab(28).ControlCount=   0
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   12720
      TabIndex        =   13
      TabStop         =   0   'False
      Tag             =   "FTTF*/"
      Top             =   120
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   714
   End
   Begin VB.Label LblUnit 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Convertion"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8040
      TabIndex        =   14
      Top             =   8130
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00A6D2FF&
      Height          =   555
      Left            =   300
      Top             =   8400
      Width           =   9900
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1605
      TabIndex        =   8
      Top             =   8130
      Width           =   4335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
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
      Left            =   480
      TabIndex        =   7
      Top             =   8130
      Width           =   450
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   300
      Top             =   8040
      Width           =   9900
   End
   Begin VB.Label LblHeader 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Classification Master"
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
      Left            =   300
      TabIndex        =   6
      Top             =   360
      Width           =   14205
   End
End
Attribute VB_Name = "FRM_CLS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strTable As String
Dim strSQL As String
Dim UnitFlag As Boolean

Dim bytColSelect As String
Dim bytColCode As String
Dim bytColDescription As String
Dim bytColUnitConver As String

Private Sub SaveData()
    Dim adoLock As New ADODB.Connection
    Dim adoRs As New ADODB.Recordset
    
    Dim intRow As Integer
    Dim strSqlDel As String
    
    On Error GoTo errHandler
    Me.MousePointer = vbHourglass
    
    adoLock.ConnectionString = Db.ConnectionString
    adoLock.Open
    adoLock.BeginTrans
    
    If grid.FindRow("D", , bytColSelect) <= 0 Then
        If txtCode.Text = "" Then
            LblErrMsg.Caption = DisplayMsg("0001") & " Code !"
            txtCode.SetFocus
            GoTo ErrExit
        End If
        
        If txtDescription.Text = "" Then
            LblErrMsg.Caption = DisplayMsg("0001") & " Description !"
            txtDescription.SetFocus
            GoTo ErrExit
        End If
        
        If strTable = "Unit_cls" Then
            adoRs.Open strSQL & " where " & strTable & " = '" & txtCode.Text & "' ", adoLock, adOpenDynamic, adLockOptimistic, adCmdText
            If adoRs.EOF Then
                adoRs.AddNew
                adoRs.Fields(0) = txtCode.Text
                adoRs.Fields(1) = txtDescription.Text
                adoRs.Fields(2) = txtUnit.Text
                adoRs.Fields(3) = Now
                adoRs.Fields(4) = userLogin
                adoRs.Fields(5) = Now
                adoRs.update
                LblErrMsg.Caption = DisplayMsg(1000)
            Else
                If txtCode.Enabled Then
                    If MsgBox("Record already exist! Do you want to update?", vbQuestion + vbYesNo) = vbNo Then GoTo ErrExit
                End If
                adoRs.Fields(1) = txtDescription.Text
                adoRs.Fields(2) = txtUnit.Text
                adoRs.Fields(3) = Now
                adoRs.Fields(4) = userLogin
                adoRs.update
                LblErrMsg.Caption = DisplayMsg(1101)
            End If
            adoRs.Close
        Else
            adoRs.Open strSQL & " where " & strTable & " = '" & txtCode.Text & "' ", adoLock, adOpenDynamic, adLockOptimistic, adCmdText
            If adoRs.EOF Then
                adoRs.AddNew
                adoRs.Fields(0) = txtCode.Text
                adoRs.Fields(1) = txtDescription.Text
                adoRs.Fields(2) = Now
                adoRs.Fields(3) = userLogin
                adoRs.Fields(4) = Now
                adoRs.update
                LblErrMsg.Caption = DisplayMsg(1000)
            Else
                If txtCode.Enabled Then
                    If MsgBox("Record already exist! Do you want to update?", vbQuestion + vbYesNo) = vbNo Then GoTo ErrExit
                End If
                adoRs.Fields(1) = txtDescription.Text
                adoRs.Fields(2) = Now
                adoRs.Fields(3) = userLogin
                adoRs.update
                LblErrMsg.Caption = DisplayMsg(1101)
            End If
            adoRs.Close
        End If
    Else
        If MsgBox("Are you sure want to delete?", vbQuestion + vbYesNo) = vbYes Then
            For intRow = 1 To grid.Rows - 1
                If grid.TextMatrix(intRow, bytColSelect) = "D" Then
                    strSqlDel = "delete from " & strTable & " where " & strTable & " = '" & grid.TextMatrix(intRow, bytColCode) & "'"
                    adoLock.Execute strSqlDel
                End If
            Next
            LblErrMsg.Caption = DisplayMsg(1201)
        Else
            GoTo ErrExit
        End If
    End If
    
    adoLock.CommitTrans
    adoLock.Close
    
    ShowData
    txtCode.Enabled = True
    txtCode.Text = ""
    txtDescription.Text = ""
    txtUnit.Text = ""
    txtCode.SetFocus
    
ErrExit:
    Me.MousePointer = vbDefault
    Set adoRs = Nothing
    Set adoLock = Nothing
    Exit Sub
errHandler:
    LblErrMsg.Caption = "[" & err.number & "] " & err.Description
    err.clear
    Resume ErrExit
End Sub

Private Sub SetGridHeader()
    bytColSelect = 0
    bytColCode = 1
    bytColDescription = 2
        
    With grid
        .Redraw = flexRDNone
        .clear
        .ColS = 3
        .FixedCols = 0
        .Rows = 1
        .FixedRows = 1
        
        .TextMatrix(0, bytColCode) = "Code"
        .TextMatrix(0, bytColDescription) = "Description"
        
        .ColAlignment(bytColSelect) = flexAlignCenterCenter
        .ColAlignment(bytColCode) = flexAlignLeftCenter
        .ColAlignment(bytColDescription) = flexAlignLeftCenter
        
        .ColWidth(bytColSelect) = 300
        .ColWidth(bytColCode) = 1000
        .ColWidth(bytColDescription) = 6000
                
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub ShowData()
    Dim adoRs As New ADODB.Recordset
    
   On Error GoTo errHandler
    Me.MousePointer = vbHourglass
    
    SetGridHeader
    
    Shape1.Width = 7740
    Shape2.Width = 7740
    txtUnit.Visible = False
    lblUnit.Visible = False
    
    strTable = Replace(SSTab.Caption, " ", "") & "_cls"
    If strTable = "Group_cls" Then
        txtCode.MaxLength = 8
    Else
        txtCode.MaxLength = 2
    End If
    
    txtDescription.MaxLength = 25
    grid.ColHidden(bytColSelect) = False
    
    Select Case SSTab.Caption
    Case "Control"
        grid.ColHidden(bytColSelect) = True
    Case "Currency"
        strTable = "Curr_Cls"
        grid.ColHidden(bytColSelect) = True
    Case "Unit"
        UnitFlag = 1
        grid.ColHidden(bytColSelect) = True
        Shape1.Width = 9900
        Shape2.Width = 9900
        txtUnit.Visible = True
        lblUnit.Visible = True
    Case "Mat. Consump"
        strTable = "MaterialConsump_Cls"
    Case "Loss Time"
        strTable = "WorkingLossTime_Cls"
    Case "PO Class"
        strTable = "PO_Cls"
    Case "Remarks"
        strTable = "Remarks_Cls"
    Case "Model"
        strTable = "Model_Cls"
    Case "Clasification Part"
        strTable = "ClasificationPart_Cls"
    End Select
    
    If UnitFlag = True Then
        SetGridHeaderUnit
        strSQL = "select * from " & strTable
        adoRs.Open strSQL, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
        grid.Redraw = flexRDNone
        While adoRs.EOF = False

            grid.AddItem ""
            grid.TextMatrix(grid.Rows - 1, bytColCode) = Trim(adoRs.Fields(0) & "")
            grid.TextMatrix(grid.Rows - 1, bytColDescription) = Trim(adoRs.Fields(1) & "")
            grid.TextMatrix(grid.Rows - 1, bytColUnitConver) = Trim(adoRs.Fields(2) & "")
            
            grid.Cell(flexcpBackColor, grid.Rows - 1, bytColSelect) = vbWhite
            
            adoRs.MoveNext
        Wend
        grid.Redraw = flexRDDirect
        adoRs.Close
        UnitFlag = False
        
    Else
        strSQL = "select * from " & strTable
        adoRs.Open strSQL, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
        grid.Redraw = flexRDNone
        While adoRs.EOF = False
            grid.AddItem ""
            grid.TextMatrix(grid.Rows - 1, bytColCode) = Trim(adoRs.Fields(0) & "")
            grid.TextMatrix(grid.Rows - 1, bytColDescription) = Trim(adoRs.Fields(1) & "")
            
            
            
            grid.Cell(flexcpBackColor, grid.Rows - 1, bytColSelect) = vbWhite
            
            adoRs.MoveNext
        Wend
        grid.Redraw = flexRDDirect
        adoRs.Close
    End If
    
ErrExit:
    Me.MousePointer = vbDefault
    Set adoRs = Nothing
    Exit Sub
errHandler:
    LblErrMsg.Caption = "[" & err.number & "] " & err.Description
    err.clear
    Resume ErrExit
End Sub

Private Sub cmdClear_Click()
    LblErrMsg.Caption = ""
    ShowData
    
    txtCode.Enabled = True
    txtCode.Text = ""
    txtDescription.Text = ""
    txtCode.SetFocus
End Sub

Private Sub CmdSubMenu_Click()
    frmMainMenu.Show
    Unload Me
End Sub

Private Sub CmdSubmit_Click()
    SaveData
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
    If ErrMsg = "" Then
        Unload Me
    Else
        LblErrMsg.Caption = ErrMsg
    End If
End Sub

Private Sub Form_Load()
    If gb_Simulation = True Then Call up_InitSimulation(Me)
    Shape1.Width = 7740
    Shape2.Width = 7740
    
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
    
    HakU = hakUpdate(Me.Name)
    ShowData
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim intRow As Integer
    If Col = bytColSelect Then
        If grid.TextMatrix(Row, Col) = "S" Then
            For intRow = 1 To grid.Rows - 1
                If intRow <> Row Then grid.TextMatrix(intRow, bytColSelect) = ""
            Next
            txtCode.Text = grid.TextMatrix(Row, bytColCode)
            txtDescription.Text = grid.TextMatrix(Row, bytColDescription)
            txtCode.Enabled = False
        ElseIf grid.TextMatrix(Row, Col) = "D" Then
            For intRow = 1 To grid.Rows - 1
                If grid.TextMatrix(intRow, bytColSelect) = "S" Then grid.TextMatrix(intRow, bytColSelect) = ""
            Next
            txtCode.Text = ""
            txtDescription.Text = ""
            txtUnit.Text = ""
            txtCode.Enabled = True
        End If
    End If
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    LblErrMsg.Caption = ""
    If Col = bytColSelect Then
        grid.EditMaxLength = 1
    Else
        Cancel = True
    End If
End Sub

Private Sub grid_KeyPress(KeyAscii As Integer)
    If grid.Col = bytColSelect Then
        If InStr(1, "SD", UCase(Chr(KeyAscii))) = 0 And KeyAscii <> vbKeyReturn Then
            KeyAscii = 0
        Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
    End If
End Sub

Private Sub Grid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If grid.Col = bytColSelect Then
        If InStr(1, "SD", UCase(Chr(KeyAscii))) = 0 And KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyEscape And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack Then
            KeyAscii = 0
        Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
    End If
End Sub

Private Sub SSTab_Click(PreviousTab As Integer)
    LblErrMsg.Caption = ""
    If PreviousTab <> SSTab.Tab Then cmdClear_Click
End Sub

Private Sub SetGridHeaderUnit()
    bytColSelect = 0
    bytColCode = 1
    bytColDescription = 2
    bytColUnitConver = 3
        
    With grid
        .Redraw = flexRDNone
        .clear
        .ColS = 4
        .FixedCols = 0
        .Rows = 1
        .FixedRows = 1
        
        .TextMatrix(0, bytColCode) = "Code"
        .TextMatrix(0, bytColDescription) = "Description"
        .TextMatrix(0, bytColUnitConver) = "Unit Convertion"
        
        .ColAlignment(bytColSelect) = flexAlignCenterCenter
        .ColAlignment(bytColCode) = flexAlignLeftCenter
        .ColAlignment(bytColDescription) = flexAlignLeftCenter
        .ColAlignment(bytColUnitConver) = flexAlignLeftCenter
        
        .ColWidth(bytColSelect) = 300
        .ColWidth(bytColCode) = 1000
        .ColWidth(bytColDescription) = 3000
        .ColWidth(bytColUnitConver) = 3000
                
        .Redraw = flexRDDirect
    End With
End Sub


Private Sub txtUnit_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
