VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmBC40List 
   BackColor       =   &H00FDDFE3&
   Caption         =   "BC 40 List"
   ClientHeight    =   10680
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15120
   Icon            =   "FrmBC40List.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   37644
   ScaleMode       =   0  'User
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdDelete 
      BackColor       =   &H0080FFFF&
      Caption         =   "De&lete"
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
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   25
      Tag             =   "FFTT*/"
      Top             =   10080
      Width           =   1245
   End
   Begin VB.CommandButton cmdSyncronize 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Get BC No."
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
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   23
      Tag             =   "FFTT*/"
      Top             =   10080
      Width           =   1245
   End
   Begin VB.TextBox txtSupplierCode 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1440
      TabIndex        =   22
      Tag             =   "TTFF*/"
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtSuratJalan 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   0
      TabIndex        =   21
      Tag             =   "TTFF*/"
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton CmdDetail 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Detail"
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
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   15
      Tag             =   "FFTT*/"
      Top             =   10080
      Width           =   1125
   End
   Begin VB.CommandButton Cmd_SubMenu 
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   14
      Tag             =   "TFFT*/"
      Top             =   10020
      Width           =   1125
   End
   Begin VB.CommandButton CmdCreate 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Create"
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
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   13
      Tag             =   "FFTT*/"
      Top             =   10080
      Width           =   1125
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   600
      Left            =   120
      TabIndex        =   11
      Tag             =   "TFTT*/"
      Top             =   9300
      Width           =   14640
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
         TabIndex        =   12
         Tag             =   "TTTF*/"
         Top             =   180
         Width           =   14325
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1875
      Left            =   150
      TabIndex        =   0
      Tag             =   "TTTF*/"
      Top             =   1080
      Width           =   14655
      Begin VB.CommandButton CmdSearch 
         BackColor       =   &H0080FFFF&
         Caption         =   "&Search"
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
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   1
         Tag             =   "TTFF*/"
         Top             =   1320
         Width           =   1125
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         Tag             =   "TTFF*/"
         Top             =   360
         Width           =   1530
         _ExtentX        =   2699
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
         CustomFormat    =   "dd MMM yyyy"
         Format          =   157089795
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   315
         Left            =   3780
         TabIndex        =   3
         Tag             =   "TTFF*/"
         Top             =   360
         Width           =   1530
         _ExtentX        =   2699
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
         CustomFormat    =   "dd MMM yyyy"
         Format          =   157089795
         CurrentDate     =   37798
      End
      Begin VB.Label lblTampung 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   8880
         TabIndex        =   20
         Tag             =   "TTFF*/"
         Top             =   720
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lblTradeName 
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
         Left            =   3360
         TabIndex        =   19
         Tag             =   "TTFF*/"
         Top             =   960
         Width           =   60
      End
      Begin MSForms.ComboBox cboTradeCode 
         Height          =   345
         Left            =   1680
         TabIndex        =   10
         Tag             =   "TTFF*/"
         Top             =   840
         Width           =   1530
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2699;609"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         Index           =   3
         Left            =   240
         TabIndex        =   9
         Tag             =   "TTFF*/"
         Top             =   420
         Width           =   405
      End
      Begin VB.Label LblCustomer 
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
         Left            =   3420
         TabIndex        =   8
         Tag             =   "TTFF*/"
         Top             =   420
         Width           =   60
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trade Code"
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
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Tag             =   "TTFF*/"
         Top             =   900
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
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
         Left            =   3360
         TabIndex        =   6
         Tag             =   "TTFF*/"
         Top             =   450
         Width           =   210
      End
      Begin VB.Line Line1 
         X1              =   3360
         X2              =   7620
         Y1              =   1150
         Y2              =   1150
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Interface Cls"
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
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Tag             =   "TTFF*/"
         Top             =   1410
         Width           =   1110
      End
      Begin MSForms.ComboBox cbointeface 
         Height          =   345
         Left            =   1680
         TabIndex        =   4
         Tag             =   "TTFF*/"
         Top             =   1335
         Width           =   1530
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2699;609"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   5910
      Left            =   135
      TabIndex        =   16
      Tag             =   "TTTT*/"
      Top             =   3015
      Width           =   14640
      _cx             =   25823
      _cy             =   10425
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
      GridColor       =   12582912
      GridColorFixed  =   12582912
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
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
      ScrollTrack     =   -1  'True
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
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   12960
      TabIndex        =   24
      TabStop         =   0   'False
      Tag             =   "FTTF*/"
      Top             =   120
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   714
   End
   Begin EZRunnerv3.Anchor Anchor1 
      Left            =   0
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   820
   End
   Begin VB.Label CtrlMenu 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BC 40 List"
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
      Index           =   0
      Left            =   150
      TabIndex        =   18
      Tag             =   "FTTF*/"
      Top             =   360
      Width           =   14610
   End
   Begin MSForms.Label LblRecord 
      Height          =   255
      Left            =   11430
      TabIndex        =   17
      Tag             =   "FFTT*/"
      Top             =   9000
      Width           =   3345
      BackColor       =   16637923
      VariousPropertyBits=   8388627
      Caption         =   "0 Record(s)"
      Size            =   "5900;450"
      FontName        =   "Verdana"
      FontEffects     =   1073741827
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
      ParagraphAlign  =   2
      FontWeight      =   700
   End
End
Attribute VB_Name = "FrmBC40List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nilKosong As Boolean

Const ColCheck As Integer = 0
Const ColTradeCode As Integer = 1
Const ColTradeName As Integer = 2
Const colSuratJalanNo As Integer = 3
Const colReceiptDate As Integer = 4
Const colNoPengajuan As Integer = 5
Const ColBCNO As Integer = 6
Const colBCDate As Integer = 7
Const ColStatus As Integer = 8
Const colcount As Integer = 9

Dim RS As New ADODB.Recordset

Dim MySQLCon As New ADODB.Connection

Public SuratJalanNo As String
Public SupplierCode As String
Public NoPengajuan As String

Private Sub cboTradeCode_Change()
LblErrMsg = ""

    If cboTradeCode.ListIndex <> -1 Then
        lblTradeName.Caption = cboTradeCode.Column(1)
        up_GridHeader
    Else
        lblTradeName.Caption = ""
        cboTradeCode.SetFocus
        Exit Sub
    End If
End Sub

Sub Kosong()

    DTPFrom = Format(Year(Now) & "-" & Format(Month(Now), "#0") & "-01", "dd MMM yyyy")
    DTPTo = Format(Now, "dd MMM yyyy")
    
    With cbointeface
        .clear
        .AddItem "ALL"
        .AddItem "Yes"
        .AddItem "No"
        
        .ListIndex = 0
    End With

End Sub

Private Sub CmdCreate_Click()
Dim li_Row As Integer
    Dim NoBC As String
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset
    
    With Grid
        li_Row = 0
        For i = 1 To .Rows - 1
            If .Cell(flexcpChecked, i, ColCheck) = flexChecked Then
                li_Row = i
                Exit For
            End If
        Next i
        
        If li_Row = 0 Then
            LblErrMsg = DisplayMsg(8011)
            Exit Sub
        Else
            SuratJalanNo = Trim(.TextMatrix(li_Row, colSuratJalanNo))
            SupplierCode = Trim(.TextMatrix(li_Row, ColTradeCode))
            NoPengajuan = Trim(Replace(.TextMatrix(li_Row, colNoPengajuan), "-", ""))
            NoBC = Trim(.TextMatrix(li_Row, ColBCNO))
        End If
    End With
    
    If NoBC <> "" Then
        LblErrMsg = "[0000] BC No. for this data is already exists !"
        Exit Sub
    End If
    
    If NoPengajuan <> "" Then
        LblErrMsg = "[0000] NO AJU for this data is already created !"
        Exit Sub
    End If
    
    Me.MousePointer = vbHourglass
    
    Dim Aa As String
    Aa = MsgBox("Are you sure want to create new NO AJU ?", vbYesNo + vbQuestion + vbDefaultButton2, "Question")
    
    If Aa = vbYes Then
        
        GetNoPengajuan
        
        Set cmd = New ADODB.Command
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 0
        cmd.ActiveConnection = Db
        cmd.CommandText = "sp_BC40ListGenerateNoAju_Sel"
        
        cmd.Parameters.append cmd.CreateParameter("SuratJalanNo", adVarChar, adParamInput, 50, SuratJalanNo)
        cmd.Parameters.append cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, Right(NoPengajuan, 6))
        cmd.Parameters.append cmd.CreateParameter("SupplierCode", adVarChar, adParamInput, 6, SupplierCode)
        
        Set RS = cmd.Execute
        
        Set cmd = New ADODB.Command
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 0
        cmd.ActiveConnection = Db
        cmd.CommandText = "sp_BC40Detail_Sel"
        
        cmd.Parameters.append cmd.CreateParameter("SuratJalanNo", adVarChar, adParamInput, 25, SuratJalanNo)
        cmd.Parameters.append cmd.CreateParameter("SupplierCode", adVarChar, adParamInput, 6, SupplierCode)
        
        Set RS = cmd.Execute
        
        If RS.EOF = False Then
            SuratJalanNo = Trim(RS("SuratJalan_No"))
            NoPengajuan = Trim(RS("No_Pengajuan"))
        End If
        
        FrmBC40Detail.Show
        
        Me.MousePointer = vbDefault
        Me.Hide
    End If
    
    Me.MousePointer = vbDefault
    
End Sub

Private Sub cmdDelete_Click()
    Dim strSQL As String
    Dim li_Row As Integer
    Dim NoBC As String
    Dim StsInterface As Boolean
    
    With Grid
        li_Row = 0
        For i = 1 To .Rows - 1
            If .Cell(flexcpChecked, i, ColCheck) = flexChecked Then
                li_Row = i
                Exit For
            End If
        Next i
        
        If li_Row = 0 Then
            LblErrMsg = DisplayMsg(8011)
            Exit Sub
        Else
            SuratJalanNo = Trim(.TextMatrix(li_Row, colSuratJalanNo))
            NoPengajuan = Trim(Replace(.TextMatrix(li_Row, colNoPengajuan), "-", ""))
            NoBC = Trim(.TextMatrix(li_Row, ColBCNO))
            
            StsInterface = Trim(.TextMatrix(li_Row, ColStatus)) = "Yes"
        End If
    End With
        
    If NoPengajuan = "" Then
        LblErrMsg = "[0000] Please select data with NO AJU !"
        Exit Sub
    End If
    
    If StsInterface Then
        LblErrMsg = "[0000] Data already sent to CEISA !"
        Exit Sub
    End If
    
    Me.MousePointer = vbHourglass
    
    Dim Aa As String
    Aa = MsgBox("Are you sure want to delete this NO AJU ?", vbYesNo + vbQuestion + vbDefaultButton2, "Question")
    
    If Aa = vbYes Then
        
        strSQL = "EXEC sp_BC40Header_Del '" & Replace(NoPengajuan, "-", "") & "'"
        Db.Execute strSQL
        
    End If
    
    up_GridLoad
    
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdDetail_Click()
    up_Detail
End Sub

Private Sub up_Detail()
Dim cek As Integer
    Dim rsCek As New ADODB.Recordset

    Me.MousePointer = vbHourglass

    With Grid
        cek = 0
        For i = 1 To .Rows - 1
            If .Cell(flexcpChecked, i, ColCheck) = flexChecked Then
                cek = i
                Exit For
            End If
        Next i
        
        If cek = 0 Then
            LblErrMsg = DisplayMsg(8011)
            Me.MousePointer = vbDefault
            Exit Sub
        Else
            If Trim(.TextMatrix(cek, colNoPengajuan)) = "" Then
                LblErrMsg = DisplayMsg(8011)
                Me.MousePointer = vbDefault
                Exit Sub
            Else
                SuratJalanNo = Trim(.TextMatrix(cek, colSuratJalanNo))
                SupplierCode = Trim(.TextMatrix(cek, ColTradeCode))
                NoPengajuan = Trim(Replace(.TextMatrix(cek, colNoPengajuan), "-", ""))
            End If
        End If
    End With
        
    Me.MousePointer = vbDefault
    Me.Hide
    
    FrmBC40Detail.txtNoPengajuan = Format(NoPengajuan, gs_formatNoAju)
    FrmBC40Detail.Show
End Sub

Private Sub cmdSearch_Click()
    up_GridLoad
End Sub

Private Sub cmdSyncronize_Click()
    Dim cmd As ADODB.Command
    Dim RS As ADODB.Recordset
    Dim MySQLRS As New ADODB.Recordset
        
    LblErrMsg.Caption = ""
    
    If Grid.Rows = 1 Then
        Exit Sub
    End If
    
    Me.MousePointer = vbHourglass
    
    Dim Aa As String
    Aa = MsgBox("Are you sure want to get BC No. from CEISA ?", vbYesNo + vbQuestion + vbDefaultButton2, "Question")
    
    If Aa = vbYes Then
        KoneksiMysql
        
        Set cmd = New ADODB.Command
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 0
        cmd.ActiveConnection = Db
        cmd.CommandText = "sp_BC40List_Sel"
        
        cmd.Parameters.append cmd.CreateParameter("StartDate", adDBTime, adParamInput, , DTPFrom.Value)
        cmd.Parameters.append cmd.CreateParameter("EndDate", adDBTime, adParamInput, , DTPTo.Value)
        cmd.Parameters.append cmd.CreateParameter("TradeCode", adVarChar, adParamInput, 6, cboTradeCode.Text)
        cmd.Parameters.append cmd.CreateParameter("Status", adVarChar, adParamInput, 3, Null)
            
        Set RS = cmd.Execute
            
        While RS.EOF = False
            sql = "SELECT NOMOR_DAFTAR, TANGGAL_DAFTAR FROM tpbdb.tpb_header WHERE NOMOR_AJU = '" & RS.Fields("No_Pengajuan") & "'"
            
            If MySQLRS.State <> adStateClosed Then MySQLRS.Close
            MySQLRS.Open sql, MySQLCon, adOpenForwardOnly, adLockReadOnly, adCmdText
            
            If MySQLRS.EOF = False Then
                If MySQLRS.Fields("NOMOR_DAFTAR") <> "" Then
                    sql = "UPDATE Bea_Cukai_TPB_Header SET NOMOR_DAFTAR = '" & MySQLRS.Fields("NOMOR_DAFTAR") & "', TANGGAL_DAFTAR = '" & MySQLRS.Fields("TANGGAL_DAFTAR") & "' WHERE NO_PENGAJUAN = '" & RS.Fields("No_Pengajuan") & "'"
                    Db.Execute sql
                    
                    sql = "UPDATE Part_Receipt SET BC40_No = '" & MySQLRS.Fields("NOMOR_DAFTAR") & "', BC40_Date = '" & MySQLRS.Fields("TANGGAL_DAFTAR") & "' WHERE SuratJalan_No = '" & RS.Fields("SuratJalan_No") & "' AND Supplier_Code = '" & RS.Fields("Supplier_Code") & "'"
                    Db.Execute sql
                End If
            End If
            MySQLRS.Close
            
            RS.MoveNext
        Wend
        RS.Close
        
        MySQLCon.Close
        Set MySQLCon = Nothing
        
        up_GridLoad
    End If
    
    Me.MousePointer = vbDefault
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

up_FillComboTrade
up_GridHeader
Kosong

HakU = hakUpdate(Me.Name)
CtrlMenu1.FormName = Me.Name
Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"

    With Anchor1
      .RegString = "AnchorCtrl,Positions," & Me.Name & "0|0"
      .DoInit
    End With

End Sub

Private Sub up_FillComboTrade()
Dim sql As String
Dim RS As New Recordset

    sql = "Select Trade_code, trade_name From Trade_Master Where Epte_Cls = 1 and Country_Cls=0 "
    Set RS = Db.Execute(sql)

    With cboTradeCode
        .clear
        .columnCount = 2
        .ColumnWidths = "50pt;300pt"
        .ListWidth = 350
        .ListRows = 15
        
    cboTradeCode.AddItem ""
    cboTradeCode.List(0, 0) = "ALL"
    cboTradeCode.List(0, 1) = "ALL"
    
        i = 1
        Do While Not RS.EOF
            .AddItem ""
            .List(i, 0) = Trim(RS("Trade_code"))
            .List(i, 1) = IIf(IsNull(RS("trade_name")), " ", Trim(RS("Trade_Name")))
            RS.MoveNext
            i = i + 1
        Loop
        cboTradeCode.ListIndex = 0
    End With
End Sub

Private Sub up_GridHeader()
    With Grid
        .ColS = colcount
        .Rows = 1
        
        .TextMatrix(0, ColCheck) = ""
        .TextMatrix(0, ColTradeCode) = "Trade Code"
        .TextMatrix(0, ColTradeName) = "Trade Name"
        .TextMatrix(0, colSuratJalanNo) = "Surat Jalan No"
        .TextMatrix(0, colReceiptDate) = "Receipt Date"
        .TextMatrix(0, colNoPengajuan) = "No Pengajuan"
        .TextMatrix(0, ColBCNO) = "BC No"
        .TextMatrix(0, colBCDate) = "BC Date"
        .TextMatrix(0, ColStatus) = "Interface"
        
        .ColWidth(ColCheck) = 300
        .ColWidth(ColTradeCode) = 1150
        .ColWidth(ColTradeName) = 3500
        .ColWidth(colSuratJalanNo) = 2100
        .ColWidth(colReceiptDate) = 1300
        .ColWidth(colNoPengajuan) = 3100
        .ColWidth(ColBCNO) = 750
        .ColWidth(colBCDate) = 1350
        .ColWidth(ColStatus) = 900
                
        .Cell(flexcpAlignment, 0, 0, 0, 7) = flexAlignCenterCenter
        .ColAlignment(colSuratJalanNo) = flexAlignLeftCenter
        .ColAlignment(ColStatus) = flexAlignCenterCenter
        
        .FillStyle = flexFillRepeat
        .CellAlignment = flexAlignCenterCenter
     
    End With
End Sub

Private Sub up_GridLoad()
    Dim ls_sql As String
    Dim RS As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim li_Row As Integer
    Dim JmlTran As Integer
    
    LblErrMsg.Caption = ""

    up_GridHeader
    
    Me.MousePointer = vbHourglass
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC40List_Sel"
    
    cmd.Parameters.append cmd.CreateParameter("StartDate", adDBTime, adParamInput, , DTPFrom.Value)
    cmd.Parameters.append cmd.CreateParameter("EndDate", adDBTime, adParamInput, , DTPTo.Value)
    cmd.Parameters.append cmd.CreateParameter("TradeCode", adVarChar, adParamInput, 6, cboTradeCode.Text)
    cmd.Parameters.append cmd.CreateParameter("Status", adVarChar, adParamInput, 3, cbointeface.Text)
    
    Set RS = cmd.Execute
    
    If RS.EOF = False Then
    
    i = 1
    With Grid
        While Not RS.EOF
            .Rows = .Rows + 1
            
            .Cell(flexcpChecked, i, ColCheck) = flexUnchecked
            .Cell(flexcpBackColor, i, ColCheck) = vbWhite
            .TextMatrix(i, ColTradeCode) = Trim(RS("Supplier_Code"))
            .TextMatrix(i, ColTradeName) = Trim(RS("trade_name"))
            .TextMatrix(i, colSuratJalanNo) = Trim(RS("SuratJalan_No"))
            .TextMatrix(i, colNoPengajuan) = Format(RS("No_Pengajuan"), gs_formatNoAju)
            .TextMatrix(i, colReceiptDate) = Format(Trim(RS("Receipt_Date")), "dd MMM yyyy")
            .TextMatrix(i, ColBCNO) = IIf(IsNull(Trim(RS!BC40_No)), "", Trim(RS!BC40_No))
            .TextMatrix(i, colBCDate) = Format(Trim(RS("BC40_Date")), "dd MMM yyyy")
            .TextMatrix(i, ColStatus) = Trim(RS("Interface_Cls"))
                       
            i = i + 1
        RS.MoveNext
        Wend
    End With
    
    
    LblRecord = Format(i - 1, "#,##0") & " Record(s)"
    
    Me.MousePointer = vbDefault
    
    Else
    
    LblErrMsg.Caption = DisplayMsg(13)
    Me.MousePointer = vbDefault
    LblRecord = " 0 Record(s)"
     
    End If
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> ColCheck Then
        Cancel = True
    Else
        For i = 1 To Grid.Rows - 1
            If Grid.Cell(flexcpChecked, i, Col) = flexChecked Then
                Grid.Cell(flexcpChecked, i, 0) = flexChecked
            Else
                Grid.Cell(flexcpChecked, i, 0) = flexUnchecked
            End If
        Next i
    End If
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim cek As Integer
    
    If Col = ColCheck Then
        If Grid.Cell(flexcpChecked, Row, Col) = flexChecked Then
            cek = 1
        Else
            cek = 2
        End If
        
        For i = 1 To Grid.Rows - 1
            Grid.Cell(flexcpChecked, i, 0) = flexUnchecked
        Next i
        
        Grid.Cell(flexcpChecked, Row, Col) = cek
    End If
End Sub

Private Sub Cmd_SubMenu_Click()
    Unload Me
    frmMainMenu.Show
End Sub

Private Sub GetNoPengajuan()
    Dim sql As String
    Dim RS As New ADODB.Recordset
    
    Dim Period As String
    Dim NoDokumen As String
    
    Period = Format(DTPTo, "YYYY")
    
    RS.Open "SELECT No_KPPBC, NoDoc_BC40 FROM Company_Profile", Db, adOpenForwardOnly, adLockReadOnly
    If RS.EOF = False Then
        NoDokumen = Left(Trim(RS.Fields("No_KPPBC")), 4) & "40" & Trim(RS.Fields("NoDoc_BC40"))
    Else
        NoDokumen = ""
    End If
    RS.Close
    
    KoneksiMysql
    
    sql = "SELECT MAX(NOMOR_AJU) NOMOR_AJU FROM tpbdb.tpb_header WHERE NOMOR_AJU LIKE '" & NoDokumen & Period & "%" & "'"
    If RS.State <> adStateClosed Then RS.Close
    RS.Open sql, MySQLCon, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If RS.EOF = False Then
        If IsNull(RS.Fields("NOMOR_AJU")) = True Then
            NoPengajuan = "000000"
        Else
            NoPengajuan = RS.Fields("NOMOR_AJU")
        End If
    Else
        NoPengajuan = "000000"
    End If
    RS.Close
    
    MySQLCon.Close
    Set MySQLCon = Nothing
End Sub

Private Sub KoneksiMysql()
    Dim db_name As String
    Dim db_server As String
    Dim db_port As String
    Dim db_user As String
    Dim db_pass As String
    
    Dim sql As String
    Dim RS As New Recordset
    
    '//error traping
    On Error GoTo buat_koneksi_Error
    
    sql = "SELECT * FROM Connection_Mysql"
    Set RS = Db.Execute(sql)
        
    '/variable localhost
    db_name = Trim(RS("DatabaseName"))
    db_server = Trim(RS("ServerName"))
    db_port = Trim(RS("Port"))
    db_user = Trim(RS("UserId"))
    db_pass = fc_Decrypt(Trim(RS("Password")))
    
    '/buka koneksi my sql
    With MySQLCon
        .ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & db_server & ";DATABASE=" & db_name & ";UID=" & db_user & ";PWD=" & db_pass & ";PORT=" & db_port & ""
        .Open
    End With
            
    On Error GoTo 0
    Exit Sub
    
buat_koneksi_Error:
    MsgBox "[" & err.number & "] " & err.Description, vbCritical, "Error"
End Sub

