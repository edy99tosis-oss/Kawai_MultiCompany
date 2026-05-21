VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmProdScheduleInterface 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Production Schedule Interface"
   ClientHeight    =   10245
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15540
   Icon            =   "frmProdScheduleInterface.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10245
   ScaleWidth      =   15540
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1920
      Top             =   9600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSubmit 
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
      Left            =   14280
      Style           =   1  'Graphical
      TabIndex        =   7
      Tag             =   "FFTT*/"
      Top             =   9720
      Width           =   1125
   End
   Begin VB.CommandButton cmdClear 
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
      Left            =   13080
      Style           =   1  'Graphical
      TabIndex        =   8
      Tag             =   "FFTT*/"
      Top             =   9720
      Width           =   1125
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FDDFE3&
      Height          =   615
      Left            =   120
      TabIndex        =   24
      Tag             =   "TFTT*/"
      Top             =   8880
      Width           =   15240
      Begin VB.Label lblErrMsg 
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
         Height          =   300
         Left            =   120
         TabIndex        =   25
         Tag             =   "TFTF*/"
         Top             =   240
         Width           =   15000
      End
   End
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   9
      Tag             =   "TFFT*/"
      Top             =   9720
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1650
      Left            =   120
      TabIndex        =   11
      Tag             =   "TTTF*/"
      Top             =   1080
      Width           =   15240
      Begin MSComCtl2.DTPicker scheduledate1 
         Height          =   315
         Left            =   10080
         TabIndex        =   4
         Tag             =   "TTFF*/"
         Top             =   720
         Width           =   1545
         _ExtentX        =   2725
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
         Format          =   127533059
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker scheduledate2 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd MMM yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Left            =   12000
         TabIndex        =   5
         Tag             =   "TTFF*/"
         Top             =   720
         Width           =   1545
         _ExtentX        =   2725
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
         Format          =   127533059
         CurrentDate     =   37798
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Interface Status"
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
         Index           =   4
         Left            =   7680
         TabIndex        =   26
         Tag             =   "TTFF*/"
         Top             =   1230
         Width           =   1380
      End
      Begin MSForms.ComboBox cboFactory 
         Height          =   315
         Left            =   1920
         TabIndex        =   22
         Tag             =   "TTFF*/"
         Top             =   280
         Width           =   1305
         VariousPropertyBits=   746604569
         MaxLength       =   10
         DisplayStyle    =   3
         Size            =   "2302;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblcust 
         BackStyle       =   0  'Transparent
         Caption         =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
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
         TabIndex        =   21
         Tag             =   "TTFF*/"
         Top             =   795
         Width           =   3135
      End
      Begin VB.Label LblCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Process Code"
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
         Left            =   240
         TabIndex        =   20
         Tag             =   "TTFF*/"
         Top             =   795
         Width           =   1170
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   3360
         X2              =   6480
         Y1              =   1035
         Y2              =   1035
      End
      Begin MSForms.ComboBox cbocust 
         Height          =   315
         Left            =   1920
         TabIndex        =   1
         Tag             =   "TTFF*/"
         Top             =   735
         Width           =   1305
         VariousPropertyBits=   612386843
         MaxLength       =   10
         DisplayStyle    =   3
         Size            =   "2302;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Line Code"
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
         TabIndex        =   19
         Tag             =   "TTFF*/"
         Top             =   1230
         Width           =   855
      End
      Begin MSForms.ComboBox cbolinecd 
         Height          =   315
         Left            =   1920
         TabIndex        =   2
         Tag             =   "TTFF*/"
         Top             =   1200
         Width           =   1305
         VariousPropertyBits=   746604571
         MaxLength       =   10
         DisplayStyle    =   3
         Size            =   "2302;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lbllinecd 
         BackStyle       =   0  'Transparent
         Caption         =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
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
         TabIndex        =   18
         Tag             =   "TTFF*/"
         Top             =   1210
         Width           =   3255
      End
      Begin VB.Line Line1 
         X1              =   3360
         X2              =   6480
         Y1              =   1470
         Y2              =   1470
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "to"
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
         Left            =   11715
         TabIndex        =   17
         Tag             =   "TTFF*/"
         Top             =   795
         Width           =   165
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Production Schedule Date"
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
         Left            =   7680
         TabIndex        =   16
         Tag             =   "TTFF*/"
         Top             =   795
         Width           =   2205
      End
      Begin VB.Line Line5 
         X1              =   11400
         X2              =   13660
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "a"
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
         Left            =   11400
         TabIndex        =   15
         Tag             =   "TTFF*/"
         Top             =   390
         Width           =   2235
      End
      Begin MSForms.ComboBox CboGroup 
         Height          =   315
         Left            =   10080
         TabIndex        =   3
         Tag             =   "TTFF*/"
         Top             =   285
         Width           =   1125
         VariousPropertyBits=   746604571
         MaxLength       =   10
         DisplayStyle    =   3
         Size            =   "1984;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Group Cls"
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
         Index           =   1
         Left            =   7680
         TabIndex        =   14
         Tag             =   "TTFF*/"
         Top             =   390
         Width           =   855
      End
      Begin VB.Line Line6 
         X1              =   3360
         X2              =   6480
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label lblFactory 
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
         TabIndex        =   13
         Tag             =   "TTFF*/"
         Top             =   390
         Width           =   3255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Factory Code"
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
         TabIndex        =   12
         Tag             =   "TTFF*/"
         Top             =   390
         Width           =   1140
      End
      Begin MSForms.ComboBox cboStatus 
         Height          =   315
         Left            =   10080
         TabIndex        =   6
         Tag             =   "TTFF*/"
         Top             =   1200
         Width           =   1515
         VariousPropertyBits=   746604571
         MaxLength       =   10
         DisplayStyle    =   3
         Size            =   "2672;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   14760
      TabIndex        =   0
      TabStop         =   0   'False
      Tag             =   "FTTF*/"
      Top             =   240
      Width           =   1845
      _extentx        =   3254
      _extenty        =   714
   End
   Begin VSFlex8Ctl.VSFlexGrid grid 
      Height          =   5865
      Left            =   120
      TabIndex        =   23
      TabStop         =   0   'False
      Tag             =   "TTTT*/"
      Top             =   2880
      Width           =   15240
      _cx             =   26882
      _cy             =   10345
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
      HighLight       =   2
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
      RowHeightMax    =   0
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
      ExplorerBar     =   1
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
   Begin EZRunnerv3.Anchor Anchor1 
      Left            =   240
      Top             =   0
      _extentx        =   847
      _extenty        =   820
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Production Schedule Interface"
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
      Left            =   0
      TabIndex        =   10
      Tag             =   "TTTF*/"
      Top             =   480
      Width           =   16680
   End
End
Attribute VB_Name = "frmProdScheduleInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim sqlGrid As String
Dim startDaily As String
Dim dbTransfer As New ADODB.Connection

'Dim ubah As Boolean, ubahBulan As Boolean, ubahBulanQty As Boolean, Status As String
'Dim PONO As String, poSEqNo As String, lblFNo As String, startDaily As String
'Dim lbllotno As String, lblschdate As Date, lblQty As Double, changeQty As Double
'Public dailypanggil As String
'Dim gabung As Boolean, gabungState As Boolean
'Dim OrderEntry As Double, ForecastAwal As Double, TotalDaily As Double
'Dim tOrderEntry As Double, tForecastAwal As Double, tTotalDaily As Double, tUpdateQty As Double
'Dim insertActForecast As Boolean, insertActOrder As Boolean, insertActDaily As Boolean, insertActWIP As Boolean
'Dim tinsertActForecast As Boolean, tinsertActOrder As Boolean, tinsertActDaily As Boolean, tinsertActWIP As Boolean
'
'Dim notFinishGood As Boolean, tnotFinishGood As Boolean
'Dim WIPLimit As Double, tWIPLimit As Double
'Dim tmpParentLotNo As String, ttmpParentLotNo As String

Dim bteColProdCode As Byte
Dim bteColPart As Byte
Dim bteColDesc As Byte
Dim bteColLotNo As Byte
Dim bteColQty As Byte
Dim bteColUnitCls As Byte
Dim bteColUnit As Byte
Dim BteColSerialFrom As Byte
Dim BteColSerialTo As Byte
Dim bteColDate As Byte
Dim bteColRemark As Byte
Dim bteColAuto As Byte
Dim bteColCustCode As Byte
Dim bteColCustName As Byte
Dim bteColPONo As Byte
Dim bteColSeqNo As Byte
Dim bteColApproved As Byte
Dim bteColApprovedDate As Byte
Dim bteColApprovedUser As Byte
Dim bteColVoid As Byte
Dim bteColVoidDate As Byte
Dim bteColVoidUser As Byte

Dim Auto_Cls As Byte
Dim StrStartSerial As String

Private Sub CboCust_Change()
    If cbocust.ListIndex <> -1 Then
        lblcust.Caption = cbocust.Column(1)
        adtocbolinecd
    End If
    If cbocust.Text <> "" And cbolinecd.Text <> "" Then
        Browse
    Else
        Header
    End If
End Sub

Private Sub cboFactory_Change()
    If cboFactory.ListIndex <> -1 Then
        lblFactory.Caption = cboFactory.Column(1)
        adtocboCust
    End If
    
   Header
End Sub

Private Sub cboGroup_Change()
     If CboGroup.ListIndex <> -1 Then
        Label11.Caption = CboGroup.Column(1)
        Else
        Label11.Caption = "All"
    End If
    Browse
End Sub

Private Sub cbolinecd_Change()
    If cbolinecd.ListIndex <> -1 Then
        lbllinecd.Caption = cbolinecd.Column(1)
    End If
    If cbocust.Text <> "" And cbolinecd.Text <> "" Then
        Browse
    Else
        Header
    End If
End Sub

Private Sub cboLotNo_Change()
    If cbocust.Text <> "" And cbolinecd.Text <> "" Then
        Browse
    End If
End Sub

Private Sub cbostatus_Change()
    If cbocust.Text <> "" And cbolinecd.Text <> "" Then
        Browse
    End If
End Sub

Private Sub cmdClear_Click()
    Kosong
End Sub

Private Sub CmdSubmit_Click()

    Dim tanya As VbMsgBoxResult
    Dim sql As String
    Dim statusProses As String
    Dim StatInsertUpdate As Boolean
    Dim i As Long

    On Error GoTo ErrorMesage

    MousePointer = vbHourglass

    ' =====================================================
    ' HAK AKSES
    ' =====================================================
    If HakU = 0 Then
        lblErrMsg = DisplayMsg(3008)
        MousePointer = vbDefault
        Exit Sub
    End If

    ' =====================================================
    ' KONFIRMASI
    ' =====================================================
    tanya = vbYes
     tanya = MsgBox("Do you really want to submit data ?", vbQuestion + vbYesNo, "Confirmation")

    If tanya = vbNo Then
        MousePointer = vbDefault
        Exit Sub
    End If

    ' =====================================================
    ' VALIDASI
    ' =====================================================
    If Trim(cbocust.Text) = "" Then

        lblErrMsg = DisplayMsg(9017) & " Process Code"
        cbocust.SetFocus
        MousePointer = vbDefault
        Exit Sub

    ElseIf Trim(cbolinecd.Text) = "" Then

        lblErrMsg = DisplayMsg(9017) & " Line Code"
        cbolinecd.SetFocus
        MousePointer = vbDefault
        Exit Sub

    ElseIf Trim(CboGroup.Text) = "" Then

        lblErrMsg = DisplayMsg(9017) & " Group Code"
        CboGroup.SetFocus
        MousePointer = vbDefault
        Exit Sub

    End If

    lblErrMsg.Caption = ""

    ' =====================================================
    ' OPEN CONNECTION
    ' =====================================================
    dbTransfer.ConnectionTimeout = 0
    dbTransfer.CommandTimeout = 1800

    dbTransfer.Open Db.ConnectionString
    dbTransfer.BeginTrans

    StatInsertUpdate = False

    ' =====================================================
    ' PROCESS GRID
    ' =====================================================
    With grid

        For i = 1 To .Rows - 1

        statusProses = ""
    
        ' =================================================
        ' VOID
        ' PRIORITAS PERTAMA
        ' =================================================
        If .Cell(flexcpChecked, i, bteColVoid) = flexChecked Then
    
            ' Jika memang belum pernah void
            If Trim(.TextMatrix(i, bteColVoidDate)) = "" Then
                statusProses = "VOID"
            End If
    
        ' =================================================
        ' APPROVED
        ' =================================================
        ElseIf .Cell(flexcpChecked, i, bteColApproved) = flexChecked Then
    
            ' Jika memang belum pernah approve
            If Trim(.TextMatrix(i, bteColApprovedDate)) = "" Then
                statusProses = "APPROVE"
            End If
    
        End If
    
        ' =================================================
        ' JIKA TIDAK ADA PROSES
        ' SKIP
        ' =================================================
        If statusProses = "" Then GoTo NextData
    
        ' =================================================
        ' EXECUTE STORED PROCEDURE
        ' =================================================
        sql = ""
    
        sql = sql & " EXEC sp_ProdScheduleInf_status "
        sql = sql & " '" & Trim(cbocust.Text) & "', "
        sql = sql & " '" & Trim(cbolinecd.Text) & "', "
        sql = sql & " '" & Format(scheduledate1.Value, "yyyy-mm-dd") & "', "
        sql = sql & " '" & Format(scheduledate2.Value, "yyyy-mm-dd") & "', "
        sql = sql & " '" & Trim(.TextMatrix(i, bteColProdCode)) & "', "
        sql = sql & " '" & Trim(.TextMatrix(i, bteColLotNo)) & "', "
        sql = sql & "  " & Val(.TextMatrix(i, bteColQty)) & ", "
        sql = sql & "  '" & statusProses & "', "
        sql = sql & " '" & userLogin & "' "
    
        dbTransfer.Execute sql
    
        StatInsertUpdate = True
    
NextData:
    
    Next i

    End With

    ' =====================================================
    ' COMMIT
    ' =====================================================
    dbTransfer.CommitTrans
    dbTransfer.Close
    
    Call Export_ToCsv

    ' =====================================================
    ' REFRESH GRID
    ' =====================================================
    If StatInsertUpdate = True Then
        Browse
    End If

    lblErrMsg = DisplayMsg(1101)

    MousePointer = vbDefault
    Exit Sub

' =====================================================
' ERROR
' =====================================================
ErrorMesage:

    If dbTransfer.State = 1 Then
        dbTransfer.RollbackTrans
        dbTransfer.Close
    End If

    lblErrMsg.Caption = err.number & " - " & err.Description

    MousePointer = vbDefault

End Sub

Private Sub Export_ToCsv()

    Dim rsCsv As New ADODB.Recordset
    Dim sql As String
    Dim filename As String
    Dim FullPath As String
    Dim FileNo As Integer

    Dim scheduledate As String
    Dim LineCode As String
    Dim ItemCode As String
    Dim SerialNumber As String
    Dim StatusData As String

    On Error GoTo Errhandle

    ' =====================================================
    ' SAVE FILE DIALOG
    ' =====================================================
    CommonDialog1.CancelError = True
    CommonDialog1.filter = "CSV File (*.csv)|*.csv"
    CommonDialog1.DefaultExt = "csv"
    CommonDialog1.filename = "Schedule-" & Format(Now, "yyyymmdd") & ".csv"
    CommonDialog1.ShowSave

    FullPath = CommonDialog1.filename

    If Trim(FullPath) = "" Then Exit Sub

    ' =====================================================
    ' QUERY
    ' =====================================================
   sql = ""

    sql = sql & " EXEC sp_ProdScheduleInterface_ExportCsv "
    sql = sql & " '" & Trim(cbocust.Text) & "', "
    sql = sql & " '" & Trim(cbolinecd.Text) & "', "
    sql = sql & " '" & Format(scheduledate1.Value, "yyyy-mm-dd") & "', "
    sql = sql & " '" & Format(scheduledate2.Value, "yyyy-mm-dd") & "', "
    sql = sql & " '" & UCase(Trim(cboStatus.Text)) & "' "

    rsCsv.Open sql, Db, adOpenForwardOnly, adLockReadOnly

    If rsCsv.EOF Then

        MsgBox "No data to export.", vbExclamation
        rsCsv.Close
        Exit Sub

    End If

    ' =====================================================
    ' CREATE CSV
    ' =====================================================
    FileNo = FreeFile

    Open FullPath For Output As #FileNo

    ' HEADER
    Print #FileNo, _
        "Schedule Date,Line Code,Item Code,Serial Number,Status"

    ' DETAIL
    Do While Not rsCsv.EOF

        scheduledate = rsCsv.Fields("Schedule Date").Value
         LineCode = "=""" & rsCsv.Fields("Line Code").Value & """"
        ItemCode = "=""" & rsCsv.Fields("Item Code").Value & """"
        SerialNumber = "=""" & rsCsv.Fields("Serial Number").Value & """"
        StatusData = rsCsv.Fields("Status").Value

        Print #FileNo, _
            scheduledate & "," & _
            LineCode & "," & _
            ItemCode & "," & _
            SerialNumber & "," & _
            StatusData

        rsCsv.MoveNext

    Loop

    Close #FileNo

    rsCsv.Close

    MsgBox "Export CSV Success", vbInformation

    Exit Sub

Errhandle:

    If rsCsv.State = 1 Then rsCsv.Close

    MsgBox err.number & " - " & err.Description, vbCritical

End Sub

'Private Sub Export_ToCsv()
'
'    Dim rsCsv As New ADODB.Recordset
'    Dim rsFtp As New ADODB.Recordset
'
'    Dim sql As String
'    Dim ftpSql As String
'
'    Dim FullPath As String
'    Dim FileNo As Integer
'
'    Dim scheduledate As String
'    Dim LineCode As String
'    Dim ItemCode As String
'    Dim SerialNumber As String
'    Dim StatusData As String
'
'    Dim ftpHost As String
'    Dim ftpUser As String
'    Dim ftpPass As String
'    Dim ftpFolder As String
'
'    Dim ftpUrl As String
'    Dim fileName As String
'
'    On Error GoTo Errhandle
'
'    ' =====================================================
'    ' AUTO FILE NAME
'    ' =====================================================
'
'    fileName = "Schedule-" & Format(Now, "yyyymmdd_hhnnss") & ".csv"
'
'    FullPath = App.path & "\" & fileName
'
'    ' =====================================================
'    ' QUERY EXPORT
'    ' =====================================================
'
'    sql = ""
'
'    sql = sql & " EXEC sp_ProdScheduleInterface_ExportCsv "
'    sql = sql & " '" & Trim(cbocust.Text) & "', "
'    sql = sql & " '" & Trim(cbolinecd.Text) & "', "
'    sql = sql & " '" & Format(scheduledate1.Value, "yyyy-mm-dd") & "', "
'    sql = sql & " '" & Format(scheduledate2.Value, "yyyy-mm-dd") & "', "
'    sql = sql & " '" & UCase(Trim(cboStatus.Text)) & "' "
'
'    rsCsv.Open sql, Db, adOpenForwardOnly, adLockReadOnly
'
'    If rsCsv.EOF Then
'
'        MsgBox "No data to export.", vbExclamation
'        rsCsv.Close
'        Exit Sub
'
'    End If
'
'    ' =====================================================
'    ' CREATE CSV
'    ' =====================================================
'
'    FileNo = FreeFile
'
'    Open FullPath For Output As #FileNo
'
'    ' HEADER
'    Print #FileNo, _
'        "Schedule Date,Line Code,Item Code,Serial Number,Status"
'
'    ' DETAIL
'    Do While Not rsCsv.EOF
'
'        scheduledate = rsCsv.Fields("Schedule Date").Value
'        LineCode = "=""" & rsCsv.Fields("Line Code").Value & """"
'        ItemCode = "=""" & rsCsv.Fields("Item Code").Value & """"
'        SerialNumber = "=""" & rsCsv.Fields("Serial Number").Value & """"
'        StatusData = rsCsv.Fields("Status").Value
'
'        Print #FileNo, _
'            scheduledate & "," & _
'            LineCode & "," & _
'            ItemCode & "," & _
'            SerialNumber & "," & _
'            StatusData
'
'        rsCsv.MoveNext
'
'    Loop
'
'    Close #FileNo
'    rsCsv.Close
'
'    ' =====================================================
'    ' GET FTP CONFIG
'    ' =====================================================
'
'    ftpSql = "EXEC sp_GetFTPConfig"
'
'    rsFtp.Open ftpSql, Db, adOpenForwardOnly, adLockReadOnly
'
'    If rsFtp.EOF Then
'
'        MsgBox "FTP Config not found", vbCritical
'        Exit Sub
'
'    End If
'
'    ftpHost = Trim(rsFtp.Fields("FTP_HOST").Value)
'    ftpUser = Trim(rsFtp.Fields("FTP_USER").Value)
'    ftpPass = Trim(rsFtp.Fields("FTP_PASS").Value)
'    ftpFolder = Trim(rsFtp.Fields("FTP_FOLDER").Value)
'
'    rsFtp.Close
'
'    ' =====================================================
'    ' FTP UPLOAD
'    ' =====================================================
'
'    ftpUrl = "ftp://" & ftpHost & ftpFolder & "/" & fileName
'
'    Inet1.URL = ftpUrl
'    Inet1.userName = ftpUser
'    Inet1.Password = ftpPass
'
'    Inet1.Execute , "PUT " & FullPath & " " & fileName
'
'    Do While Inet1.StillExecuting
'        DoEvents
'    Loop
'
'    MsgBox "Export CSV & Upload FTP Success", vbInformation
'
'    Exit Sub
'
'Errhandle:
'
'    If rsCsv.State = 1 Then rsCsv.Close
'    If rsFtp.State = 1 Then rsFtp.Close
'
'    MsgBox err.number & " - " & err.Description, vbCritical
'
'End Sub

Private Sub Form_Load()
 
    If gb_Simulation = True Then Call up_InitSimulation(Me)
    
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
    HakU = hakUpdate(Me.Name)
    
    adcboCompany
    adtocboCust
    adtocboGroup
    adtocboStatus
    
    Call Kosong
    
    'GET START DAILY
    Dim Ret As String, NC As Long, TempPWD As String
    Ret = String(255, 0)
    NC = GetPrivateProfileString("StartDaily", "Date", "", Ret, 255, IniFile)
    If NC <> 0 Then Ret = Left$(Ret, NC)
    startDaily = Ret
    
    With Anchor1
      .RegString = "AnchorCtrl,Positions," & Me.Name & "0|0"
      .DoInit
    End With
End Sub

Sub Header()
    Dim i As Long
    
    bteColProdCode = 0
    bteColPart = 1
    bteColDesc = 2
    bteColLotNo = 3
    bteColQty = 4
    bteColUnitCls = 5
    bteColUnit = 6
    BteColSerialFrom = 7
    BteColSerialTo = 8
    bteColDate = 9
    bteColRemark = 10
    bteColAuto = 11
    bteColCustCode = 12
    bteColCustName = 13
    bteColPONo = 14
    bteColSeqNo = 15
    bteColApproved = 16
    bteColApprovedDate = 17
    bteColApprovedUser = 18
    bteColVoid = 19
    bteColVoidDate = 20
    bteColVoidUser = 21
    
    
    With grid
      .clear
      .Rows = 1
      .ColS = 22
      
      .ColWidth(bteColProdCode) = 1400
      .ColWidth(bteColPart) = 1400
      .ColWidth(bteColDesc) = 3000
      .ColWidth(bteColLotNo) = 1000
      .ColWidth(bteColQty) = 1230
      .ColWidth(bteColUnit) = 650
      .ColWidth(BteColSerialFrom) = 1100
      .ColWidth(BteColSerialTo) = 1100
      .ColWidth(bteColDate) = 1450
      .ColWidth(bteColRemark) = 3250
      .ColWidth(bteColAuto) = 1000
      .ColWidth(bteColCustCode) = 1200
      .ColWidth(bteColCustName) = 2800
      .ColWidth(bteColPONo) = 2000
      .ColWidth(bteColApproved) = 1050
      .ColWidth(bteColApprovedDate) = 1450
      .ColWidth(bteColApprovedUser) = 1500
      .ColWidth(bteColVoid) = 1000
      .ColWidth(bteColVoidDate) = 1450
      .ColWidth(bteColVoidUser) = 1500
      
      .TextMatrix(0, bteColPart) = "Part Number"
      .TextMatrix(0, bteColProdCode) = "Product Code"
      .TextMatrix(0, bteColDesc) = "Description"
      .TextMatrix(0, bteColLotNo) = "Lot No"
      .TextMatrix(0, bteColQty) = "Qty"
      .TextMatrix(0, bteColUnitCls) = "UnitCls"
      .TextMatrix(0, bteColUnit) = "Unit"
      .TextMatrix(0, BteColSerialFrom) = "Serial From"
      .TextMatrix(0, BteColSerialTo) = "Serial To"
      .TextMatrix(0, bteColDate) = "Schedule Date"
      .TextMatrix(0, bteColRemark) = "Remark"
      .TextMatrix(0, bteColAuto) = "Auto"
      .TextMatrix(0, bteColCustCode) = "Cust. Code"
      .TextMatrix(0, bteColCustName) = "Cust. Name"
      .TextMatrix(0, bteColPONo) = "PO No."
      .TextMatrix(0, bteColSeqNo) = "SeqNo"
      .TextMatrix(0, bteColApproved) = "Approved"
      .TextMatrix(0, bteColApprovedDate) = "Approved Date"
      .TextMatrix(0, bteColApprovedUser) = "Approved User"
      .TextMatrix(0, bteColVoid) = "Void"
      .TextMatrix(0, bteColVoidDate) = "Void Date"
      .TextMatrix(0, bteColVoidUser) = "Void User"
      
      
      .ColHidden(bteColUnitCls) = True
      .ColHidden(bteColSeqNo) = True
      .ColHidden(bteColCustCode) = True
      .ColHidden(bteColCustName) = True
      .ColHidden(bteColPONo) = True
      
      .ColDataType(bteColDate) = flexDTDate
      
      .Cell(flexcpAlignment, 0, 0, 0, bteColVoidUser) = flexAlignCenterCenter
      .ColAlignment(bteColPart) = flexAlignLeftCenter
      .ColAlignment(bteColProdCode) = flexAlignLeftCenter
      .ColAlignment(bteColDesc) = flexAlignLeftCenter
      .ColAlignment(bteColLotNo) = flexAlignCenterCenter
      .ColAlignment(bteColQty) = flexAlignRightCenter
      .ColAlignment(bteColUnit) = flexAlignCenterCenter
      .ColAlignment(BteColSerialFrom) = flexAlignCenterCenter
      .ColAlignment(BteColSerialTo) = flexAlignCenterCenter
      .ColAlignment(bteColDate) = flexAlignCenterCenter
      .ColAlignment(bteColRemark) = flexAlignLeftCenter
      .ColAlignment(bteColAuto) = flexAlignCenterCenter
      .ColAlignment(bteColCustCode) = flexAlignLeftCenter
      .ColAlignment(bteColCustName) = flexAlignLeftCenter
      .ColAlignment(bteColPONo) = flexAlignLeftCenter
      .ColAlignment(bteColApproved) = flexAlignCenterCenter
      .ColAlignment(bteColVoid) = flexAlignCenterCenter
      
      .EditMaxLength = 1
    End With
 
End Sub

Sub Kosong()
    scheduledate1.Value = Format(Now, "dd MMM yyyy")
    scheduledate2.Value = Format(Now, "dd MMM yyyy")
    lblcust.Caption = ""
    cbocust.Text = ""
    cbolinecd.clear
    cbolinecd.Text = ""
    
    lbllinecd.Caption = ""
    CboGroup.ListIndex = 0
    
    lblErrMsg = ""
    
    Header
End Sub

Sub adcboCompany()
    FillCompanyCombo cboFactory
End Sub

Sub adtocboCust()
Dim sqlcust As String
Dim RsCust As New Recordset
Dim i As Integer

    sqlcust = "EXEC dbo.sp_GetCompanyCode @CompanyCode = '" & cboFactory.Text & "'"

    Set RsCust = Db.Execute(sqlcust)
    
    With cbocust
        .clear
        .columnCount = 2
        .ColumnWidths = "50pt;175pt"
        .ListWidth = 225
        .ListRows = 15
        
        i = 0
        Do While Not RsCust.EOF
            .AddItem
            .List(i, 0) = Trim(RsCust("manufacture_code"))
            .List(i, 1) = IIf(IsNull(RsCust("trade_name")), "", Trim(RsCust("Trade_Name")))
            RsCust.MoveNext
            i = i + 1
        Loop
    End With
End Sub

Sub adtocboGroup()
Dim SqlGroup As String
Dim RsGroup As New Recordset
Dim i As Integer

    SqlGroup = "EXEC dbo.sp_FillComboProdScheduleInf @Type = '2', @Param1 = '', @Param2 = ''"

    Set RsGroup = Db.Execute(SqlGroup)
    
    With CboGroup
        .clear
        .columnCount = 2
        .ColumnWidths = "50pt;125pt"
        .ListWidth = 175
        .ListRows = 5
        
       .AddItem
        .List(0, 0) = "All"
        .List(0, 1) = "All"
        
        i = 1
        Do While Not RsGroup.EOF
            .AddItem
            .List(i, 0) = Trim(RsGroup("Group_Cls"))
            .List(i, 1) = IIf(IsNull(RsGroup("Description")), "", Trim(RsGroup("Description")))
            RsGroup.MoveNext
            i = i + 1
        Loop
    End With
End Sub

Sub adtocboStatus()
Dim SqlStatus As String
Dim RsStatus As New Recordset
Dim i As Integer

    SqlStatus = "EXEC dbo.sp_FillComboProdScheduleInf @Type = '3', @Param1 = '', @Param2 = ''"

    Set RsStatus = Db.Execute(SqlStatus)
    
    With cboStatus
        .clear
        .columnCount = 1
        .ColumnWidths = "82pt"
        .ListWidth = 82
        .ListRows = 3
        
        i = 0
        Do While Not RsStatus.EOF
            .AddItem
            .List(i, 0) = Trim(RsStatus("Interface_Status"))
            RsStatus.MoveNext
            i = i + 1
        Loop
        
        .ListIndex = 0
    End With
End Sub

'Sub adtoCboLotNo()
'Dim SqlStatus As String
'Dim RsStatus As New Recordset
'Dim i As Integer
'
'    SqlStatus = " EXEC dbo.sp_FillComboProdScheduleInf @Type = '4', @Param1 = '" & cbocust.Text & "', @Param2 = '" & cbolinecd.Text & "', " & vbCrLf & _
'                " @Param3 = '" & scheduledate1.Value & "', @Param4 = '" & scheduledate2.Value & "'"
'
'    Set RsStatus = Db.Execute(SqlStatus)
'
'    With cboLotNo
'        .clear
'        .columnCount = 1
'        .ColumnWidths = "82pt"
'        .ListWidth = 82
'        .ListRows = 3
'
'        i = 0
'        Do While Not RsStatus.EOF
'            .AddItem
'            .List(i, 0) = Trim(RsStatus("Lot_No"))
'            RsStatus.MoveNext
'            i = i + 1
'        Loop
'
'        .ListIndex = 0
'    End With
'End Sub

Private Sub CmdSubMenu_Click()
    Unload Me
    frmMainMenu.Show
End Sub

Sub Browse()
    
    Dim i As Long
    Dim strSQL As String
    Dim rsGrid As New ADODB.Recordset

    Header
    Auto_Cls = 0
              
    strSQL = "EXEC dbo.SP_ProdScheduleInf_Browse @FactoryCode = '" & Trim(cbocust.Text) & "', @LineCode = '" & Trim(cbolinecd.Text) & "', " & _
            " @GroupCls = '" & Trim(CboGroup.Text) & "', @DateFrom = '" & Format(scheduledate1.Value, "yyyymmdd") & "', " & _
            " @DateTo =  '" & Format(scheduledate2.Value, "yyyymmdd") & "', @Status = '" & UCase(Trim(cboStatus.Text)) & "'"
    
    If rsGrid.State <> adStateClosed Then rsGrid.Close
    Set rsGrid = Db.Execute(strSQL)
          
    i = 1
    If Not (rsGrid.BOF And rsGrid.EOF) Then
        With grid
            Do While Not rsGrid.EOF
                .Rows = .Rows + 1
                .TextMatrix(i, bteColProdCode) = Trim(rsGrid("Item_Code"))
                .TextMatrix(i, bteColPart) = Trim(rsGrid("MakerItem_Code"))
                .TextMatrix(i, bteColDesc) = IIf(IsNull(rsGrid("item_name")), "", Trim(rsGrid("item_name")))
                
                .TextMatrix(i, bteColLotNo) = IIf(IsNull(rsGrid("lot_no")), "", Trim(rsGrid("lot_no")))
                
                .TextMatrix(i, bteColQty) = IIf(IsNull(rsGrid("Qty")), 0, Trim(rsGrid("Qty")))
                If InStr(1, .TextMatrix(i, bteColQty), ".") > 0 Then
                    .TextMatrix(i, bteColQty) = Format(.TextMatrix(i, bteColQty), gs_formatQty)
                Else
                    .TextMatrix(i, bteColQty) = Format(.TextMatrix(i, bteColQty), gs_formatQty)
                End If
        
                If IsNull(rsGrid("unit_cls")) Then
                  .TextMatrix(i, bteColUnitCls) = ""
                  .TextMatrix(i, bteColUnit) = ""
                Else
                  .TextMatrix(i, bteColUnitCls) = Trim(rsGrid("Unit_cls"))
                  .TextMatrix(i, bteColUnit) = uf_GetUnitDescription(Trim(rsGrid("Unit_Cls")))
                End If
                
                .TextMatrix(i, BteColSerialFrom) = IIf(IsNull(rsGrid("SerialNoFrom")), "", Trim(rsGrid("SerialNoFrom"))) 'Add 20090207
                .TextMatrix(i, BteColSerialTo) = IIf(IsNull(rsGrid("SerialNoTo")), "", Trim(rsGrid("SerialNoTo"))) 'Add 20090207
                
                .TextMatrix(i, bteColDate) = Format(Trim(rsGrid("schedule_date")), "dd MMM yyyy")
                .TextMatrix(i, bteColRemark) = IIf(IsNull(rsGrid("remark")), "", Trim(rsGrid("remark")))
                .TextMatrix(i, bteColSeqNo) = Val(rsGrid("seq_no"))
                
                If Val(rsGrid("auto_cls") & "") = 0 Then .TextMatrix(i, bteColAuto) = "No" Else .TextMatrix(i, bteColAuto) = "Yes"
                
                .TextMatrix(i, bteColCustCode) = rsGrid("cust_code") & ""
                .TextMatrix(i, bteColCustName) = rsGrid("trade_name") & ""
                .TextMatrix(i, bteColPONo) = rsGrid("po_no") & ""
                
                .Cell(flexcpBackColor, i, bteColApproved) = vbWhite
                .ColDataType(bteColApproved) = flexDTBoolean
                If IsNull(rsGrid.Fields("Approved_Date").Value) Then
                    .TextMatrix(i, bteColApproved) = False
                Else
                    .TextMatrix(i, bteColApproved) = True
                End If
                
                 .TextMatrix(i, bteColApprovedDate) = Format(Trim(rsGrid("Approved_Date")), "dd MMM yyyy") & ""
                 .TextMatrix(i, bteColApprovedUser) = rsGrid("Approved_User") & ""
                 
                .Cell(flexcpBackColor, i, bteColVoid) = vbWhite
                .ColDataType(bteColVoid) = flexDTBoolean
                If IsNull(rsGrid.Fields("Void_Date").Value) Then
                    .TextMatrix(i, bteColVoid) = False
                Else
                    .TextMatrix(i, bteColVoid) = True
                End If
                
                 .TextMatrix(i, bteColVoidDate) = Format(Trim(rsGrid("Void_Date")), "dd MMM yyyy") & ""
                 .TextMatrix(i, bteColVoidUser) = rsGrid("Void_User") & ""
                
                rsGrid.MoveNext
                i = i + 1
            Loop
        End With
    End If
    rsGrid.Close
    Set rsGrid = Nothing
End Sub

Private Sub scheduledate1_Change()
    If CDate(scheduledate1) > CDate(scheduledate2) Then
          lblErrMsg.Caption = DisplayMsg(4068)
          Exit Sub
    Else
       lblErrMsg.Caption = ""
    End If
     
    If cbocust.Text <> "" And cbolinecd.Text <> "" Then
        Browse
    End If
End Sub

Private Sub scheduledate2_Change()
    If CDate(scheduledate2) < CDate(scheduledate1) Then
      lblErrMsg.Caption = DisplayMsg(4066)
      Exit Sub
    Else
      lblErrMsg.Caption = ""
    End If
    
    If cbocust.Text <> "" And cbolinecd.Text <> "" Then
        Browse
    End If
End Sub

Sub adtocbolinecd()
Dim sqlLine As String
Dim RsLine As New Recordset
Dim i As Integer

    sqlLine = "EXEC dbo.sp_FillComboProdScheduleInf @Type = '1', @Param1 = '" & cbocust.Text & "', @Param2 = ''"
    Set RsLine = Db.Execute(sqlLine)
    
    With cbolinecd
        .clear
        .columnCount = 2
        .ColumnWidths = "50pt;175pt"
        .ListWidth = 225
        .ListRows = 15
        
        lbllinecd.Caption = ""
        i = 0
        Do While Not RsLine.EOF
            .AddItem
            .List(i, 0) = Trim(RsLine("Line_code"))
            .List(i, 1) = Trim(RsLine("Line_Name"))
            RsLine.MoveNext
            i = i + 1
        Loop
    End With
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    With grid

        ' =========================================
        ' APPROVED
        ' =========================================
        If Col = bteColApproved Then

            ' Jika sudah submit dan checkbox sudah checked
            ' maka tidak boleh di-uncheck lagi
            If Trim(.TextMatrix(Row, bteColApprovedDate)) <> "" Then

                If .Cell(flexcpChecked, Row, bteColApproved) = flexChecked Then

                    MsgBox "Data already submitted. Please void first.", vbExclamation
                    Cancel = True
                    Exit Sub

                End If

            End If

        End If


        ' =========================================
        ' VOID
        ' =========================================
        If Col = bteColVoid Then

            ' Belum submit tidak boleh void
            If Trim(.TextMatrix(Row, bteColApprovedDate)) = "" Then

                MsgBox "Data not submitted yet.", vbExclamation
                Cancel = True
                Exit Sub

            End If

        End If


        ' =========================================
        ' READONLY COLUMN
        ' =========================================
        If Col <> bteColApproved And Col <> bteColVoid Then
            Cancel = True
        End If

    End With

End Sub



