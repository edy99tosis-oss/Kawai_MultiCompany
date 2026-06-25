VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmProdScheduleInterface 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Production Schedule Interface"
   ClientHeight    =   10245
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14940
   Icon            =   "frmProdScheduleInterface.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10245
   ScaleWidth      =   14940
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   4080
      Top             =   9600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1920
      Top             =   9600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Left            =   12360
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
      Width           =   14640
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
         Width           =   14520
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
      Width           =   14640
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
         Enabled         =   0   'False
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
         Format          =   129761283
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
         Enabled         =   0   'False
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
         Format          =   129761283
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
         Top             =   720
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
      Left            =   12960
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
      Width           =   14640
      _cx             =   25823
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
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   7
      Tag             =   "FFTT*/"
      Top             =   9720
      Width           =   1125
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
      Left            =   120
      TabIndex        =   10
      Tag             =   "TTTF*/"
      Top             =   600
      Width           =   14640
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


Dim bteColDate As Byte
Dim bteColProdCode As Byte
Dim bteColPart As Byte
Dim bteColDesc As Byte
Dim bteColLotNo As Byte
Dim bteColQty As Byte
Dim bteColUnitCls As Byte
Dim bteColUnit As Byte
Dim BteColSerialFrom As Byte
Dim BteColSerialTo As Byte
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

Private Type tApproveData
    ItemCode As String
    lotno As String
    Qty As Double
End Type

Private arrApprove() As tApproveData
Private ApproveCount As Long

Private Sub CboCust_Change()

lblErrMsg.Caption = ""

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

lblErrMsg.Caption = ""

    If CboGroup.ListIndex <> -1 Then
        Label11.Caption = CboGroup.Column(1)
        Else
        Label11.Caption = "All"
    End If
    Browse
End Sub

Private Sub cbolinecd_Change()

lblErrMsg.Caption = ""

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

lblErrMsg.Caption = ""

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
    ' CHECK REVISE
    ' =====================================================
    Dim blnRevisi As Boolean
    

    With grid
    
        For i = 1 To .Rows - 1
    
            If .Cell(flexcpChecked, i, bteColApproved) = flexChecked Then
    
                If Trim(.TextMatrix(i, bteColApprovedDate)) <> "" Then
    
                    blnRevisi = True
                    Exit For
    
                End If
    
            End If
    
        Next i
    
    End With
    
    If blnRevisi Then
    
              If MsgBox("Approved data detected. Submit as revision?", _
          vbQuestion + vbYesNo + vbDefaultButton2, _
          "Confirmation") = vbNo Then
            
            MousePointer = vbDefault
            
              If dbTransfer.State = 1 Then
                    dbTransfer.RollbackTrans
                    dbTransfer.Close
                End If
                
            Exit Sub

        End If

    End If
    
       
    blnRevisi = False
    
    ApproveCount = 0
    
    Erase arrApprove


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
        
         If statusProses = "APPROVE" Then
    
            ApproveCount = ApproveCount + 1
        
            ReDim Preserve arrApprove(1 To ApproveCount)
        
            arrApprove(ApproveCount).ItemCode = _
                Trim(.TextMatrix(i, bteColProdCode))
        
            arrApprove(ApproveCount).lotno = _
                Trim(.TextMatrix(i, bteColLotNo))
        
            arrApprove(ApproveCount).Qty = _
                Val(.TextMatrix(i, bteColQty))
        
        End If
    
NextData:
    
    Next i

    End With
    
    ' =====================================================
    ' COMMIT
    ' =====================================================
    dbTransfer.CommitTrans
    dbTransfer.Close
    
    If Export_ToCsv() = False Then
    
        Call RollbackApprove
    
        MsgBox lblErrMsg, vbExclamation, "FTP Upload Failed"
    
        MousePointer = vbDefault
    
        Exit Sub
    
    End If
  
        Browse

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
    
    If GetAdminStatus() = 1 Then
        scheduledate1.Enabled = True
        scheduledate2.Enabled = True
    Else
        scheduledate1.Enabled = False
        scheduledate2.Enabled = False
    End If


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

Public Function Export_ToCsv() As Boolean
   
    On Error GoTo Errhandle
     
    Export_ToCsv = False
         
    Dim rsCsv As New ADODB.Recordset
    Dim RsFtp As New ADODB.Recordset

    Dim sql As String
    Dim ftpSql As String

    Dim ExportFolder As String
    Dim FullPath As String

    Dim FileNo As Integer

    Dim scheduledate As String
    Dim PrevScheduleDate As String

    Dim LineCode As String
    Dim ItemCode As String
    Dim SerialNumber As String
    Dim StatusData As String

    Dim ftpHost As String
    Dim ftpUser As String
    Dim ftpPass As String
    Dim ftpFolder As String
    Dim ftpRemarks As String
    
    Dim ftpCommand As String

    Dim filename As String
    
    Dim ErrMsg As String

    Dim i As Long
    
    Dim HeaderFile As String
    Dim DetailFile As String
    Dim HeaderFullPath As String
    Dim DetailFullPath As String

    
    Dim FTPStatus As String
    Dim ExportStatus As String
    
    Dim RevisionNo As Long
    
    Dim Qty As Long
    Dim SerialCount As Long
    Dim LogSql As String
    
    ' =====================================================
    ' EXPORT FOLDER
    ' =====================================================

    ExportFolder = App.path & "\IFProductionSchedule"

    If Dir(ExportFolder, vbDirectory) = "" Then
        MkDir ExportFolder
    End If

    If Right(ExportFolder, 1) <> "\" Then
        ExportFolder = ExportFolder & "\"
    End If

    ' =====================================================
    ' QUERY EXPORT
    ' =====================================================

    sql = ""

    sql = sql & " EXEC sp_ProdScheduleInterface_ExportCsv "
    sql = sql & " '" & Trim(cbocust.Text) & "', "
    sql = sql & " '" & Trim(cbolinecd.Text) & "', "
    sql = sql & " '" & Format(scheduledate1.Value, "yyyy-mm-dd") & "', "
    sql = sql & " '" & Format(scheduledate2.Value, "yyyy-mm-dd") & "', "
    sql = sql & " '" & UCase(Trim(cboStatus.Text)) & "' "
    
    Dim rsdetail As ADODB.Recordset

    rsCsv.Open sql, Db, adOpenForwardOnly, adLockReadOnly
    
    If rsCsv.EOF Then
    
        MsgBox "No data to export.", vbExclamation
    
        rsCsv.Close
    
        Exit Function
    
    End If

    ' =====================================================
    ' GET FTP CONFIG
    ' =====================================================
    
    ftpSql = "EXEC sp_GetFTPConfig"
    
    RsFtp.Open ftpSql, Db, adOpenForwardOnly, adLockReadOnly
    
    If RsFtp.EOF Then
    
        MsgBox _
            "FTP Config Not Found", _
            vbCritical
    
        Exit Function
    
    End If
    
    ftpHost = Trim("" & RsFtp.Fields("FTP_HOST").Value)
    
    ftpUser = Trim("" & RsFtp.Fields("FTP_USER").Value)
    
    ftpPass = Trim("" & RsFtp.Fields("FTP_PASS").Value)
    
    ftpFolder = Trim("" & RsFtp.Fields("FTP_FOLDER").Value)
    
    ftpRemarks = Trim("" & RsFtp.Fields("FTP_REMARKS").Value)
    
    RsFtp.Close

    
    '=========================================
    ' HEADER RESULTSET
    '=========================================

'Dim RevisionNo As Long
Dim CheckFile As String

'scheduledate = sche
'    Trim("" & rsCsv.Fields("Schedule Date").Value)

scheduledate = Format$(Now, "yyyymmdd")
LineCode = Right$("00" & Mid$(cbolinecd.Text, 5, 1), 3) 'cbolinecd.Text

'=========================================
' GET REVISION NUMBER
'=========================================

RevisionNo = 0

Do

    If RevisionNo = 0 Then

        CheckFile = _
            ExportFolder & _
            "SCHEDULE_" & _
             Trim(LineCode) & "_" & _
            scheduledate & _
            ".csv"

    Else

        CheckFile = _
            ExportFolder & _
            "SCHEDULE_" & _
            Trim(LineCode) & "_" & _
            scheduledate & _
            "_R_" & _
            RevisionNo & _
            ".csv"

    End If

    If Dir(CheckFile) = "" Then Exit Do

    RevisionNo = RevisionNo + 1

Loop

'=========================================
' HEADER FILE NAME
'=========================================

If RevisionNo = 0 Then

    filename = _
        "SCHEDULE_" & _
        Trim(LineCode) & "_" & _
         Trim(scheduledate) & _
        ".csv"

Else

    filename = _
        "SCHEDULE_" & _
         Trim(LineCode) & "_" & _
        scheduledate & _
        "_R_" & _
        RevisionNo & _
        ".csv"

End If


FullPath = ExportFolder & filename

HeaderFile = filename
HeaderFullPath = FullPath

FileNo = FreeFile

Open FullPath For Output As #FileNo

Print #FileNo, _
    "Schedule Date,Item Code,Qty"

Qty = 0
SerialCount = 0

Do While Not rsCsv.EOF

    Qty = Qty + CLng(Val(rsCsv.Fields("Qty").Value))
    
   Print #FileNo, _
    rsCsv.Fields("Schedule Date").Value & "," & _
    """" & rsCsv.Fields("Item Code").Value & """," & _
    rsCsv.Fields("Qty").Value

    rsCsv.MoveNext

    
Loop

Close #FileNo

'=========================================
' DETAIL RESULTSET
'=========================================

On Error Resume Next
Set rsdetail = rsCsv.NextRecordset
On Error GoTo Errhandle

If Not rsdetail Is Nothing Then

    If RevisionNo = 0 Then

        filename = _
            "SERIAL_" & _
            Trim(LineCode) & "_" & _
            scheduledate & _
            ".csv"

    Else

        filename = _
            "SERIAL_" & _
            Trim(LineCode) & "_" & _
            scheduledate & _
            "_R_" & _
            RevisionNo & _
            ".csv"

    End If
    
   

    FullPath = ExportFolder & filename
    
    DetailFile = filename
    DetailFullPath = FullPath
    
    FileNo = FreeFile

    Open FullPath For Output As #FileNo

    Print #FileNo, _
        "Schedule Date,Item Code,Serial Number,Status"

    Do While Not rsdetail.EOF
        ItemCode = """" & rsdetail.Fields("Item Code").Value & """"
        
        SerialNumber = """" & rsdetail.Fields("Serial Number").Value & """"

        StatusData = _
            Trim("" & rsdetail.Fields("Status").Value)

        Print #FileNo, _
            rsdetail.Fields("Schedule Date").Value & "," & _
            ItemCode & "," & _
            SerialNumber & "," & _
            StatusData
            
        SerialCount = SerialCount + 1
            
        rsdetail.MoveNext

    Loop

    Close #FileNo

End If

'=========================================
' CLEANUP
'=========================================

If Not rsdetail Is Nothing Then
    If rsdetail.State = adStateOpen Then
        rsdetail.Close
    End If
End If

If rsCsv.State = adStateOpen Then
    rsCsv.Close
End If

    ExportStatus = "SUCCESS"
    FTPStatus = "PENDING"
    
    LogSql = ""
    
    LogSql = LogSql & " EXEC sp_Interface_Export_Log_Ins "
    LogSql = LogSql & " '" & Trim(cbocust.Text) & "', "
    LogSql = LogSql & " '" & Trim(cbolinecd.Text) & "', "
    LogSql = LogSql & " '" & scheduledate & "', "
    LogSql = LogSql & RevisionNo & ", "
    LogSql = LogSql & " '" & HeaderFile & "', "
    LogSql = LogSql & " '" & DetailFile & "', "
    LogSql = LogSql & " '" & ExportStatus & "', "
    LogSql = LogSql & " '" & FTPStatus & "', "
    LogSql = LogSql & Qty & ", "
    LogSql = LogSql & SerialCount & ", "
    LogSql = LogSql & " '" & Trim(userName) & "', "
    LogSql = LogSql & " '' "
    
    Dim rsLog As New ADODB.Recordset
    Dim LogID As Long
    
    rsLog.Open LogSql, Db, adOpenForwardOnly, adLockReadOnly
    
    If Not rsLog.EOF Then
        LogID = rsLog.Fields("Log_ID").Value
    End If
    
    rsLog.Close
    
'    If UCase$(Left$(FileName, 9)) = "SCHEDULE_" Then
'
'        FileType = "ORDER"
'
'    ElseIf UCase$(Left$(FileName, 7)) = "SERIAL_" Then
'
'        FileType = "SERIAL"
'
'    End If

    
    '=========================================
    ' FTP HEADER
    '=========================================
    
    ftpFolder = GetFTPFolder(HeaderFile)

    If UploadFTP( _
        HeaderFullPath, _
        HeaderFile, _
        ftpHost, _
        ftpUser, _
        ftpPass, _
        ftpFolder, _
        ftpRemarks, _
        ErrMsg) = False Then
    
        Db.Execute _
            "EXEC dbo.sp_Interface_Export_Log_UpdFTPStatus " & _
            "@LogID = " & LogID & _
            ", @FTPStatus = 'FAILED'" & _
            ", @UserID = '" & Replace(Trim$(userLogin), "'", "''") & "'" & _
            ", @ErrorMessage = '" & Replace(ErrMsg, "'", "''") & "'"
    
        ' Hapus file CSV yang sudah dibuat
        On Error Resume Next
    
        If Len(Dir$(HeaderFullPath)) > 0 Then
            Kill HeaderFullPath
        End If
    
        If Len(Dir$(DetailFullPath)) > 0 Then
            Kill DetailFullPath
        End If
    
        On Error GoTo 0
    
        lblErrMsg = ErrMsg
    
        Exit Function
    
    End If

    '=========================================
    ' FTP DETAIL
    '=========================================
    
    ftpFolder = GetFTPFolder(DetailFile)
    
    If UploadFTP( _
        DetailFullPath, _
        DetailFile, _
        ftpHost, _
        ftpUser, _
        ftpPass, _
        ftpFolder, _
        ftpRemarks, _
        ErrMsg) = False Then
    
    Db.Execute _
            "EXEC dbo.sp_Interface_Export_Log_UpdFTPStatus " & _
            "@LogID = " & LogID & _
            ", @FTPStatus = 'FAILED'" & _
            ", @UserID = '" & Replace(Trim$(userLogin), "'", "''") & "'" & _
            ", @ErrorMessage = '" & Replace(ErrMsg, "'", "''") & "'"
    
'        MsgBox ErrMsg, vbCritical

        lblErrMsg = ErrMsg
        
        Exit Function
    
    End If
    
    '=========================================
    ' BOTH SUCCESS
    '=========================================
    
    Db.Execute _
        "EXEC sp_Interface_Export_Log_UpdFTPStatus " & _
        LogID & _
        ", 'SUCCESS', '" & Trim(userLogin) & "'"
        
        MsgBox "Schedule Production sent successfully.", _
        vbInformation, _
        "Success"
        
        lblErrMsg = ErrMsg
           
        Export_ToCsv = True
          
        Exit Function

Errhandle:

    MsgBox _
        err.number & " - " & err.Description, _
        vbCritical

End Function

Sub Header()
    Dim i As Long
    
    bteColDate = 0
    bteColProdCode = 1
    bteColPart = 2
    bteColDesc = 3
    bteColLotNo = 4
    bteColQty = 5
    bteColUnitCls = 6
    bteColUnit = 7
    BteColSerialFrom = 8
    BteColSerialTo = 9
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
      
      .ColWidth(bteColDate) = 1450
      .ColWidth(bteColProdCode) = 1400
      .ColWidth(bteColPart) = 1400
      .ColWidth(bteColDesc) = 3000
      .ColWidth(bteColLotNo) = 1000
      .ColWidth(bteColQty) = 1230
      .ColWidth(bteColUnit) = 650
      .ColWidth(BteColSerialFrom) = 1100
      .ColWidth(BteColSerialTo) = 1100
      .ColWidth(bteColRemark) = 3250
      .ColWidth(bteColAuto) = 1000
      .ColWidth(bteColCustCode) = 1200
      .ColWidth(bteColCustName) = 2800
      .ColWidth(bteColPONo) = 2000
      .ColWidth(bteColApproved) = 1050
      .ColWidth(bteColApprovedDate) = 2150
      .ColWidth(bteColApprovedUser) = 1500
      .ColWidth(bteColVoid) = 1000
      .ColWidth(bteColVoidDate) = 2150
      .ColWidth(bteColVoidUser) = 1500
      
      .TextMatrix(0, bteColDate) = "Schedule Date"
      .TextMatrix(0, bteColPart) = "Part Number"
      .TextMatrix(0, bteColProdCode) = "Product Code"
      .TextMatrix(0, bteColDesc) = "Description"
      .TextMatrix(0, bteColLotNo) = "Lot No"
      .TextMatrix(0, bteColQty) = "Qty"
      .TextMatrix(0, bteColUnitCls) = "UnitCls"
      .TextMatrix(0, bteColUnit) = "Unit"
      .TextMatrix(0, BteColSerialFrom) = "Serial From"
      .TextMatrix(0, BteColSerialTo) = "Serial To"
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

    sqlcust = "EXEC dbo.sp_GetCompanyCodeProdScheduleInterface @CompanyCode = '" & cboFactory.Text & "'"

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
                
                .TextMatrix(i, bteColDate) = Format(Trim(rsGrid("schedule_date")), "dd MMM yyyy")
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
                
                 .TextMatrix(i, bteColApprovedDate) = Format(Trim(rsGrid("Approved_Date")), "dd MMM yyyy HH:nn:ss") & ""
                 .TextMatrix(i, bteColApprovedUser) = rsGrid("Approved_User") & ""
                 
                .Cell(flexcpBackColor, i, bteColVoid) = vbWhite
                .ColDataType(bteColVoid) = flexDTBoolean
                If IsNull(rsGrid.Fields("Void_Date").Value) Then
                    .TextMatrix(i, bteColVoid) = False
                Else
                    .TextMatrix(i, bteColVoid) = True
                End If
                
                .TextMatrix(i, bteColVoidDate) = Format(Trim(rsGrid("Void_Date")), "dd MMM yyyy HH:nn:ss") & ""
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

'Private Sub UploadFTP( _
'    ByVal FullPath As String, _
'    ByVal filename As String, _
'    ByVal ftpHost As String, _
'    ByVal ftpUser As String, _
'    ByVal ftpPass As String, _
'    ByVal ftpFolder As String)
'
'    Dim ftpCommand As String
'    Dim i As Long
'
'    Inet1.Protocol = icFTP
'
'    Inet1.URL = ftpHost
'
'    Inet1.userName = ftpUser
'
'    Inet1.Password = ftpPass
'
'    ftpCommand = _
'        "PUT """ & _
'        FullPath & _
'        """ " & _
'        ftpFolder & _
'        filename
'
'    Debug.Print ftpCommand
'
'    Inet1.Execute , ftpCommand
'
'    Do While Inet1.StillExecuting
'        DoEvents
'    Loop
'
'    For i = 1 To 500000
'        DoEvents
'    Next i
'
'End Sub

Private Function UploadFTP( _
ByVal FullPath As String, _
ByVal filename As String, _
ByVal ftpHost As String, _
ByVal ftpUser As String, _
ByVal ftpPass As String, _
ByVal ftpFolder As String, _
ByVal ftpRemarks As String, _
Optional ByRef ErrMsg As String = "") As Boolean


On Error GoTo Errhandle

Dim ftpCommand As String
Dim StartTime As Single
Dim lsResponse As String

UploadFTP = False
ErrMsg = ""

' ==========================================
' FILE CHECK
' ==========================================

If Dir(FullPath) = "" Then

    ErrMsg = "Local file not found : " & FullPath

    Exit Function

End If

' ==========================================
' TRIAL MODE
' ==========================================
If UCase$(Trim$(ftpRemarks)) = "TRIAL" Then

    ftpFolder = Replace(ftpFolder, "\", "/")

    If Left$(ftpFolder, 1) <> "/" Then
        ftpFolder = "/" & ftpFolder
    End If

    If Right$(ftpFolder, 1) <> "/" Then
        ftpFolder = ftpFolder & "/"
    End If

    ftpCommand = _
        "PUT """ & _
        FullPath & _
        """ """ & _
        ftpFolder & _
        filename & """"

    Debug.Print String(80, "=")
    Debug.Print "FTP MODE     : TRIAL"
    Debug.Print "FTP COMMAND  : " & ftpCommand
    Debug.Print "SOURCE FILE  : " & FullPath

    Inet1.URL = _
    "ftp://" & _
    EncodeURL(ftpUser) & ":" & _
    EncodeURL(ftpPass) & "@" & _
    ftpHost

    StartTime = Timer

    Debug.Print "INET URL : " & Inet1.URL
    
    Inet1.Execute , ftpCommand

    Do While Inet1.StillExecuting

        DoEvents

        If Abs(Timer - StartTime) > 120 Then

            UploadFTP = False
        
            ErrMsg = "FTP Timeout (120 Seconds)"
        
            WriteLog "STATUS        : FAILED"
            WriteLog "ERROR         : " & ErrMsg
            WriteLog "FILE          : " & filename
        
            Exit Function
        
        End If

    Loop

    Debug.Print "FTP RESPONSE : " & Inet1.ResponseInfo
    
    WriteLog String(80, "=")
    WriteLog "DATE/TIME     : " & Format(Now, "yyyy-mm-dd HH:mm:ss")
    WriteLog "FTP MODE      : TRIAL"
    WriteLog "SOURCE FILE   : " & filename
    WriteLog "DEST FILE     : " & ftpFolder & filename
    WriteLog "FTP RESPONSE  : " & Inet1.ResponseInfo

    
    lsResponse = Trim$(Inet1.ResponseInfo)
    
    If InStr(1, lsResponse, "completed successfully", vbTextCompare) > 0 Then
        UploadFTP = True
        WriteLog "STATUS        : SUCCESS"
        ErrMsg = lsResponse
    Else
        UploadFTP = False
        WriteLog "STATUS        : FAILED"
        ErrMsg = lsResponse
    End If
    
    Exit Function

End If

' ==========================================
' LIVE FTP MODE (ftp.exe)
' ==========================================

Dim ScriptFile As String
Dim LogFile As String
Dim FF As Integer
Dim Wsh As Object
Dim Result As String
Dim FSO As Object
Dim TS As Object

UploadFTP = False

If Right$(ftpFolder, 1) = "\" Then
    ftpFolder = Left$(ftpFolder, Len(ftpFolder) - 1)
End If

ScriptFile = App.path & "\ftp_upload.txt"
LogFile = App.path & "\ftp_upload.log"

FF = FreeFile

Open ScriptFile For Output As #FF

Print #FF, "open " & ftpHost
Print #FF, "user " & ftpUser & " " & ftpPass
Print #FF, "binary"

Print #FF, "put """ & _
            FullPath & _
            """ " & _
            ftpFolder & "\" & filename

Print #FF, "bye"

Close #FF

Debug.Print String(80, "=")
Debug.Print "FTP HOST     : " & ftpHost
Debug.Print "FTP USER     : " & ftpUser
Debug.Print "FTP FOLDER   : " & ftpFolder
Debug.Print "LOCAL FILE   : " & FullPath
Debug.Print "REMOTE FILE  : " & ftpFolder & "\" & filename

Set Wsh = CreateObject("WScript.Shell")

Wsh.Run _
    "cmd /c ftp -n -s:""" & _
    ScriptFile & _
    """ > """ & _
    LogFile & _
    """ 2>&1", _
    0, _
    True

Set FSO = CreateObject("Scripting.FileSystemObject")

Result = ""

If FSO.FileExists(LogFile) Then

    Set TS = FSO.OpenTextFile(LogFile, 1)

    Result = TS.ReadAll

    TS.Close

End If

Debug.Print Result

WriteLog String(80, "=")
WriteLog "DATE/TIME     : " & Format(Now, "yyyy-mm-dd HH:mm:ss")
WriteLog "FTP MODE      : LIVE"
WriteLog "SOURCE FILE   : " & filename
WriteLog "DEST FILE     : " & ftpFolder & "\" & filename
WriteLog "FTP RESPONSE  : "
WriteLog Result

If InStr(1, Result, "226", vbTextCompare) > 0 _
Or InStr(1, Result, "Transfer complete", vbTextCompare) > 0 _
Or InStr(1, Result, "completed successfully", vbTextCompare) > 0 Then

    UploadFTP = True
    ErrMsg = "Upload completed successfully."

    WriteLog "STATUS        : SUCCESS"

Else

    UploadFTP = False
    ErrMsg = Trim$(Result)

    WriteLog "STATUS        : FAILED"

End If

On Error Resume Next

If FSO.FileExists(ScriptFile) Then Kill ScriptFile
If FSO.FileExists(LogFile) Then Kill LogFile

Set TS = Nothing
Set FSO = Nothing
Set Wsh = Nothing

On Error GoTo Errhandle

Exit Function

Errhandle:

ErrMsg = "VB Error " & err.number & " - " & err.Description

WriteLog String(80, "=")
WriteLog "DATE/TIME     : " & Format(Now, "yyyy-mm-dd HH:mm:ss")
WriteLog "FTP MODE      : " & ftpRemarks
WriteLog "SOURCE FILE   : " & filename
WriteLog "STATUS        : FAILED"
WriteLog "ERROR         : " & ErrMsg

UploadFTP = False

End Function

Private Sub WriteLog(ByVal msg As String)

    On Error Resume Next

    Dim LogFolder As String
    Dim LogFile As String
    Dim F As Integer

    LogFolder = App.path & "\IFProductionSchedule\Log"

    If Dir(LogFolder, vbDirectory) = "" Then
        MkDir LogFolder
    End If

    LogFile = LogFolder & "\Log_" & _
              Format(Date, "yyyymmdd") & ".txt"

    F = FreeFile

    Open LogFile For Append As #F

    Print #F, msg

    Close #F

End Sub


Private Function EncodeURL(ByVal Text As String) As String

    Text = Replace(Text, "%", "%25")
    Text = Replace(Text, "@", "%40")
    Text = Replace(Text, " ", "%20")
    Text = Replace(Text, "#", "%23")
    Text = Replace(Text, "&", "%26")
    Text = Replace(Text, "+", "%2B")
    Text = Replace(Text, "?", "%3F")
    Text = Replace(Text, "=", "%3D")

    EncodeURL = Text

End Function

Private Function GetFTPFolder(ByVal filename As String) As String

    Dim RS As ADODB.Recordset
    Dim sql As String

    On Error GoTo ErrHandler

    Set RS = New ADODB.Recordset

    sql = "EXEC sp_GetFTPFileType_ProdScheduleInf '" & _
          Replace(filename, "'", "''") & "'"

    RS.Open sql, Db, adOpenForwardOnly, adLockReadOnly

    If Not RS.EOF Then

        GetFTPFolder = Trim$(RS!FTP_Folder)

    End If

    RS.Close
    Set RS = Nothing

    Exit Function

ErrHandler:

    GetFTPFolder = ""

    If Not RS Is Nothing Then

        If RS.State = adStateOpen Then RS.Close

        Set RS = Nothing

    End If

End Function

Private Function GetAdminStatus() As Integer

    Dim RS As ADODB.Recordset

    Set RS = Db.Execute( _
        "SELECT ISNULL(Status_Admin,0) AS Status_Admin " & _
        "FROM dbo.User_Setup " & _
        "WHERE Username = '" & Replace(userLogin, "'", "''") & "'")

    If Not RS.EOF Then
        GetAdminStatus = RS!status_Admin
    Else
        GetAdminStatus = 0
    End If

    RS.Close
    Set RS = Nothing

End Function

Private Sub RollbackApprove()

    Dim j As Long
    Dim sql As String

    For j = 1 To ApproveCount

        sql = ""

        sql = sql & " EXEC sp_ProdScheduleInf_Status "
        sql = sql & " '" & Trim(cbocust.Text) & "', "
        sql = sql & " '" & Trim(cbolinecd.Text) & "', "
        sql = sql & " '" & Format(scheduledate1.Value, "yyyy-mm-dd") & "', "
        sql = sql & " '" & Format(scheduledate2.Value, "yyyy-mm-dd") & "', "
        sql = sql & " '" & arrApprove(j).ItemCode & "', "
        sql = sql & " '" & arrApprove(j).lotno & "', "
        sql = sql & arrApprove(j).Qty & ", "
        sql = sql & " 'UNAPPROVE', "
        sql = sql & " '" & userLogin & "' "

        Db.Execute sql

    Next j

End Sub

