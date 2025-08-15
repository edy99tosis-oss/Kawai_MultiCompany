VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPROther 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Purchase Request (Others)"
   ClientHeight    =   11010
   ClientLeft      =   90
   ClientTop       =   390
   ClientWidth     =   15165
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPROther.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   11010
   ScaleWidth      =   15165
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1515
      Left            =   728
      TabIndex        =   39
      Top             =   1080
      Width           =   13785
      Begin VB.TextBox lblSection 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDDFE3&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   750
         Width           =   3645
      End
      Begin VB.TextBox lblPerson 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDDFE3&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   300
         Width           =   3645
      End
      Begin VB.TextBox lblDept 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDDFE3&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   300
         Width           =   3645
      End
      Begin MSComCtl2.DTPicker Period 
         Height          =   315
         Left            =   2040
         TabIndex        =   4
         Top             =   660
         Width           =   1305
         _ExtentX        =   2302
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
         Format          =   141230083
         UpDown          =   -1  'True
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker RequestDate1 
         Height          =   315
         Left            =   2040
         TabIndex        =   5
         Top             =   1080
         Width           =   1500
         _ExtentX        =   2646
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
         Format          =   141230083
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker RequestDate2 
         Height          =   315
         Left            =   4020
         TabIndex        =   6
         Top             =   1080
         Width           =   1500
         _ExtentX        =   2646
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
         Format          =   141230083
         CurrentDate     =   37798
      End
      Begin VB.Line Line2 
         Index           =   2
         X1              =   9480
         X2              =   13185
         Y1              =   975
         Y2              =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Section"
         Height          =   195
         Index           =   1
         Left            =   7620
         TabIndex        =   50
         Top             =   735
         Width           =   630
      End
      Begin MSForms.ComboBox cboSection 
         Height          =   315
         Left            =   8385
         TabIndex        =   49
         Top             =   690
         Width           =   975
         VariousPropertyBits=   746604571
         MaxLength       =   2
         DisplayStyle    =   3
         Size            =   "1720;556"
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
         Caption         =   "Period"
         Height          =   195
         Index           =   3
         Left            =   195
         TabIndex        =   44
         Top             =   705
         Width           =   1710
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Request Date From "
         Height          =   195
         Index           =   4
         Left            =   195
         TabIndex        =   43
         Top             =   1130
         Width           =   1710
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "to "
         Height          =   195
         Index           =   6
         Left            =   3652
         TabIndex        =   42
         Top             =   1110
         Width           =   255
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   3120
         X2              =   6825
         Y1              =   525
         Y2              =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Person in Charge"
         Height          =   195
         Index           =   2
         Left            =   195
         TabIndex        =   41
         Top             =   280
         Width           =   1725
      End
      Begin MSForms.ComboBox cboPerson 
         Height          =   315
         Left            =   2040
         TabIndex        =   0
         Top             =   240
         Width           =   975
         VariousPropertyBits=   746604571
         MaxLength       =   2
         DisplayStyle    =   3
         Size            =   "1720;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboDept 
         Height          =   315
         Left            =   8385
         TabIndex        =   2
         Top             =   240
         Width           =   975
         VariousPropertyBits=   746604571
         MaxLength       =   2
         DisplayStyle    =   3
         Size            =   "1720;556"
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
         Caption         =   "Department"
         Height          =   195
         Index           =   0
         Left            =   7230
         TabIndex        =   40
         Top             =   285
         Width           =   1020
      End
      Begin VB.Line Line2 
         Index           =   1
         X1              =   9480
         X2              =   13185
         Y1              =   525
         Y2              =   525
      End
   End
   Begin VB.TextBox txtItemCd 
      Height          =   315
      Left            =   2400
      MaxLength       =   15
      TabIndex        =   16
      Top             =   10425
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtDesc 
      Height          =   315
      Left            =   4800
      MaxLength       =   50
      TabIndex        =   17
      Top             =   10440
      Visible         =   0   'False
      Width           =   3180
   End
   Begin VB.CommandButton cmdReport 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Preview"
      Height          =   375
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   10020
      Width           =   1125
   End
   Begin VB.TextBox txtPurpose 
      Height          =   315
      Left            =   11160
      MaxLength       =   100
      TabIndex        =   22
      Top             =   8865
      Width           =   3195
   End
   Begin VB.CommandButton command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Cancel"
      Height          =   375
      Index           =   3
      Left            =   10862
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   10020
      Width           =   1125
   End
   Begin VB.TextBox txtRequestNo 
      Height          =   315
      Left            =   3555
      MaxLength       =   25
      TabIndex        =   8
      Top             =   2760
      Width           =   2085
   End
   Begin VB.CommandButton command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Create"
      Height          =   375
      Index           =   0
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2730
      Width           =   1125
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FDDFE3&
      Height          =   533
      Left            =   728
      TabIndex        =   29
      Top             =   9311
      Width           =   13785
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
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   180
         Width           =   13515
      End
   End
   Begin VB.CommandButton command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Submit"
      Height          =   375
      Index           =   1
      Left            =   13388
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   10020
      Width           =   1125
   End
   Begin VB.CommandButton cmdSubMenu 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Sub &Menu"
      Height          =   375
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   10020
      Width           =   1125
   End
   Begin VB.CommandButton command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Clea&r"
      Height          =   375
      Index           =   2
      Left            =   12124
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   10020
      Width           =   1125
   End
   Begin VB.TextBox txtQty 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   8910
      MaxLength       =   12
      TabIndex        =   20
      Text            =   "9,999,999.99"
      Top             =   8865
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker RequestDate 
      Height          =   315
      Left            =   7710
      TabIndex        =   10
      Top             =   2760
      Width           =   1500
      _ExtentX        =   2646
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
      Format          =   141230083
      CurrentDate     =   37798
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   12668
      TabIndex        =   28
      Top             =   360
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   714
   End
   Begin MSComCtl2.DTPicker DelDate 
      Height          =   315
      Left            =   7395
      TabIndex        =   19
      Top             =   8865
      Width           =   1485
      _ExtentX        =   2619
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
      Format          =   141230083
      CurrentDate     =   37798
   End
   Begin VSFlex8Ctl.VSFlexGrid grid 
      Height          =   4650
      Left            =   735
      TabIndex        =   12
      Top             =   3240
      Width           =   13785
      _cx             =   24315
      _cy             =   8202
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
      SelectionMode   =   0
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
      ScrollTrack     =   0   'False
      ScrollBars      =   2
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
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product Cls"
      Height          =   195
      Left            =   840
      TabIndex        =   48
      Top             =   8025
      Width           =   975
   End
   Begin MSForms.ComboBox cboCls 
      Height          =   315
      Left            =   2040
      TabIndex        =   13
      Top             =   7980
      Width           =   1335
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "2355;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code"
      Height          =   195
      Left            =   840
      TabIndex        =   47
      Top             =   8490
      Width           =   1155
   End
   Begin MSForms.ComboBox cboItemCode 
      Height          =   315
      Left            =   840
      TabIndex        =   14
      Top             =   8865
      Width           =   2550
      VariousPropertyBits=   612386843
      MaxLength       =   25
      DisplayStyle    =   3
      Size            =   "4498;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Class"
      Height          =   195
      Left            =   6690
      TabIndex        =   46
      Top             =   8490
      Width           =   465
   End
   Begin MSForms.ComboBox cboClass 
      Height          =   315
      Left            =   6465
      TabIndex        =   18
      Top             =   8865
      Width           =   900
      VariousPropertyBits=   612386843
      MaxLength       =   3
      DisplayStyle    =   7
      Size            =   "1587;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cboItemName 
      Height          =   315
      Left            =   3465
      TabIndex        =   15
      Top             =   8865
      Width           =   2970
      VariousPropertyBits=   612386843
      MaxLength       =   50
      DisplayStyle    =   3
      Size            =   "5239;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblfix 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Status Fix "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Left            =   9600
      TabIndex        =   45
      Top             =   2805
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery Date"
      Height          =   195
      Index           =   0
      Left            =   7395
      TabIndex        =   38
      Top             =   8490
      Width           =   1185
   End
   Begin MSForms.ComboBox cboUnit 
      Height          =   315
      Left            =   10275
      TabIndex        =   21
      Top             =   8865
      Width           =   870
      VariousPropertyBits=   746604569
      MaxLength       =   15
      DisplayStyle    =   7
      Size            =   "1535;556"
      ListRows        =   15
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontEffects     =   1073750016
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks/Purpose"
      Height          =   195
      Index           =   1
      Left            =   11145
      TabIndex        =   37
      Top             =   8490
      Width           =   1530
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Request No"
      Height          =   255
      Index           =   1
      Left            =   2355
      TabIndex        =   36
      Top             =   2805
      Width           =   1170
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Request Date"
      Height          =   255
      Index           =   2
      Left            =   6360
      TabIndex        =   35
      Top             =   2805
      Width           =   1305
   End
   Begin MSForms.ComboBox cboRequestNo 
      Height          =   315
      Left            =   3555
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2760
      Width           =   2370
      VariousPropertyBits=   612386843
      DisplayStyle    =   3
      Size            =   "4180;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox combo1 
      Height          =   315
      Left            =   735
      TabIndex        =   7
      Top             =   2760
      Width           =   1215
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "2143;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Request (Others)"
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
      Left            =   690
      TabIndex        =   34
      Top             =   360
      Width           =   13800
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00A6D2FF&
      Height          =   510
      Index           =   0
      Left            =   735
      Top             =   8760
      Width           =   13785
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unit"
      Height          =   195
      Left            =   10545
      TabIndex        =   33
      Top             =   8490
      Width           =   330
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      Height          =   195
      Left            =   3450
      TabIndex        =   32
      Top             =   8490
      Width           =   1200
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Request Qty"
      Height          =   195
      Left            =   9195
      TabIndex        =   31
      Top             =   8490
      Width           =   1050
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00A6D2FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   0
      Left            =   735
      Top             =   8400
      Width           =   13785
   End
End
Attribute VB_Name = "frmPROther"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'0: direct , 1: others


Option Explicit


Dim sql As String, sqlGrid As String
Dim RS As New ADODB.Recordset, rsGrid As New ADODB.Recordset
Dim i As Long, lblQty As Double, lblseqno As Long
Dim ubah As Boolean, ada As Boolean, statusfix As Byte, temptgl As Byte

Sub adtocboaccount()
    Dim sqlitem As String
    Dim RsItem As New Recordset

'    sqlitem = "select accountno, accountname from " & EZRGLDb & ".dbo.accountmaster"
'    Set RsItem = Db.Execute(sqlitem)
'    With CboAccount
'        .clear
'        .ColumnCount = 2
'        .ColumnWidths = "60pt;240pt"
'        .ListWidth = 300
'        .ListRows = 15
'        i = 0
'        Do While Not RsItem.EOF
'            .AddItem
'            .List(i, 0) = Trim(RsItem("accountno"))
'            .List(i, 1) = Trim(RsItem("accountname"))
'            RsItem.MoveNext
'            i = i + 1
'        Loop
'    End With
    Set RsItem = Nothing
End Sub




Sub Kosong()
    LblErrMsg = ""
    cboPerson.Text = "": lblPerson.Text = ""
    cboDept.Text = "": lblDept.Text = ""
    cboSection.Text = ""
    Period.Value = Format(Now, "MMM yyyy")
    temptgl = Period.Month
    requestdate1.Value = Format(Now, "yyyy-mm-01")
    requestdate1.Enabled = True
    requestdate2.Value = Format(Now, "dd MMM yyyy")
    requestdate2.Enabled = True
    txtRequestNo.Text = "": cborequestno.clear
    RequestDate.Value = Format(Now, "dd MMM yyyy")
    RequestDate.Enabled = True
    
    ubah = False: ada = False
    statusfix = 0: Call kunci(False)
    Call kosongBwh
    Call Header
End Sub

Sub kosongBwh()
    CboItemCode.Text = ""
    cboItemName.Text = ""
    cboClass.ListIndex = -1
    DelDate.Value = Format(Now, "dd MMM yyyy")
    txtQty.Text = ""
    lblQty = 0
    cbounit.ListIndex = -1
    txtPurpose.Text = ""
    CboItemCode.Enabled = True
    cboItemName.Enabled = True
    lblseqno = 0
    cboCls.ListIndex = 0
    cboCls.Enabled = True
End Sub

Sub adtocboperson()
Dim sqlperson As String
Dim rsperson As New Recordset

    sqlperson = "select * from PersonInCharge_Cls order by PersonInCharge_cls"
    Set rsperson = Db.Execute(sqlperson)
    
    With cboPerson
        .clear
        .columnCount = 2
        .ColumnWidths = "30pt;120pt"
        .ListWidth = 150
        .ListRows = 15
        
        i = 0
        Do While Not rsperson.EOF
            .AddItem
            .List(i, 0) = Trim(rsperson("PersonInCharge_cls"))
            .List(i, 1) = IIf(IsNull(rsperson("description")), "", Trim(rsperson("description")))
            rsperson.MoveNext
            i = i + 1
        Loop
    End With
    Set rsperson = Nothing
End Sub
Private Sub addCboSection()
    Dim adoRs As New ADODB.Recordset
    Dim intCount As Integer
    
    Me.MousePointer = vbHourglass
    On Error GoTo errHandler
    
    With cboSection
        .clear
        .columnCount = 2
        .ColumnWidths = "30pt;70pt"
        .ListWidth = 100
        .ListRows = 600
        
        sql = "select Section_Cls,Description FROM section_cls"
        adoRs.Open sql, Db, adOpenDynamic, adLockReadOnly, adCmdText
        While Not adoRs.EOF
            .AddItem ""
            .Column(1, intCount) = Trim(adoRs.Fields("Description"))
            .Column(0, intCount) = Trim(adoRs.Fields("Section_Cls"))
            intCount = intCount + 1
            adoRs.MoveNext
        Wend
        adoRs.Close
    End With
    
ErrExit:
    Set adoRs = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
errHandler:
    LblErrMsg.Caption = "[" & err.number & "] " & err.Description
    err.clear
    Resume ErrExit
End Sub



Sub adtocboDept()
Dim sqldept As String
Dim rsdept As New Recordset

    sqldept = "select * from Department_Cls order by Department_cls"
    Set rsdept = Db.Execute(sqldept)
    
    With cboDept
        .clear
        .columnCount = 2
        .ColumnWidths = "50pt;150pt"
        .ListWidth = 200
        .ListRows = 15
        
        i = 0
        Do While Not rsdept.EOF
            .AddItem
            .List(i, 0) = Trim(rsdept("Department_cls"))
            .List(i, 1) = IIf(IsNull(rsdept("Description")), "", Trim(rsdept("Description")))
            rsdept.MoveNext
            i = i + 1
        Loop
    End With
    Set rsdept = Nothing
End Sub

Sub adtocborequestno()
Dim sqlno As String
Dim rsno As New Recordset
    
    sqlno = "select PORequest_no from PORequest_Master " & _
            "where PORequest_date >= '" & Format(requestdate1.Value, "yyyy-mm-dd") & "' " & _
            "and PORequest_date <= '" & Format(requestdate2.Value, "yyyy-mm-dd") & "' " & _
            "and PersonInCharge_Cls = '" & Trim(cboPerson.Text) & "' and others_cls = '1' " & _
            "order by PORequest_date desc, PORequest_No desc "
    Set rsno = Db.Execute(sqlno)

    With cborequestno
        .clear
        Do While Not rsno.EOF
            .AddItem Trim(rsno("PORequest_No"))
            rsno.MoveNext
        Loop
        .ColumnWidths = "100pt"
        .ListWidth = 100
        .ListRows = 15
    End With
    Set rsno = Nothing
End Sub

Sub adtocboitem(ByVal nmCombo, ByVal field1 As String, ByVal field2 As String, ByVal colWidth1 As Integer, ByVal colwidth2 As Integer, ByVal Orderby As String)
Dim sqlitem As String
Dim RsItem As New Recordset

    sqlitem = "select accounting_code, unit_Cls, " & field1 & ", " & field2 & " from item_master " & _
              "where stockcontrol_cls = '02' and makebuy_cls = '02' order by " & Orderby
    Set RsItem = Db.Execute(sqlitem)
    
    With nmCombo
        .clear
        .columnCount = 4
        .ColumnWidths = colWidth1 & "pt;" & colwidth2 & "pt;0pt;80pt"
        .ListWidth = colWidth1 + colwidth2 + 80
        .ListRows = 15
        
        i = 0
        Do While Not RsItem.EOF
            .AddItem
            .List(i, 0) = IIf(IsNull(RsItem(field1)), "", Trim(RsItem(field1)))
            .List(i, 1) = IIf(IsNull(RsItem(field2)), "", Trim(RsItem(field2)))
            .List(i, 2) = IIf(IsNull(RsItem("unit_cls")), "", Trim(RsItem("unit_Cls")))
            .List(i, 3) = IIf(IsNull(RsItem("accounting_code")), "", Trim(RsItem("accounting_code")))
            RsItem.MoveNext
            i = i + 1
        Loop
    End With
    Set RsItem = Nothing
End Sub

Sub adtocboitemOthers()
Dim sqlitem As String
Dim RsItem As New Recordset

    sqlitem = "select * from OthersItem_Master order by item_desc"
    Set RsItem = Db.Execute(sqlitem)
    
    With cboItemName
        .clear
        .columnCount = 2
        .ColumnWidths = "250pt;130pt"
        .ListWidth = 380
        .ListRows = 15
        i = 0
        Do While Not RsItem.EOF
            .AddItem ""
            .List(i, 0) = Trim(RsItem("item_desc"))
            .List(i, 1) = Trim(RsItem("accounting_code") & "")
            RsItem.MoveNext
            i = i + 1
        Loop
    End With
    Set RsItem = Nothing
End Sub

Sub adtocboClass()
Dim sqlitem As String
Dim RsItem As New Recordset

    sqlitem = "select * FROM PO_Cls ORDER BY PO_CLS"
    Set RsItem = Db.Execute(sqlitem)
    
    With cboClass
        .clear
        .columnCount = 2
        .ColumnWidths = "50pt;150pt"
        .ListWidth = 200
        .ListRows = 5
        i = 0
        Do While Not RsItem.EOF
            .AddItem ""
            .List(i, 0) = Trim(RsItem("PO_Cls"))
            .List(i, 1) = Trim(RsItem("Description") & "")
            RsItem.MoveNext
            i = i + 1
        Loop
    End With
    Set RsItem = Nothing
    End Sub

Sub requestno(ByVal thn As String, ByVal bln As String)
Dim sqlno As String, SqlS As String
Dim rsno As New Recordset, rsS As New Recordset
    'PRYYMM99999
'    If Format(RequestDate, "YYYY-MM-01") > "2006-07-30" Then
'        sqlno = "select top 1 rtrim(PORequest_No) from PORequest_Master " & _
'                "where substring(rtrim(PORequest_No),3,2) = '" & thn & "' and substring(rtrim(PORequest_No),5,2) > '" & bln & "'  " & _
'                "order by right(rtrim(PORequest_No),5) desc"
'    Else
        sqlno = "select top 1 rtrim(PORequest_No) from PORequest_Master " & _
                "where substring(rtrim(PORequest_No),3,2) = '" & thn & "' " & _
                "order by right(rtrim(PORequest_No),5) desc"
'    End If
    Set rsno = Db.Execute(sqlno)
    If Not (rsno.BOF And rsno.EOF) Then
        txtRequestNo.Text = Left(Trim(rsno(0)), 4) & bln & Format(Right(Trim(rsno(0)), 5) + 1, "0000#")
    Else
        txtRequestNo.Text = "PR" & thn & bln & "00001"
    End If
    txtRequestNo.locked = True
    Set rsno = Nothing
End Sub

Sub kunci(l As Boolean)
    Period.Enabled = Not l
    RequestDate.Enabled = Not l
    cboDept.Enabled = Not l
    cboSection.Enabled = Not l
    Command1(1).Enabled = Not l
    lblFix.Caption = "Status Fix"
    lblFix.Visible = l
End Sub

Sub Header()
    With grid
        .clear
        .Rows = 1
        .ColS = 12
    
        .ColWidth(0) = 300
        .ColWidth(1) = 2500 'item code
        .ColWidth(2) = 3000
        .ColWidth(3) = 0    'class cls
        .ColWidth(4) = 1500
        .ColWidth(5) = 1700
        .ColWidth(6) = 1400
        .ColWidth(7) = 0    'unit cls
        .ColWidth(8) = 600
        .ColWidth(9) = 2100
        .ColWidth(11) = 1500
        
        .ColHidden(10) = True    'seq no
        
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "Product Code"
        .TextMatrix(0, 2) = "Product Name"
        .TextMatrix(0, 4) = "Class"
        .TextMatrix(0, 5) = "Req Delivery Date"
        .ColDataType(5) = flexDTDate
        .TextMatrix(0, 6) = "Request Qty"
        .TextMatrix(0, 8) = "Unit"
        .TextMatrix(0, 9) = "Remarks/Purpose"
        .TextMatrix(0, 11) = "Account No."
    
        .Cell(flexcpAlignment, 0, 0, 0, 11) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(4) = flexAlignLeftCenter
        .ColAlignment(5) = flexAlignCenterCenter
        .ColAlignment(6) = flexAlignRightCenter
        .ColAlignment(8) = flexAlignCenterCenter
        .ColAlignment(9) = flexAlignLeftCenter
        .ColAlignment(11) = flexAlignCenterCenter
        .EditMaxLength = 1
    End With
End Sub

Sub Browse()
    LblErrMsg = ""
    sql = "select * from PORequest_Master " & _
          "where porequest_no = '" & txtRequestNo.Text & "' and others_Cls = '1'"
    If RS.State <> adStateClosed Then RS.Close
    RS.Open sql, Db, adOpenKeyset, adLockOptimistic
    
    If Not (RS.BOF And RS.EOF) Then
        ada = True: ubah = True
        statusfix = IIf(IsNull(RS("fix_cls")), 0, RS("fix_cls"))
        Call BrowseGrid
        If statusfix = 1 Then Call kunci(True) Else Call kunci(False)
    Else
        ada = False
    End If
End Sub
Function Get_Field(sql, Field)
Dim Rdata As New ADODB.Recordset
Set Rdata = Db.Execute(sql)
Get_Field = ""
If Not Rdata.EOF Then
 Get_Field = IIf(IsNull(Rdata.Fields(Field)), "", Rdata.Fields(Field))
End If
End Function
Sub BrowseGrid()
Dim j As Integer
Dim SClas As New ADODB.Recordset
    Call Header
    Call kosongBwh
    
    sqlGrid = " select (select description from unit_cls uc where uc.unit_cls= PORequest_Detail.unit_cls ) unit_desc ," & _
              " * from PORequest_Detail where porequest_no = '" & txtRequestNo.Text & "' order by item_name, reqdelivery_date"
    If rsGrid.State <> adStateClosed Then rsGrid.Close
    rsGrid.Open sqlGrid, Db, adOpenKeyset, adLockOptimistic
    
    i = 1
    With grid
    Do While Not rsGrid.EOF
        .Rows = .Rows + 1
        .Cell(flexcpBackColor, i, 0) = &HFFFFFF
        .TextMatrix(i, 1) = Trim(rsGrid("Item_Code"))
        .TextMatrix(i, 2) = IIf(IsNull(rsGrid("item_name")), "", Trim(rsGrid("Item_name")))
        If IsNull(rsGrid("Class")) Then
            .TextMatrix(i, 3) = ""
            .TextMatrix(i, 4) = ""
        Else
        
                    .TextMatrix(i, 3) = Trim(rsGrid("class"))
                    .TextMatrix(i, 4) = Get_Field("SELECT * FROM PO_Cls WHERE PO_CLs='" & .TextMatrix(i, 3) & "'", 1)
                    
        End If
        .TextMatrix(i, 5) = Format(Trim(rsGrid("ReqDelivery_date")), "dd MMM yyyy")
        .TextMatrix(i, 6) = IIf(IsNull(rsGrid("Qty")), 0, Format(Trim(rsGrid("Qty")), "##,##0.#0"))
        If IsNull(rsGrid("unit_cls")) Then
          .TextMatrix(i, 7) = " "
          .TextMatrix(i, 8) = " "
        Else
          .TextMatrix(i, 7) = Trim(rsGrid("Unit_cls"))
          '.TextMatrix(i, 8) = Split(isiunit, ",")(Val(Trim(rsGrid("Unit_Cls"))) - 1)
          .TextMatrix(i, 8) = Trim(rsGrid("Unit_desc"))
        End If
        .TextMatrix(i, 9) = IIf(IsNull(rsGrid("purpose")), "", Trim(rsGrid("purpose")))
        .TextMatrix(i, 10) = rsGrid("poreq_seqno")
        .TextMatrix(i, 11) = IIf(IsNull(rsGrid("accountno")), "", Trim(rsGrid("accountno")))
        rsGrid.MoveNext
        i = i + 1
    Loop
    End With
End Sub

Sub BrowseAtas()
Dim p As String

    sql = "select * from PORequest_Master where PORequest_No = '" & txtRequestNo.Text & "' and Others_Cls = '1'"
    If RS.State <> adStateClosed Then RS.Close
    RS.Open sql, Db, adOpenKeyset, adLockOptimistic
    If Not (RS.BOF And RS.EOF) Then
        RequestDate.Value = IIf(IsNull(RS("porequest_date")), " ", Format(Trim(RS("porequest_date")), "dd MMM yyyy"))
        p = IIf(IsNull(RS("porequest_period")), " ", Left(Trim(RS("porequest_period")), 4) & "-" & Right(Trim(RS("porequest_period")), 2) & "-01")
        Period.Value = Format(p, "MMM yyyy")
        temptgl = Period.Month
        cboDept.Text = IIf(IsNull(RS("Department_Cls")), "", Trim(RS("Department_Cls")))
        statusfix = IIf(IsNull(RS("fix_cls")), 0, RS("fix_cls"))
        cboSection = IIf(IsNull(RS("Section_Cls")), "", Trim(RS("Section_Cls")))
        If statusfix = 1 Then Call kunci(True) Else Call kunci(False)
    End If
End Sub

Function seqNo() As Long
Dim sqlseqno As String
Dim rsseqno As New Recordset

    sqlseqno = "select POReq_seqno from PORequest_Detail order by POReq_seqno desc"
    If rsseqno.State <> adStateClosed Then rsseqno.Close
    rsseqno.Open sqlseqno, Db, adOpenKeyset, adLockOptimistic
    
    If Not (rsseqno.BOF And rsseqno.EOF) Then
        seqNo = rsseqno!POReq_seqno + 1
    Else
        seqNo = 1
    End If
    Set rsseqno = Nothing
End Function

Function cekpoqty(ByVal seqNo As String, ByVal requestno As String) As Double
Dim sqlcekpoqty As String
Dim rscekpoqty As New Recordset
    
    cekpoqty = 0
    sqlcekpoqty = "select pod.poreq_seqno, isnull(sum(pod.qty),0) poqty " & _
                  "from PurchaseOrder_Detail pod " & _
                  "where pod.porequest_no = '" & Trim(requestno) & "' " & _
                  "and pod.poreq_seqno='" & IIf(Trim(seqNo) = "", 0, Trim(seqNo)) & "' " & _
                  "group by pod.poreq_seqno"
    If rscekpoqty.State <> adStateClosed Then rscekpoqty.Close
    rscekpoqty.Open sqlcekpoqty, Db, adOpenKeyset, adLockOptimistic
    If Not (rscekpoqty.BOF And rscekpoqty.EOF) Then _
        cekpoqty = CDbl(rscekpoqty("poqty"))
    
    Set rscekpoqty = Nothing
End Function

Private Sub cboItemName_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub cboSection_Change()
lblSection.Text = ""
End Sub

Private Sub cboSection_Click()
 If cboSection.ListIndex <> -1 Then lblSection.Text = cboSection.Column(1)
End Sub

Private Sub cboSection_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 13 Then Call cboSection_Click
End Sub

Private Sub Form_Load()
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
        
    'ISI COMBO
    combo1.AddItem "Create"
    combo1.AddItem "Update"
    Call adtocboperson
    Call adtocboDept
    Call up_FillCombo(cbounit, "unit_Cls")
    Call addCboSection
    
    'Call isiCboUnitCurr(cbounit, isiunit, 0, 10)
    cbounit.TextColumn = 2
    Call adtocboClass
   
    cboCls.AddItem "By Code"
    cboCls.AddItem "Non Code"
    
    Call Kosong
    combo1.ListIndex = 1
End Sub

Private Sub cboperson_Change()
    lblPerson.Text = ""
End Sub

Private Sub cboperson_Click()
Dim ketemu As Boolean

    If cboPerson.ListIndex <> -1 Then lblPerson.Text = cboPerson.Column(1)
    
    If combo1.ListIndex = 1 Then
        Call adtocborequestno
        For i = 0 To cborequestno.ListCount - 1
            If txtRequestNo.Text = cborequestno.List(i) Then
                ketemu = True
                cborequestno.ListIndex = i
                Exit For
            End If
        Next i
        If ketemu = False Then txtRequestNo.Text = "": Call Header
    End If
End Sub

Private Sub cboperson_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call cboperson_Click
End Sub

Private Sub cboDept_Change()
    lblDept.Text = ""
End Sub

Private Sub cboDept_Click()
    If cboDept.ListIndex <> -1 Then lblDept.Text = cboDept.Column(1)
End Sub

Private Sub cboDept_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call cboDept_Click
End Sub

Private Sub period_Change()
    Call period_Click
    temptgl = Period.Month
    If combo1.ListIndex = 1 Then Call Header
End Sub

Private Sub period_Click()
    If Period.Month = 1 And Val(temptgl) = 12 Then Period.Year = Period.Year + 1
    If Period.Month = 12 And Val(temptgl) = 1 Then Period.Year = Period.Year - 1
End Sub

Private Sub requestdate_Change()
Dim t As String
    If combo1.ListIndex = 0 Then
        t = Right(Year(RequestDate), 2) & "-" & Format(Month(RequestDate), "0#")
        Call requestno(Right(Year(RequestDate), 2), Format(Month(RequestDate), "0#"))
    End If
End Sub

Private Sub requestdate1_Change()
Dim ketemu As Boolean
    
    LblErrMsg.Caption = ""
    If Format(requestdate1, "yyyy-mm-dd") > Format(requestdate2, "yyyy-mm-dd") Then
       LblErrMsg.Caption = DisplayMsg(4025) & " " & Format(requestdate2, "MMM yyyy")    '"Start Date must be lower than "
       Exit Sub
    End If

    If combo1.ListIndex = 1 Then
        Call adtocborequestno
        For i = 0 To cborequestno.ListCount - 1
            If txtRequestNo.Text = cborequestno.List(i) Then
                ketemu = True
                cborequestno.ListIndex = i
                Exit For
            End If
        Next i
        If ketemu = False Then txtRequestNo.Text = "": Call Header
    End If
End Sub

Private Sub requestdate2_Change()
Dim ketemu As Boolean

    LblErrMsg.Caption = ""
    If Format(requestdate2, "yyyy-mm-01") < Format(requestdate1, "yyyy-mm-01") Then
       LblErrMsg.Caption = DisplayMsg(4024) & " " & Format(requestdate1, "MMM yyyy")    '"End Date must be higher than "
       Exit Sub
    End If

    If combo1.ListIndex = 1 Then
        Call adtocborequestno
        For i = 0 To cborequestno.ListCount - 1
            If txtRequestNo.Text = cborequestno.List(i) Then
                ketemu = True
                cborequestno.ListIndex = i
                Exit For
            End If
        Next i
        If ketemu = False Then txtRequestNo.Text = "": Call Header
    End If
End Sub

Private Sub Combo1_Click()
Dim ketemu As Boolean, t As String

    ketemu = False
    LblErrMsg = ""
    Call kunci(False)
    Call Header

    If combo1.ListIndex = 0 Then    'CREATE
        Command1(0).Caption = "&Create"
        ubah = False
        requestdate1.Enabled = False
        requestdate2.Enabled = False
        RequestDate.Value = Format(Now, "dd mmm yyyy")
        RequestDate.Enabled = False
        cborequestno.locked = True
        txtRequestNo.Text = ""
        t = Right(Year(RequestDate), 2) & "-" & Format(Month(RequestDate), "0#")
        Call requestno(Right(Year(RequestDate), 2), Format(Month(RequestDate), "0#"))
        Call kosongBwh
    Else    'UPDATE
        Call adtocborequestno
        Command1(0).Caption = "&Update"
        ubah = True
        cborequestno.locked = False
        txtRequestNo.locked = False
        requestdate1.Enabled = True
        requestdate2.Enabled = True
        RequestDate.Enabled = True

        For i = 0 To cborequestno.ListCount - 1
            If txtRequestNo.Text = cborequestno.List(i) Then
                ketemu = True
                cborequestno.ListIndex = i
                Exit For
            End If
        Next i
        If ketemu = False Then txtRequestNo.Text = ""
    End If
End Sub

Private Sub combo1_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call Combo1_Click
End Sub

Private Sub cborequestno_Click()
    LblErrMsg = ""
    txtRequestNo.Text = cborequestno.Text
    Call Header
    Call BrowseAtas
End Sub

Private Sub cborequestno_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call cborequestno_Click
End Sub

Private Sub txtrequestno_Change()
Dim ketemu As Boolean
    
    LblErrMsg = ""
    If combo1.ListIndex = 1 Then
        For i = 0 To cborequestno.ListCount - 1
            If txtRequestNo.Text = cborequestno.List(i) Then
                ketemu = True
                cborequestno.ListIndex = i
                Exit For
            End If
        Next i
        If ketemu = False Then cborequestno.ListIndex = -1
    End If
End Sub

Private Sub txtrequestno_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = Asc("'") Then KeyAscii = 0
    
    If KeyAscii = 13 Then
        If combo1.ListIndex = 0 Then
            SendKeys vbTab
        Else
            Call Header
            Call BrowseAtas
        End If
    End If
End Sub

Private Sub cboCls_Click()
    If UCase(Trim(cboCls.Text)) = "BY CODE" Then
        CboItemCode.Enabled = True
        Call adtocboitem(CboItemCode, "item_code", "item_name", 130, 200, "item_code")
        Call adtocboitem(cboItemName, "item_name", "item_code", 200, 130, "item_name")
        cbounit.ListIndex = -1
        cbounit.Enabled = False
        
    ElseIf UCase(Trim(cboCls.Text)) = "NON CODE" Then
        Call adtocboitemOthers
        CboItemCode.clear
        CboItemCode.Text = ""
        CboItemCode.Enabled = False
        cbounit.ListIndex = -1
        cbounit.Enabled = True
        
    End If
End Sub

Private Sub CboItemCode_Change()
    LblErrMsg = ""
'    cboItemName.Text = ""
'    cbounit.ListIndex = -1
End Sub

Private Sub cboitemcode_Click()
    LblErrMsg = ""
    If CboItemCode.ListIndex <> -1 Then
        cboItemName.Text = CboItemCode.Column(1)
        For i = 0 To cbounit.ListCount - 1
            If Trim(cbounit.List(i, 0)) = Trim(CboItemCode.Column(2)) Then
                cbounit.ListIndex = i
                Exit For
            End If
        Next i
    
    End If
End Sub

Private Sub cboitemcode_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call cboitemcode_Click
End Sub

Private Sub cboItemCode_LostFocus()
    Call cboitemcode_Click
End Sub

Private Sub CboItemName_Change()
    LblErrMsg = ""
'    cboItemCode.Text = ""
'    cbounit.ListIndex = -1
End Sub

Private Sub cboitemname_Click()
    LblErrMsg = ""
    If UCase(Trim(cboCls.Text)) = "BY CODE" Then
        If cboItemName.ListIndex <> -1 Then
            CboItemCode.Text = cboItemName.Column(1)
            For i = 0 To cbounit.ListCount - 1
                If Trim(cbounit.List(i, 0)) = Trim(CboItemCode.Column(2)) Then
                    cbounit.ListIndex = i
                    Exit For
                End If
            Next i
            
        End If
    Else
        'If cboItemName.ListIndex <> -1 Then CboAccount.Text = cboItemName.Column(1)
    End If
End Sub

Private Sub cboitemname_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call cboitemname_Click
End Sub

Private Sub cboItemname_LostFocus()
    Call cboitemname_Click
End Sub

Private Sub txtqty_Change()
    If InStr(1, txtQty.Text, ",") = 1 Then txtQty.Text = Right(txtQty, Len(txtQty) - 1)
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then _
          KeyAscii = 0
    If InStr(1, txtQty.Text, ".") > 0 Then If KeyAscii = Asc(".") Then KeyAscii = 0
    If (txtQty.Text & Chr(KeyAscii)) > 9999999.99 And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub txtQty_LostFocus()
Dim z As Double
    If IsNumeric(txtQty.Text) = False Then txtQty.Text = 0
    If txtQty.Text <> "" Then
        z = CDbl(txtQty.Text)
        If z > 9999999.99 Then txtQty.Text = Left(z, 7)
    End If
    txtQty.Text = Format(txtQty.Text, "#,##0.#0")
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim TextGrid As String
Dim sql1 As String, rs1 As New Recordset

With grid
    TextGrid = grid.Text
    If TextGrid = "S" Then
'        txtItemCd = .TextMatrix(Row, 1)
        If Trim(.TextMatrix(Row, 1)) = "" Then cboCls.ListIndex = 1 Else cboCls.ListIndex = 0
        CboItemCode.Text = .TextMatrix(Row, 1)
        cboItemName.Text = .TextMatrix(Row, 2)
        sql1 = "select * from PurchaseOrder_Detail where PORequest_No = '" & txtRequestNo.Text & "' " & _
               "and POReq_Seqno = '" & .TextMatrix(Row, 10) & "' "
        Set rs1 = Db.Execute(sql1)
        If Not (rs1.BOF And rs1.EOF) Then
'            txtItemCd.Enabled = False
            cboCls.Enabled = False
            CboItemCode.Enabled = False
            cboItemName.Enabled = False
        Else
'            txtItemCd.Enabled = True
            cboCls.Enabled = True
            If cboCls.ListIndex = 0 Then CboItemCode.Enabled = True
            cboItemName.Enabled = True
        End If
        Set rs1 = Nothing
        
        txtDesc.Text = .TextMatrix(Row, 2)
        cboClass.ListIndex = -1
        For i = 0 To cboClass.ListCount - 1
            If .TextMatrix(Row, 3) = cboClass.List(i, 0) Then
                cboClass.ListIndex = i
                Exit For
            End If
        Next i
        DelDate = Format(.TextMatrix(Row, 5), "dd mmm yyyy")
        txtQty.Text = Format(.TextMatrix(Row, 6), "#,##0.#0")
        lblQty = CDbl(.TextMatrix(Row, 6))
        cbounit.ListIndex = -1
        For i = 0 To cbounit.ListCount - 1
            If .TextMatrix(Row, 7) = cbounit.List(i, 0) Then
                cbounit.ListIndex = i
                Exit For
            End If
        Next i
        txtPurpose.Text = .TextMatrix(Row, 9)
'        CboAccount.Text = .TextMatrix(Row, 11)
        lblseqno = .TextMatrix(Row, 10)
       Call kosongColGrid
    ElseIf TextGrid = "D" Then
       Call kosongColGrid("S")
    End If
    .TextMatrix(Row, Col) = TextGrid
End With
End Sub

Private Sub kosongColGrid(Optional Kolom As String)
    With grid
        .Col = 0
        If Kolom <> "" Then
           For i = 1 To .Rows - 1
              If .Text = Kolom Then .Text = ""
              If .TextMatrix(i, 0) <> "D" Then .TextMatrix(i, 0) = ""
           Next i
           kosongBwh
        Else
           For i = 1 To .Rows - 1
              If .TextMatrix(i, 0) <> "" Then .TextMatrix(i, 0) = ""
           Next i
        End If
    End With
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If grid.Col <> 0 Then Cancel = True
End Sub

Private Sub Grid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col = 0 Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii <> Asc("S") And KeyAscii <> Asc("D") And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then _
            KeyAscii = 0
        If KeyAscii = Asc(".") Then KeyAscii = 0
    End If
End Sub


Private Sub Command1_Click(Index As Integer)
Dim sql1 As String, rs1 As New Recordset
Dim sql2 As String, rs2 As New Recordset
Dim t As String, tanya, hapus As Boolean
Dim ubahgrid As Boolean
MousePointer = vbHourglass
    ubahgrid = False
    LblErrMsg = ""
    
    Select Case Index
    Case 0: 'CREATE / UPDATE
            If hakUpdate(Me.Name) = 0 Then _
                LblErrMsg = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
    
            'HEADER VALIDATION
            If cboPerson.Text = "" Then
                LblErrMsg = DisplayMsg(1070)
                cboPerson.SetFocus
                Exit Sub
            End If
            If cboPerson.Text <> "" Then
                If cboPerson.MatchFound = False Then
                    LblErrMsg = DisplayMsg(4086)    'Record with this Person in Charge not found
                    cboPerson.SetFocus
                    Exit Sub
                End If
            End If
            If cboDept.Text <> "" Then
                If cboDept.MatchFound = False Then
                    LblErrMsg = DisplayMsg(4142)    'Record with this Department not found
                    cboDept.SetFocus
                    Exit Sub
                End If
            End If
            If txtRequestNo.Text = "" Then
                LblErrMsg = DisplayMsg(8121) '"Please Input Request No"
                txtRequestNo.SetFocus
                Exit Sub
            End If
    
            If combo1.ListIndex = 0 Then    'CREATE
                If ubah = False Then
                    sql = "select * from PORequest_Master where porequest_no = '" & txtRequestNo.Text & "' "
                    If RS.State <> adStateClosed Then RS.Close
                    RS.Open sql, Db, adOpenKeyset, adLockOptimistic
                    If Not (RS.BOF And RS.EOF) Then
                        LblErrMsg.Caption = DisplayMsg(1023)
                        txtRequestNo.SetFocus
                        Exit Sub
                    Else
                        RS.AddNew
                        RS("PORequest_No") = txtRequestNo.Text
                    End If
                End If
                RS("PORequest_Period") = Year(Period.Value) & Format(Month(Period.Value), "0#")
                RS("PORequest_Date") = Format(RequestDate.Value, "yyyy-mm-dd")
                RS("PersonInCharge_Cls") = Trim(cboPerson.Text)
                RS("Department_Cls") = Trim(cboDept.Text)
                RS("Others_Cls") = "1"
                RS("Username") = userLogin
                RS("Section_Cls") = Trim(cboSection.Text)
                RS("Last_Update") = Format(Now, "yyyy-mm-dd hh:mm:ss")


'On Error Resume Next
                RS.update
errHandler:
                If InStr(1, err.Description, "Violation of PRIMARY KEY constraint") > 0 Then
                    t = Right(Year(RequestDate), 2) & "-" & Format(Month(RequestDate), "0#")
                    Call requestno(Right(Year(RequestDate), 2), Format(Month(RequestDate), "0#"))
                    RS("porequest_No") = txtRequestNo.Text
                    RS.update
                    If InStr(1, err.Description, "Violation of PRIMARY KEY constraint") > 0 Then
                        GoTo errHandler
                    Else
                        If Trim$(err.Description) <> "" Then
                            LblErrMsg = Trim$(err.number) + " : " + Trim$(err.Description)
                            Exit Sub
                        End If
                    End If
                Else
                    If Trim$(err.Description) <> "" Then
                        LblErrMsg = Trim$(err.number) + " : " + Trim$(err.Description)
                        Exit Sub
                    End If
                End If
                
                If CDate(RequestDate.Value) > CDate(requestdate1.Value) Then
                    If CDate(RequestDate.Value) > CDate(requestdate2.Value) Then _
                        requestdate2.Value = Format(RequestDate.Value, "dd MMM yyyy")
                Else
                    requestdate1.Value = Format(RequestDate.Value, "dd MMM yyyy")
                End If
    
                combo1.Text = "Update"
                MousePointer = vbDefault
                LblErrMsg.Caption = DisplayMsg(1000)
                ubah = True
            
            Else    'UPDATE
                Call Browse
                If ada = False Then
                    LblErrMsg.Caption = DisplayMsg(4144)    'Record with this Request No not found
                    txtRequestNo.SetFocus
                    Exit Sub
                End If
            End If

    Case 1: 'SUBMIT
            If hakUpdate(Me.Name) = 0 Then _
                LblErrMsg = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
            
            'HEADER VALIDATION
            If cboPerson.Text = "" Then
                LblErrMsg = DisplayMsg(8119)
                cboPerson.SetFocus
                Exit Sub
            End If
            If cboPerson.Text <> "" Then
                If cboPerson.MatchFound = False Then
                    LblErrMsg = DisplayMsg(4086)    'Record with this Person in Charge not found
                    cboPerson.SetFocus
                    Exit Sub
                End If
            End If
            If cboDept.Text <> "" Then
                If cboDept.MatchFound = False Then
                    LblErrMsg = DisplayMsg(4142)    'Record with this Department not found
                    cboDept.SetFocus
                    Exit Sub
                End If
            End If
            If txtRequestNo.Text = "" Then
                LblErrMsg = DisplayMsg(1067) '"Please Input Request No"
                txtRequestNo.SetFocus
                Exit Sub
            End If
              
            sql = "select * from PORequest_Master where PORequest_No = '" & txtRequestNo.Text & "' and Others_Cls = '1'"
            If RS.State <> adStateClosed Then RS.Close
            RS.Open sql, Db, adOpenKeyset, adLockOptimistic
            If RS.BOF And RS.EOF Then
                LblErrMsg.Caption = DisplayMsg(8122)
                txtRequestNo.SetFocus
                Exit Sub
            End If

            If ubah = True Then
                RS("PORequest_Period") = Year(Period.Value) & Format(Month(Period.Value), "0#")
                RS("PORequest_Date") = Format(RequestDate.Value, "yyyy-mm-dd")
                RS("PersonInCharge_Cls") = Trim(cboPerson.Text)
                RS("Department_Cls") = Trim(cboDept.Text)
                RS("Others_Cls") = "1"
                RS("Section_Cls") = Trim(cboSection)
                RS("Username") = userLogin
                RS("Last_Update") = Format(Now, "yyyy-mm-dd hh:mm:ss")
                RS.update
                
                With grid   'DELETE GRID
                    For i = 1 To .Rows - 1
                        If .TextMatrix(i, 0) = "D" Then
                            If IsEmpty(tanya) Then tanya = MsgBox("Do You Really Want To Delete This Data ?", vbQuestion & vbYesNo, "Confirmation")
                            If tanya = vbYes Then
                                sql1 = "select * from PurchaseOrder_Detail where PORequest_No = '" & txtRequestNo.Text & "' " & _
                                       "and POReq_SeqNo = '" & .TextMatrix(i, 10) & "' "
                                Set rs1 = Db.Execute(sql1)
                                If Not (rs1.BOF And rs1.EOF) Then
                                    LblErrMsg.Caption = DisplayMsg(1204)
                                    .Row = i
                                    .SetFocus
                                    Exit Sub
                                Else
                                    sql = "delete from PORequest_Detail where PORequest_No = '" & txtRequestNo.Text & "' " & _
                                          "and POReq_SeqNo = '" & .TextMatrix(i, 10) & "' "
                                    Db.Execute sql
                                    hapus = True
                                End If
                            Else
                                Exit For
                            End If
                        ElseIf .TextMatrix(i, 0) = "S" Then
                            ubahgrid = True
                        End If
                    Next i
                    If (hapus) Then Call Browse: LblErrMsg = DisplayMsg(1201): MousePointer = vbDefault: Exit Sub
                End With
                
                If CboItemCode.Text <> "" Or cboItemName.Text <> "" Or txtQty.Text <> "" Or cbounit.Text <> "" Or txtPurpose.Text <> "" Then
                    'DETAIL VALIDATION
'                    If txtDesc.Text = "" Then
'                        txtDesc.SetFocus
'                        lblErrMsg = DisplayMsg(1006) 'Please Input Description
'                        Exit Sub
                    If UCase(Trim(cboCls.Text)) = "BY CODE" Then
                        If Trim(CboItemCode.Text) = "" Then
                            If CboItemCode.Enabled = True Then CboItemCode.SetFocus
                            LblErrMsg = DisplayMsg(1024)    'Please select product code
                            Exit Sub
                        Else
                            If CboItemCode.MatchFound = False Then
                                If CboItemCode.Enabled = True Then CboItemCode.SetFocus
                                LblErrMsg = DisplayMsg(4003)
                                Exit Sub
                            End If
                        End If
                    End If
                    If Trim(cboItemName.Text) = "" Then
                        If cboItemName.Enabled = True Then cboItemName.SetFocus
                        LblErrMsg = DisplayMsg(1099)    'Please Input product name
                        Exit Sub
                    Else
                        If UCase(Trim(cboCls.Text)) = "BY CODE" Then
                        If cboItemName.MatchFound = False Then
                            If cboItemName.Enabled = True Then cboItemName.SetFocus
                            LblErrMsg = DisplayMsg(4165)
                            Exit Sub
                        End If
                        End If
                    End If
                    If cboClass.Text = "" Then
                        cboClass.SetFocus
                        LblErrMsg = "Please Select Classification"
                        Exit Sub
                    End If
                    If txtQty.Text = "" Then
                        txtQty.SetFocus
                        LblErrMsg = DisplayMsg(1012) 'Please Input Quantity
                        Exit Sub
                    ElseIf cbounit.Text = "" Then
                        If cbounit.Enabled = True Then cbounit.SetFocus
                        LblErrMsg = DisplayMsg(1030)
                        Exit Sub
                    End If
                      
                    If txtQty.Text = 0 Then
                        txtQty.SetFocus
                        LblErrMsg = DisplayMsg(1012) 'Please Input Quantity
                        Exit Sub
                    ElseIf CDbl(txtQty.Text) > 9999999.99 Then
                        txtQty.SetFocus
                        LblErrMsg = DisplayMsg(4045) & " 9,999,999.99" '"Quantity must be lower or equal than 9,999,999.99"
                        Exit Sub
                    End If
                    Dim poqty As Double
                    poqty = cekpoqty(lblseqno, txtRequestNo.Text)
                    If CDbl(txtQty.Text) < poqty Then
                        txtQty.SetFocus
                        LblErrMsg = DisplayMsg(4036) & " " & poqty '"Quantity must be higher or equal than "
                        Exit Sub
                    End If
                                                  
                    'INSERT PO REQUEST DETAIL
                    If ubahgrid = False Then
                        sqlGrid = "select * from PORequest_Detail " & _
                                  "where PORequest_No = '" & txtRequestNo.Text & "' " & _
                                  "and POReq_seqno = " & lblseqno & _
                                  " order by item_name"
                        If rsGrid.State <> adStateClosed Then rsGrid.Close
                        rsGrid.Open sqlGrid, Db, adOpenKeyset, adLockOptimistic
                        If rsGrid.BOF And rsGrid.EOF Then
                            rsGrid.AddNew
                            rsGrid("POReq_seqno") = seqNo
                            rsGrid("PORequest_No") = Trim(txtRequestNo.Text)
                        End If
                    Else    'UPDATE PO REQUEST DETAIL
                        sqlGrid = "select * from PORequest_Detail " & _
                                  "where PORequest_No = '" & txtRequestNo.Text & "' " & _
                                  "and POReq_SeqNo = " & lblseqno & _
                                  " order by item_name"
                        If rsGrid.State <> adStateClosed Then rsGrid.Close
                        rsGrid.Open sqlGrid, Db, adOpenKeyset, adLockOptimistic
                    End If
                                                
'                    rsGrid("Item_Code") = Trim(txtItemCd.Text)
                    rsGrid("Item_Code") = Trim(CboItemCode.Text)
                    rsGrid("item_name") = Trim(cboItemName.Text)
                    rsGrid("Class") = Trim(cboClass.Text)
                    rsGrid("ReqDelivery_Date") = Format(DelDate, "yyyy-mm-dd")
                    rsGrid("qty") = CDbl(txtQty.Text)
                    rsGrid("unit_cls") = cbounit.Column(0)
                    rsGrid("Purpose") = Trim(txtPurpose.Text)
                    
                    'rsGrid("AccountNo") = Trim(CboAccount.Text)
                    rsGrid("username") = userLogin
                    rsGrid("last_update") = Format(Now, "yyyy-mm-dd hh:mm:ss")
                    rsGrid.update
                    
                    'ADD ITEM NAME TO OTHERSITEM_MASTER
                    If UCase(Trim(cboCls.Text)) = "NON CODE" Then
                        sql2 = "select * From OthersItem_Master where Item_Desc = '" & Trim(cboItemName.Text) & "' "
                        Set rs2 = Db.Execute(sql2)
                        If rs2.BOF And rs2.EOF Then
                            sql2 = "insert into OthersItem_Master (Item_Desc) values ('" & Trim(cboItemName.Text) & "') "
                            Db.Execute sql2
                            Call adtocboitemOthers
                        End If
                        Set rs2 = Nothing
'                        If cboItemName.List(cboItemName.ListIndex, 1) <> Trim(CboAccount.Text) Then
'                            sql2 = "update othersitem_master set accounting_code = '" & Trim(CboAccount.Text) & "' " & _
'                                "where item_desc = '" & Trim(cboItemName.Text) & "'"
'                            Db.Execute sql2
'                            cboItemName.List(cboItemName.ListIndex, 1) = Trim(CboAccount.Text)
'                        End If
                    Else
'                        If cboItemCode.List(cboItemCode.ListIndex, 3) <> Trim(CboAccount.Text) Then
'                            sql2 = "update item_master set accounting_code = '" & Trim(CboAccount.Text) & "' " & _
'                                "where item_code = '" & Trim(cboItemCode.Text) & "'"
'                            Db.Execute sql2
'                            cboItemCode.List(cboItemCode.ListIndex, 3) = Trim(CboAccount.Text)
'                            cboItemName.List(cboItemName.ListIndex, 3) = Trim(CboAccount.Text)
'                        End If
                    End If
                    
                    Call Browse
                    ubahgrid = True
                End If
                MousePointer = vbDefault
                LblErrMsg = DisplayMsg(1101)
            End If

    Case 2: 'CLEAR
            Call Kosong
            combo1.ListIndex = 1

    Case 3: 'CANCEL
            If Trim(txtRequestNo.Text) <> "" Then
                If cboPerson.Text = "" Then
                    LblErrMsg = DisplayMsg(1070)
                    cboPerson.SetFocus
                    Exit Sub
                End If
                If cboPerson.Text <> "" Then
                    If cboPerson.MatchFound = False Then
                        LblErrMsg = DisplayMsg(4086)    'Record with this Person in Charge not found
                        cboPerson.SetFocus
                        Exit Sub
                    End If
                End If
                If cboDept.Text <> "" Then
                    If cboDept.MatchFound = False Then
                        LblErrMsg = DisplayMsg(4142)    'Record with this Department not found
                        cboDept.SetFocus
                        Exit Sub
                    End If
                End If
                
                Call BrowseAtas
                Call Browse
            End If
    End Select
MousePointer = vbDefault
End Sub

Private Sub cmdReport_Click()
Dim application As New CRAXDDRT.application
Dim report As New CRAXDDRT.report
Dim rsRpt As New ADODB.Recordset
Dim Rpt As New FrmRpt3
Dim sqlcekdet As String, SqlRpt As String
Dim rscekdet As New Recordset
  
    If combo1.ListIndex = 1 And txtRequestNo.Text <> "" Then
        sqlcekdet = "select prd.porequest_no from PORequest_Detail prd " & _
                    "inner join PORequest_Master prm on prm.porequest_no = prd.porequest_no " & _
                    "where prd.porequest_no = '" & txtRequestNo.Text & "' "
                    
        Set rscekdet = Db.Execute(sqlcekdet)
        If Not (rscekdet.BOF And rscekdet.EOF) Then
            Me.MousePointer = vbHourglass
            
'            SqlRpt = "select rtrim(pm.porequest_no) porequest_no, pm.porequest_Date, pm.department_cls, rtrim(dc.description) Department, " & _
'                     "pd.POReq_SeqNo, rtrim(pd.item_code) item_code, rtrim(pd.item_name) item_name, rtrim(pd.class) class, isnull(pd.qty,0) qty, pd.unit_cls, pd.ReqDelivery_Date, " & _
'                     "rtrim(pd.Purpose) Purpose, rtrim(pd.AccountNo) Accounting_Code, " & _
'                     "(select upper(rtrim(company_name)) from Company_Profile) comp_name, " & _
'                     "ISNULL((select isnull(case when datediff(month, ClosingDate, StartDate) = 0  then sum(lm_inventory) " & _
'                     "             when datediff(month, ClosingDate, StartDate) = 1  then sum(tm_current) " & _
'                     "             when datediff(month, ClosingDate, StartDate) >= 2 then sum(nm_current) end,0) Stock " & _
'                     "    From ( select " & _
'                     "            (select cast( cast(year as varchar(4)) + case when month < 10 then '0' else '' end + cast(month as varchar(2)) + '01' as dateTime) ClosingDate " & _
'                     "        from (select top 1 max(inventory_month) month, inventory_year year from inventory_control " & _
'                     "            where fix_cls='1' group by inventory_year order by inventory_year desc ) tbI " & _
'                     "            ) ClosingDate, StartDate = '" & Format(Now, "yyyy-mm-dd") & "', * " & _
'                     "        From Stock_Master " & _
'                     "    ) tbS " & _
'                     "    Where tbS.Item_Code = pd.Item_Code group by ClosingDate, Item_code, StartDate " & _
'                     "),0) Current_Stock, '' Currency_Code, 0 Price, pm.Others_cls, rtrim((select isnull(PO_Person,'') from Company_Profile)) PO_Person " & _
'                     "from PORequest_Master pm " & _
'                     "inner join PORequest_Detail pd on pd.porequest_no = pm.porequest_no " & _
'                     "left join Department_Cls dc on dc.Department_Cls = pm.Department_Cls " & _
'                     "where pm.PORequest_No = '" & Trim(txtRequestNo.Text) & "' and pm.Others_cls = '1' " & _
'                     "order by pd.item_name, pd.reqdelivery_date "
            
            
        Dim rsclosing As ADODB.Recordset
        Dim CloseThn As Long
        Dim CloseBln As Long
        Dim selisih As Long
        Dim FPilih As String
        
        Set rsclosing = New ADODB.Recordset
        rsclosing.Open "select * from inventory_control " & _
                              " Where Inventory_Month=(Select Max(Inventory_Month) from Inventory_Control " & _
                              " Where inventory_Year=(Select Max(Inventory_Year) from Inventory_Control)) ", Db
    
        CloseThn = rsclosing(0)
        CloseBln = rsclosing(1)
            
        selisih = (Year(RequestDate) * 12 + Month(RequestDate)) - (CloseThn * 12 + CloseBln)
        
        If selisih < 0 And selisih > 2 Then
            LblErrMsg = "Invalid Date Periode !"
        Else
            If selisih = 0 Then
                FPilih = "LM_Current"
            ElseIf selisih = 1 Then
                FPilih = "TM_Current"
            Else
                FPilih = "NM_Current"
            End If
        End If

SqlRpt = " Select PRM.PoRequest_No, PRM.PoRequest_Date, " & _
            vbLf & " PRM.Department_Cls,D.Description Department, " & _
            vbLf & " PRM.Section_Cls,S.Description Section, " & _
            vbLf & " PRM.PersonInCharge_Cls,P.Description PIC, " & _
            vbLf & " PRD.Item_Code, ISNULL(IM.Item_Name,PRD.Item_Name),IM.WH_Code,IM.Supplier_Code,IM.Control_Cls, " & _
            vbLf & " TM.Trade_Name Supplier_Name, " & _
            vbLf & " isnull(SM." & FPilih & ",0) Stock, " & _
            vbLf & " PRD.Qty,PRD.ReqDelivery_Date,dateadd(month,1,PRD.ReqDelivery_Date) BlnF1,dateadd(month,2,PRD.ReqDelivery_Date) BlnF2, " & _
            vbLf & " isnull((Select ChildRequirement_Qty from requirement_Master Where ChildRequirement_Month=month(PRD.ReqDelivery_Date)+1 and ChildRequirement_Year=year(PRD.ReqDelivery_Date) and ChildItem_Code=PRD.Item_code),0) F1, " & _
            vbLf & " isnull((Select ChildRequirement_Qty from requirement_Master Where ChildRequirement_Month=month(PRD.ReqDelivery_Date)+2 and ChildRequirement_Year=year(PRD.ReqDelivery_Date) and ChildItem_Code=PRD.Item_code),0) F2 " & _
            vbLf & " From PoRequest_Master PRM inner Join PORequest_Detail PRD " & _
            vbLf & " On PRM.PoREquest_No=PRD.PoRequest_No " & _
            vbLf & " Left Join Item_Master IM on PRD.Item_Code=IM.Item_Code " & _
            vbLf & " Left Join Trade_Master TM on IM.Supplier_Code=TM.Trade_Code " & _
            vbLf & " Inner Join Department_Cls D on PRM.Department_Cls=D.Department_Cls " & _
            vbLf & " Inner Join Section_Cls S on PRM.Section_Cls=S.Section_Cls " & _
            vbLf & " Inner Join PersonInCharge_Cls P on PRM.PersonInCharge_Cls=P.PersonInCharge_Cls " & _
            vbLf & " Left Join Stock_Master SM on PRD.Item_Code=SM.Item_Code and IM.WH_Code=SM.WareHouse_Code " & _
            vbLf & " where PRM.PORequest_No = '" & Trim(txtRequestNo.Text) & "' and PRM.Others_cls = '1' " & _
            vbLf & " order by PRD.POReq_SeqNo,PRD.Item_Code, PRD.ReqDelivery_Date"

' -----
            
            If rsRpt.State <> adStateClosed Then rsRpt.Close
            rsRpt.Open SqlRpt, Db, adOpenDynamic, adLockOptimistic
            
            sqlprint = SqlRpt
            reportcode = "PORequestDirect"
            
            If rsRpt.EOF Then LblErrMsg.Caption = DisplayMsg(4006): Me.MousePointer = vbDefault: Exit Sub
            
            Set report = application.OpenReport(App.path & "\Reports\rptPORequestDirectNewGroup.rpt")
            report.Database.Tables(1).SetDataSource rsRpt
            
            Rpt.CRViewer1.ReportSource = report
            Rpt.CRViewer1.ViewReport
            Rpt.CRViewer1.Zoom 1
            Rpt.WindowState = 2
            Rpt.Show 1
            Me.MousePointer = vbDefault
        Else
            LblErrMsg.Caption = DisplayMsg(4006): Me.MousePointer = vbDefault
        End If
    End If

    Set rscekdet = Nothing
    Set rsRpt = Nothing
End Sub

Private Sub CmdSubMenu_Click()
    sql = "delete from PORequest_Master Where porequest_no = '" & Trim(txtRequestNo.Text) & "' " & _
          "And porequest_no not in (select porequest_no from PORequest_Detail) and others_cls = '1'"
    Db.Execute sql

    Unload Me
    frmMainMenu.Show
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
    If ErrMsg = "" Then
        Unload Me
    Else
        LblErrMsg.Caption = ErrMsg
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RS = Nothing
    Set rsGrid = Nothing
End Sub

