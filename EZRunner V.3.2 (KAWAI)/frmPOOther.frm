VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPOOther 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Purchase Order Unscheduled (Others)"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox TxtDisc 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5550
      MaxLength       =   25
      TabIndex        =   86
      TabStop         =   0   'False
      Top             =   8985
      Width           =   2355
   End
   Begin VB.TextBox TxtSubAmount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3150
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   85
      TabStop         =   0   'False
      Top             =   8985
      Width           =   2355
   End
   Begin VB.TextBox TxtService 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   12210
      MaxLength       =   19
      TabIndex        =   16
      Text            =   "9,999,999.99"
      Top             =   6030
      Width           =   1275
   End
   Begin VB.TextBox txtSubTotal 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   13590
      Locked          =   -1  'True
      MaxLength       =   19
      TabIndex        =   17
      TabStop         =   0   'False
      Text            =   "9,999,999.99"
      Top             =   6030
      Width           =   1365
   End
   Begin VB.TextBox txtGrandTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   12870
      Locked          =   -1  'True
      MaxLength       =   35
      TabIndex        =   79
      TabStop         =   0   'False
      Top             =   8970
      Width           =   2235
   End
   Begin VB.TextBox txtPONo2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   300
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   8970
      Width           =   2490
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7950
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   77
      TabStop         =   0   'False
      Top             =   8970
      Width           =   2355
   End
   Begin VB.TextBox txtPPn 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10410
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   76
      TabStop         =   0   'False
      Top             =   8970
      Width           =   2355
   End
   Begin VB.TextBox txtMarking 
      Height          =   315
      Index           =   2
      Left            =   10110
      MaxLength       =   25
      TabIndex        =   25
      Top             =   7095
      Width           =   2085
   End
   Begin VB.TextBox txtMarking 
      Height          =   315
      Index           =   0
      Left            =   7350
      MaxLength       =   25
      TabIndex        =   23
      Top             =   7095
      Width           =   2085
   End
   Begin VB.TextBox txtMarking 
      Height          =   315
      Index           =   4
      Left            =   12930
      MaxLength       =   25
      TabIndex        =   27
      Top             =   7080
      Width           =   2085
   End
   Begin VB.TextBox txtRemarks 
      Height          =   540
      Left            =   8040
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   29
      Top             =   7950
      Width           =   6990
   End
   Begin VB.TextBox txtMarking 
      Height          =   315
      Index           =   3
      Left            =   10110
      MaxLength       =   25
      TabIndex        =   26
      Top             =   7470
      Width           =   2085
   End
   Begin VB.TextBox txtMarking 
      Height          =   315
      Index           =   1
      Left            =   7350
      MaxLength       =   25
      TabIndex        =   24
      Top             =   7470
      Width           =   2085
   End
   Begin VB.TextBox txtMarking 
      Height          =   315
      Index           =   5
      Left            =   12900
      MaxLength       =   25
      TabIndex        =   28
      Top             =   7485
      Width           =   2085
   End
   Begin VB.TextBox txtPriceCondition 
      BackColor       =   &H00FDDFE3&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3090
      Locked          =   -1  'True
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   7140
      Width           =   3570
   End
   Begin VB.TextBox txtPaymentTerm 
      BackColor       =   &H00FDDFE3&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3090
      Locked          =   -1  'True
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   6750
      Width           =   3585
   End
   Begin VB.TextBox txtTransport 
      BackColor       =   &H00FDDFE3&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3090
      Locked          =   -1  'True
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   8190
      Width           =   3585
   End
   Begin VB.TextBox txtInsurance 
      BackColor       =   &H00FDDFE3&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3090
      Locked          =   -1  'True
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   7815
      Width           =   3585
   End
   Begin VB.TextBox txtPacking 
      BackColor       =   &H00FDDFE3&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   3090
      Locked          =   -1  'True
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   7485
      Width           =   3585
   End
   Begin VB.TextBox txtPrice 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   10770
      MaxLength       =   19
      TabIndex        =   15
      Text            =   "9,999,999.99"
      Top             =   6030
      Width           =   1395
   End
   Begin VB.CommandButton command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Create"
      Height          =   375
      Index           =   0
      Left            =   11850
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1830
      Width           =   1125
   End
   Begin VB.TextBox txtpono 
      Height          =   315
      Left            =   2400
      MaxLength       =   25
      TabIndex        =   3
      Top             =   1870
      Width           =   2470
   End
   Begin VB.TextBox txtRev 
      Height          =   315
      Left            =   8760
      MaxLength       =   20
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1870
      Width           =   525
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1035
      Left            =   188
      TabIndex        =   44
      Top             =   720
      Width           =   14865
      Begin VB.TextBox lblcust 
         BackColor       =   &H00FDDFE3&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   1
         Left            =   8760
         Locked          =   -1  'True
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   280
         Width           =   5835
      End
      Begin VB.TextBox lblcust 
         BackColor       =   &H00FDDFE3&
         BorderStyle     =   0  'None
         Height          =   200
         Index           =   0
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   280
         Width           =   4755
      End
      Begin MSComCtl2.DTPicker Period 
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Top             =   640
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
      Begin VB.Label LblCode 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier CD"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   250
         Width           =   1215
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   2880
         X2              =   7680
         Y1              =   525
         Y2              =   525
      End
      Begin MSForms.ComboBox cbocust 
         Height          =   315
         Left            =   1440
         TabIndex        =   0
         Top             =   210
         Width           =   1305
         VariousPropertyBits=   612386843
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2302;556"
         ColumnCount     =   12
         ListRows        =   15
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   255
         Index           =   4
         Left            =   7920
         TabIndex        =   48
         Top             =   270
         Width           =   840
      End
      Begin VB.Line Line2 
         Index           =   1
         X1              =   8760
         X2              =   14640
         Y1              =   525
         Y2              =   525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Period"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   47
         Top             =   690
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdReport 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Preview"
      Height          =   375
      Left            =   10260
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   10050
      Width           =   1125
   End
   Begin VB.CommandButton command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Cancel"
      Height          =   375
      Index           =   3
      Left            =   11505
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   10050
      Width           =   1125
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FDDFE3&
      Height          =   533
      Left            =   120
      TabIndex        =   37
      Top             =   9420
      Width           =   15045
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
         Left            =   60
         TabIndex        =   38
         Top             =   120
         Width           =   14895
      End
   End
   Begin VB.CommandButton command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Submit"
      Height          =   375
      Index           =   1
      Left            =   13980
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   10050
      Width           =   1125
   End
   Begin VB.CommandButton cmdSubMenu 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Sub &Menu"
      Height          =   375
      Left            =   188
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   10140
      Width           =   1125
   End
   Begin VB.CommandButton command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Clea&r"
      Height          =   375
      Index           =   2
      Left            =   12750
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   10050
      Width           =   1125
   End
   Begin VB.TextBox txtQty 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   7620
      MaxLength       =   12
      TabIndex        =   12
      Text            =   "9,999,999.99"
      Top             =   6030
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DelDate 
      Height          =   315
      Left            =   6090
      TabIndex        =   11
      Top             =   6030
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
   Begin MSComCtl2.DTPicker podate 
      Height          =   315
      Left            =   6360
      TabIndex        =   4
      Top             =   1870
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
   Begin VSFlex8Ctl.VSFlexGrid grid 
      Height          =   3150
      Left            =   195
      TabIndex        =   7
      Top             =   2295
      Width           =   14955
      _cx             =   26379
      _cy             =   5556
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
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   13170
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   180
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   714
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PO Number"
      Height          =   195
      Index           =   2
      Left            =   360
      TabIndex        =   89
      Top             =   8670
      Width           =   975
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Discount"
      Height          =   195
      Index           =   28
      Left            =   6285
      TabIndex        =   88
      Top             =   8670
      Width           =   735
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Total"
      Height          =   195
      Index           =   29
      Left            =   3885
      TabIndex        =   87
      Top             =   8670
      Width           =   810
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Service"
      Height          =   255
      Left            =   12300
      TabIndex        =   84
      Top             =   5670
      Width           =   705
   End
   Begin MSForms.ComboBox combo1 
      Height          =   315
      Left            =   195
      TabIndex        =   2
      Top             =   1870
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      Height          =   255
      Left            =   13590
      TabIndex        =   83
      Top             =   5670
      Width           =   1455
   End
   Begin VB.Line Line8 
      Index           =   1
      X1              =   3060
      X2              =   6630
      Y1              =   8430
      Y2              =   8430
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PPn"
      Height          =   195
      Index           =   3
      Left            =   11460
      TabIndex        =   82
      Top             =   8640
      Width           =   315
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
      Height          =   195
      Index           =   1
      Left            =   8400
      TabIndex        =   81
      Top             =   8670
      Width           =   1140
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grand Total"
      Height          =   195
      Index           =   0
      Left            =   13350
      TabIndex        =   80
      Top             =   8670
      Width           =   1005
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00A6D2FF&
      BackStyle       =   1  'Opaque
      Height          =   300
      Index           =   1
      Left            =   300
      Top             =   8610
      Width           =   14805
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00A6D2FF&
      Height          =   2685
      Index           =   1
      Left            =   150
      Top             =   6660
      Width           =   15015
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Line3"
      Height          =   195
      Index           =   20
      Left            =   9600
      TabIndex        =   75
      Top             =   7155
      Width           =   450
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Line1"
      Height          =   195
      Index           =   18
      Left            =   6825
      TabIndex        =   74
      Top             =   7155
      Width           =   450
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Line5"
      Height          =   195
      Index           =   23
      Left            =   12330
      TabIndex        =   73
      Top             =   7155
      Width           =   450
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PO Marking"
      Height          =   195
      Index           =   22
      Left            =   6720
      TabIndex        =   72
      Top             =   6750
      Width           =   975
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      Height          =   195
      Index           =   8
      Left            =   6930
      TabIndex        =   71
      Top             =   8010
      Width           =   765
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Line4"
      Height          =   195
      Index           =   21
      Left            =   9570
      TabIndex        =   70
      Top             =   7545
      Width           =   450
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Line2 "
      Height          =   195
      Index           =   19
      Left            =   6825
      TabIndex        =   69
      Top             =   7530
      Width           =   510
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00A6D2FF&
      Height          =   915
      Left            =   6750
      Top             =   6990
      Width           =   8325
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Line6"
      Height          =   195
      Index           =   27
      Left            =   12360
      TabIndex        =   68
      Top             =   7545
      Width           =   450
   End
   Begin VB.Line Line4 
      X1              =   3105
      X2              =   6660
      Y1              =   7395
      Y2              =   7395
   End
   Begin VB.Line Line5 
      X1              =   3090
      X2              =   6645
      Y1              =   7005
      Y2              =   7005
   End
   Begin VB.Line Line6 
      X1              =   3090
      X2              =   6660
      Y1              =   7725
      Y2              =   7725
   End
   Begin VB.Line Line8 
      Index           =   0
      X1              =   3090
      X2              =   6660
      Y1              =   8055
      Y2              =   8055
   End
   Begin MSForms.ComboBox cboPriceCondition 
      Height          =   315
      Left            =   1980
      TabIndex        =   19
      Top             =   7080
      Width           =   975
      VariousPropertyBits=   612386843
      MaxLength       =   6
      DisplayStyle    =   3
      Size            =   "1720;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Price Condition"
      Height          =   195
      Index           =   13
      Left            =   645
      TabIndex        =   62
      Top             =   7020
      Width           =   1290
   End
   Begin MSForms.ComboBox cboPaymentTerm 
      Height          =   315
      Left            =   1980
      TabIndex        =   18
      Top             =   6720
      Width           =   975
      VariousPropertyBits=   612386843
      MaxLength       =   6
      DisplayStyle    =   3
      Size            =   "1720;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Term"
      Height          =   195
      Index           =   14
      Left            =   675
      TabIndex        =   61
      Top             =   6750
      Width           =   1260
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Packing"
      Height          =   195
      Index           =   15
      Left            =   1275
      TabIndex        =   60
      Top             =   7380
      Width           =   660
   End
   Begin MSForms.ComboBox cboTransport 
      Height          =   315
      Left            =   1980
      TabIndex        =   22
      Top             =   8130
      Width           =   975
      VariousPropertyBits=   612386843
      MaxLength       =   6
      DisplayStyle    =   3
      Size            =   "1720;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transportation"
      Height          =   195
      Index           =   17
      Left            =   690
      TabIndex        =   59
      Top             =   8220
      Width           =   1245
   End
   Begin MSForms.ComboBox cboInsuranceCls 
      Height          =   315
      Left            =   1980
      TabIndex        =   21
      Top             =   7770
      Width           =   975
      VariousPropertyBits=   612386843
      MaxLength       =   6
      DisplayStyle    =   3
      Size            =   "1720;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Insurance Covered"
      Height          =   195
      Index           =   16
      Left            =   285
      TabIndex        =   58
      Top             =   7830
      Width           =   1650
   End
   Begin MSForms.ComboBox cboPacking 
      Height          =   315
      Left            =   1980
      TabIndex        =   20
      Top             =   7440
      Width           =   975
      VariousPropertyBits=   612386843
      MaxLength       =   6
      DisplayStyle    =   3
      Size            =   "1720;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cboCls 
      Height          =   315
      Left            =   180
      TabIndex        =   8
      Top             =   6030
      Width           =   1275
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "2249;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Cls"
      Height          =   255
      Left            =   270
      TabIndex        =   57
      Top             =   5670
      Width           =   1335
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code"
      Height          =   255
      Left            =   1470
      TabIndex        =   56
      Top             =   5670
      Width           =   1350
   End
   Begin MSForms.ComboBox cboItemCode 
      Height          =   315
      Left            =   1500
      TabIndex        =   9
      Top             =   6030
      Width           =   1890
      VariousPropertyBits=   612386843
      MaxLength       =   15
      DisplayStyle    =   3
      Size            =   "3334;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cboItemName 
      Height          =   315
      Left            =   3420
      TabIndex        =   10
      Top             =   6030
      Width           =   2625
      VariousPropertyBits=   612386843
      MaxLength       =   50
      DisplayStyle    =   3
      Size            =   "4630;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Curr"
      Height          =   255
      Left            =   9900
      TabIndex        =   55
      Top             =   5670
      Width           =   840
   End
   Begin VB.Label e 
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      Height          =   255
      Left            =   10830
      TabIndex        =   54
      Top             =   5670
      Width           =   705
   End
   Begin MSForms.ComboBox cbocurr 
      Height          =   315
      Left            =   9870
      TabIndex        =   14
      Top             =   6030
      Width           =   840
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "1482;556"
      TextColumn      =   2
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cbopono 
      Height          =   315
      Left            =   2400
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   1870
      Width           =   2745
      VariousPropertyBits=   612386843
      DisplayStyle    =   3
      Size            =   "4842;556"
      ListWidth       =   5291
      ListRows        =   15
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "PO Date"
      Height          =   255
      Index           =   6
      Left            =   5520
      TabIndex        =   53
      Top             =   1910
      Width           =   825
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "PO No"
      Height          =   255
      Index           =   2
      Left            =   1680
      TabIndex        =   52
      Top             =   1910
      Width           =   720
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
      Left            =   13875
      TabIndex        =   51
      Top             =   1910
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Rev"
      Height          =   255
      Index           =   1
      Left            =   8280
      TabIndex        =   50
      Top             =   1910
      Width           =   525
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery Date"
      Height          =   255
      Index           =   0
      Left            =   6210
      TabIndex        =   43
      Top             =   5670
      Width           =   1485
   End
   Begin MSForms.ComboBox cboUnit 
      Height          =   315
      Left            =   8940
      TabIndex        =   13
      Top             =   6030
      Width           =   825
      VariousPropertyBits=   746604569
      MaxLength       =   15
      DisplayStyle    =   7
      Size            =   "1455;556"
      ListRows        =   15
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontEffects     =   1073750016
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Order Unscheduled (Others)"
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
      Left            =   2130
      TabIndex        =   42
      Top             =   240
      Width           =   10980
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00A6D2FF&
      Height          =   1005
      Index           =   0
      Left            =   150
      Top             =   5550
      Width           =   15015
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit"
      Height          =   255
      Left            =   9030
      TabIndex        =   41
      Top             =   5670
      Width           =   825
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      Height          =   255
      Left            =   3420
      TabIndex        =   40
      Top             =   5670
      Width           =   1635
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Qty"
      Height          =   255
      Left            =   7740
      TabIndex        =   39
      Top             =   5670
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00A6D2FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   0
      Left            =   150
      Top             =   5625
      Width           =   14865
   End
End
Attribute VB_Name = "frmPOOther"
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
Dim countrycls As Byte, isippn As Long, j As Integer, Baris As Integer
Const isiPOTerm = "after B/L Date,after Delivery Date,after Invoice Date,prior before Shipment,from Custom Clearance Date,after Receive Invoice,after Receive Goods"
Dim ColItem As Byte, ColProd As Byte, ColDelDate As Byte, ColUnit As Byte, ColCurr As Byte, ColPrice As Byte
Dim ColAmount As Byte, colRemark As Byte, ColQty As Byte, colSeqNo As Byte, ColService As Byte



Private Sub ClearData()

    sql = "Delete from PurchaseOrder_Master Where PO_No = '" & Trim(txtPoNo.Text) & "' " & _
          "and PO_No not in (select PO_No from PurchaseOrder_Detail) and others_cls = '1' and period is not null"
    Db.Execute sql

End Sub

Private Sub CekPONumber()
    
    Dim adoRs As New ADODB.Recordset
    Set adoRs = Nothing
    sql = "Select * From PurchaseOrder_Master Where PO_No = '" & Trim(txtPoNo.Text) & "'"
    adoRs.Open sql, Db, 1, 3
    If Not adoRs.EOF Then
        Call PONO(Right(Year(PODate), 2), Format(Month(PODate), "0#"))
        adoRs!po_no = Trim(txtPoNo.Text)
        adoRs.update
    End If
    adoRs.Close
    Set adoRs = Nothing
    
End Sub

Sub Kosong()
    LblErrMsg = ""
    cboCust.Text = ""
    lblcust(0).Text = "": lblcust(1).Text = ""
    Period.Value = Format(Now, "MMM yyyy")
    temptgl = Period.Month
    txtPoNo.Text = "": txtPONo2.Text = ""
    PODate.Value = Format(Now, "dd MMM yyyy")
    PODate.Enabled = True
    Call ppn(PODate.Value)
    txtRev.Text = ""
    cboPaymentTerm = "": txtPaymentTerm = ""
    cboPriceCondition = "": txtPriceCondition = ""
    CboPacking = "": TxtPacking = ""
    cboInsuranceCls = "": txtInsurance = ""
    cboTransport = "": txtInsurance = ""
    txtMarking(0) = "": txtMarking(1) = "": txtMarking(2) = "": txtMarking(3) = "": txtMarking(4) = "": txtMarking(5) = ""
    cboPriceCondition.Text = ""
   
    ' Add 20090112
    TxtSubAmount.Text = 0
    TxtDisc.Text = 0
    ' ---
    txtamount.Text = 0
    txtPPN.Text = 0
    txtGrandTotal.Text = 0
    
    ubah = False: ada = False
    statusfix = 0: Call kunci(False)
    Call kosongBwh
    'Call header
    'Call adtocbopono
    
End Sub

Sub kosongBwh()
    CboItemCode.Text = ""
    CboItemCode.Enabled = True
    cboItemName.Text = ""
    cboItemName.Enabled = True
    cboCls.Enabled = True
    txtRev = ""
    DelDate.Value = Format(Now, "dd MMM yyyy")
    txtQty.Text = ""
    lblQty = 0
    cbounit.ListIndex = -1
    cbocurr.ListIndex = -1
    TxtService = ""
    txtremarks = ""
    txtprice.Text = ""
    txtSubTotal = ""
    lblseqno = 0
        cboPaymentTerm = "": txtPaymentTerm = ""
    cboPriceCondition = "": txtPriceCondition = ""
    CboPacking = "": TxtPacking = ""
    cboInsuranceCls = "": txtInsurance = ""
    cboTransport = "": txtInsurance = ""
    txtMarking(0) = "": txtMarking(1) = "": txtMarking(2) = "": txtMarking(3) = "": txtMarking(4) = "": txtMarking(5) = ""
    cboPriceCondition.Text = ""

    
    
End Sub

Function adtocboCust()
Dim sqlcust As String
Dim RsCust As New Recordset

    sqlcust = "select trade_code, trade_name, address1, country_cls, po_cls, isnull(trade_Abbr,'') Trade_Abbr " & _
              ", popayment_code, popayment_day, popayment_terms, transportation_Cls " & _
              "from trade_master where ((trade_cls='2'   or trade_cls='3')) Order By Trade_Abbr" ' ((trade_cls='2' And TradeExternal_Cls in ('1','2'))  or trade_cls='3') Order By Trade_Abbr"
    Set RsCust = Db.Execute(sqlcust)

    With cboCust
        .clear
        .columnCount = 8
        .ColumnWidths = "100pt;150pt;0pt;0pt;0pt;0pt;0pt;0pt"
        .ListWidth = 360
        .ListRows = 15

        i = 0
        Do While Not RsCust.EOF
            .AddItem
            .List(i, 0) = Trim(RsCust("Trade_code"))
            .List(i, 1) = IIf(IsNull(RsCust("trade_abbr")), "", Trim(RsCust("Trade_name")))
            .List(i, 2) = IIf(IsNull(RsCust("address1")), "", Trim(RsCust("Address1")))
            .List(i, 3) = IIf(IsNull(RsCust("country_cls")), 3, Trim(RsCust("country_cls")))
            .List(i, 4) = IIf(IsNull(RsCust("po_cls")), 0, Trim(RsCust("po_cls")))
            .List(i, 5) = IIf(IsNull(RsCust("popayment_code")), "", RsCust("popayment_code")) & "," & IIf(IsNull(RsCust("popayment_day")), "", RsCust("popayment_day")) & "," & IIf(IsNull(RsCust("popayment_terms")), "", RsCust("popayment_terms"))
            .List(i, 6) = IIf(IsNull(RsCust("transportation_Cls")), -1, RsCust("transportation_Cls"))
            .List(i, 7) = IIf(IsNull(RsCust("trade_name")), "", Trim(RsCust("Trade_abbr")))
            RsCust.MoveNext
            i = i + 1
        Loop
    End With
    Set RsCust = Nothing
End Function

Sub adtocboDept()
Dim sqldept As String
Dim rsdept As New Recordset

    sqldept = "select * from Department_Cls order by Department_cls"
    Set rsdept = Db.Execute(sqldept)
    
'    With cboDept
'        .clear
'        .ColumnCount = 2
'        .ColumnWidths = "50pt;150pt"
'        .ListWidth = 200
'        .ListRows = 15
'
'        i = 0
'        Do While Not rsdept.EOF
'            .AddItem
'            .List(i, 0) = Trim(rsdept("Department_cls"))
'            .List(i, 1) = IIf(IsNull(rsdept("Description")), "", Trim(rsdept("Description")))
'            rsdept.MoveNext
'            i = i + 1
'        Loop
'    End With
    Set rsdept = Nothing
End Sub

Sub adtocboClass()
'    With cboClass
'        .clear
'        .ColumnCount = 2
'        .TextColumn = 1
'
'        For i = 0 To UBound(Split(POClass1, ","))
'            .AddItem
'            .List(i, 0) = Split(POClass1, ",")(i)
'            .List(i, 1) = Split(POClass2, ",")(i)
'        Next i
'        .ListRows = 10
'        .ListWidth = 140
'        .ColumnWidths = "30pt;110pt"
'    End With
End Sub

Sub adtocbopono()
Dim sqlno As String
Dim rsno As New Recordset
    
    sqlno = "select pom.PO_No from PurchaseOrder_Master pom " & _
            "where pom.others_cls = '1' and pom.period is not null " & _
            "and year(po_date) = '" & Year(PODate) & "' " & _
            "and month(po_date) = '" & Month(PODate) & "'  AND Supplier_Code='" & cboCust.Column(0) & "'"
    Set rsno = Db.Execute(sqlno)
    With CboPOnO
        .clear
        Do While Not rsno.EOF
            .AddItem Trim(rsno("PO_No"))
            rsno.MoveNext
        Loop
        .ColumnWidths = "150pt"
        .ListWidth = 150
        .ListRows = 15
    End With
    Set rsno = Nothing
End Sub

Sub adtocboitem(ByVal nmCombo, ByVal field1 As String, ByVal field2 As String, ByVal colWidth1 As Integer, ByVal colwidth2 As Integer, ByVal Orderby As String)
Dim sqlitem As String
Dim RsItem As New Recordset

    sqlitem = "select accounting_code, unit_Cls, " & field1 & ", " & field2 & " from item_master " & _
              "where finishgoodpart_cls = '02' and stockcontrol_Cls = '02' order by " & Orderby
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

Sub adtocbo(nmCombo, nmConst, start As Integer, col1 As Integer, Col2 As Integer, Kosong As Boolean)
Dim j As Integer, k As Integer

    With nmCombo
        .clear
        .columnCount = 2
        .TextColumn = 2
        j = 0: k = start
        
        If Kosong Then
            .AddItem
            .List(0, 0) = "Null"
            .List(0, 1) = ""
            j = 1
        End If
        
        For i = 0 To UBound(Split(nmConst, ","))
            .AddItem
            .List(j, 0) = k
            .List(j, 1) = Split(nmConst, ",")(i)
            k = k + 1: j = j + 1
        Next i
        .ListRows = 10
        .ListWidth = col1 + Col2
        .ColumnWidths = col1 & "pt;" & Col2 & "pt"
    End With
End Sub

Sub adtocboitemOthers()
Dim sqlitem As String
Dim RsItem As New Recordset

    sqlitem = "select * from OthersItem_Master order by item_desc"
    Set RsItem = Db.Execute(sqlitem)
    
    With cboItemName
        .clear
        .columnCount = 2
        .ColumnWidths = "300pt;80pt"
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

Sub adtocombo()
Dim sql1 As String
Dim rs1 As New Recordset

    combo1.AddItem "Create"
    combo1.AddItem "Update"
    Call adtocboCust
    Call adtocboDept
    Call adtocboitem(cboItemName, "item_name", "item_code", 250, 50, "item_name")
    Call adtocboitem(CboItemCode, "item_code", "item_name", 50, 250, "item_code")
    Call up_FillCombo(cbounit, "unit_cls")
 
    Call up_FillCombo(cboPaymentTerm, "PaymentTerm_Cls")
    cboPaymentTerm.ColumnWidths = "25pt;175pt"
    cboPaymentTerm.ListWidth = 200
    
    Call up_FillCombo(cboInsuranceCls, "Insurance_Cls")
    cboInsuranceCls.ColumnWidths = "25pt;175pt"
    cboInsuranceCls.ListWidth = 200
    
    Call up_FillCombo(CboPacking, "POPacking_Cls")
    CboPacking.ColumnWidths = "25pt;175pt"
    CboPacking.ListWidth = 200
    
    
    Call up_FillCombo(cboTransport, "Transportation_Cls")
    cboTransport.ColumnWidths = "25pt;175pt"
    cboTransport.ListWidth = 200
    
 
    cbounit.TextColumn = 2
 
    Call up_FillCombo(cbocurr, "curr_cls")
    cbocurr.TextColumn = 2
    
    Call adtocboClass
    cboCls.AddItem "By Code"
    cboCls.AddItem "Non Code"
    
    'Call adtocbo(cbopocode, isiPOCode, 1, 0, 40, False)
    'Call adtocbo(cbopoterms, isiPOTerm, 1, 0, 140, False)
    'Call adtocbo(cbodelmode, isiTransport, 0, 0, 80, False)

    'PRICE CONDITION
    sql1 = "select * from PriceCondition_cls"
    If rs1.State <> adStateClosed Then rs1.Close
    rs1.Open sql1, Db, adOpenKeyset, adLockOptimistic
    If Not (rs1.BOF And rs1.EOF) Then
        i = 0
        With cboPriceCondition
            .clear
            .columnCount = 2
            .ColumnWidths = "25pt;175pt"
            .ListWidth = 200
            .ListRows = 15
            
            Do While Not rs1.EOF
                .AddItem ""
                .List(i, 0) = Trim(rs1!PriceCondition_Cls)
                .List(i, 1) = Trim(rs1!Description)
                i = i + 1
                rs1.MoveNext
            Loop
        End With
    End If
    Set rs1 = Nothing
End Sub

Sub PONO(ByVal thn As String, ByVal bln As String)
Dim sqlno As String, SqlS As String
Dim rsno As New Recordset, rsS As New Recordset
    
    'POYYMM999
    If Trim(txtPoNo.Text) = "" Then
'        If Format(podate, "YYYY-MM-01") > "2006-07-30" Then
'            sqlno = "select top 1 rtrim(PO_No) from PurchaseOrder_Master " & _
'                    "where substring(rtrim(PO_No),3,2) = '" & thn & "' and substring(rtrim(PO_No),5,2) > '07' " & _
'                    "order by right(rtrim(PO_No),5) desc"
'        Else
            sqlno = "select top 1 rtrim(PO_No) from PurchaseOrder_Master " & _
                    "where substring(rtrim(PO_No),3,2) = '" & thn & "' " & _
                    "order by right(rtrim(PO_No),5) desc"
'        End If
    Else
'        If Format(podate, "YYYY-MM-01") > "2006-07-30" Then
'            sqlno = "select top 1 rtrim(PO_No) from PurchaseOrder_Master " & _
'                    "where substring(rtrim(PO_No),3,2) = '" & thn & "' and substring(rtrim(PO_No),5,2) > '07' " & _
'                    "and Right(RTrim(PO_No), 9)  < '" & Right(Trim(txtpono.Text), 9) & "' " & _
'                    "order by right(rtrim(PO_No),5) desc"
'        Else
            sqlno = "select top 1 rtrim(PO_No) from PurchaseOrder_Master " & _
                    "where substring(rtrim(PO_No),3,2) = '" & thn & "' " & _
                    "and Right(RTrim(PO_No), 9)  < '" & Right(Trim(txtPoNo.Text), 9) & "' " & _
                    "order by right(rtrim(PO_No),5) desc"
'        End If
    End If
    Set rsno = Db.Execute(sqlno)
    If Not (rsno.BOF And rsno.EOF) Then
        txtPoNo.Text = Left(Trim(rsno(0)), 4) & bln & Format(Right(Trim(rsno(0)), 5) + 1, "0000#")
    Else
        txtPoNo.Text = "PO" & thn & bln & "00001"
    End If
    txtPoNo.locked = True
    Set rsno = Nothing
End Sub

Sub kunci(l As Boolean)
    Period.Enabled = Not l
    PODate.Enabled = Not l
    grid.Editable = Not l
    Command1(1).Enabled = Not l
    CboPacking.Enabled = Not l
    cboPaymentTerm.Enabled = Not l
    cboPriceCondition.Enabled = Not l
    cboTransport.Enabled = Not l
    cboInsuranceCls.Enabled = Not l
    For i = 0 To 5
    txtMarking(i).Enabled = Not l
    Next
    lblFix.Caption = "Status Fix"
    lblFix.Visible = l
End Sub

Sub ppn(ByVal d As Date)
Dim sqlppn As String
Dim rsppn As New ADODB.Recordset
    
    sqlppn = "select rate from tax_cls where tax_code='PPN' and " & _
             "start_date <= '" & Format(d, "yyyymmdd") & "' and " & _
             "end_date >= '" & Format(d, "yyyymmdd") & "' "
    Set rsppn = Db.Execute(sqlppn)
    If Not (rsppn.BOF And rsppn.EOF) Then
        isippn = IIf(IsNull(rsppn(0)), 0, CDbl(rsppn(0)))
    Else
        isippn = 0
    End If
    Set rsppn = Nothing
End Sub

Sub Header()
    ColItem = 1
    ColProd = 2: ColDelDate = 3: ColQty = 4: ColUnit = 5: ColCurr = 6
    ColPrice = 7: ColService = 8: ColAmount = 9: colRemark = 10: colSeqNo = 11

    With grid
        .clear
        .Rows = 1
        .ColS = 12
    
        .ColWidth(0) = 300
        .ColWidth(ColItem) = 1300
        .ColWidth(ColProd) = 2500
        .ColHidden(ColDelDate) = 1500
        .ColWidth(ColQty) = 1200
        .ColHidden(ColUnit) = 10000
        .ColWidth(ColCurr) = 1000
        .ColWidth(ColPrice) = 1750
        .ColWidth(ColService) = 1750
        .ColWidth(ColAmount) = 2000
        .ColWidth(colRemark) = 2500
        .ColHidden(colSeqNo) = True
        .ColHidden(colRemark) = True
        
        
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, ColItem) = "Product Code"
        .TextMatrix(0, ColProd) = "Product Name"
        .TextMatrix(0, ColDelDate) = "Delivery Date"
        .TextMatrix(0, ColQty) = "Qty"
        .TextMatrix(0, ColUnit) = "Unit"
        .TextMatrix(0, ColCurr) = "Curr"
        .TextMatrix(0, ColPrice) = "Price"
        .TextMatrix(0, ColService) = "Service"
        .TextMatrix(0, ColAmount) = "Amount"
        .TextMatrix(0, colRemark) = "RemarksPurpose"
        
    
        .Cell(flexcpAlignment, 0, 0, 0, colRemark) = flexAlignCenterCenter
        .ColAlignment(ColPrice) = flexAlignLeftCenter
        .ColAlignment(ColProd) = flexAlignLeftCenter
        .ColAlignment(ColItem) = flexAlignLeftCenter
        .ColAlignment(colRemark) = flexAlignLeftCenter
        .ColAlignment(ColQty) = flexAlignCenterCenter
        .ColAlignment(ColUnit) = flexAlignRightCenter
        .ColAlignment(ColCurr) = flexAlignCenterCenter
        
        .EditMaxLength = 1
    End With
End Sub

Sub Browse()
    LblErrMsg = ""
    sql = "select * from PurchaseOrder_Master where PO_No = '" & txtPoNo.Text & "' and others_cls = '1' and period is not null"
    If RS.State <> adStateClosed Then RS.Close
    RS.Open sql, Db, adOpenKeyset, adLockOptimistic
    
    If Not (RS.BOF And RS.EOF) Then
        ada = True: ubah = True
        statusfix = IIf(IsNull(RS("fix_cls")), 0, RS("fix_cls"))
        Call BrowseGrid
        txtamount.Text = Format(IS_NOL(RS("Amount")), "##,##0.#0")
        If cboCust.Column(3) = "4" Or cboCust.Column(3) = "6" Then
            txtPPN.Text = CDbl(isippn / 100) * CDbl(txtamount.Text)
        Else
            txtPPN.Text = 0
        End If
        txtPPN.Text = IIf(IsNull(RS("PPN")), 0, Format(Trim(RS("PPN")), "##,##0.#0"))
        txtGrandTotal.Text = IIf(IsNull(RS("Total_Amount")), 0, Format(Trim(RS("Total_Amount")), "##,##0.#0"))
        cboPaymentTerm.Text = is_null(RS!PaymentTerm_Cls)
        cboPriceCondition.Text = is_null(RS!PriceCondition_Cls)
        CboPacking.Text = is_null(RS!POPacking_Cls)
        
        cboTransport.Text = is_null(RS!Transportation_Cls)
        cboInsuranceCls.Text = is_null(RS!Insurance_Cls)
        txtMarking(0).Text = is_null(RS!POMarking1)
        txtMarking(1).Text = is_null(RS!POMarking2)
        txtMarking(2).Text = is_null(RS!POMarking3)
        txtMarking(3).Text = is_null(RS!POMarking4)
        txtMarking(4).Text = is_null(RS!POMarking5)
        txtMarking(5).Text = is_null(RS!POMarking6)
        txtremarks.Text = is_null(RS!Remarks)
        TxtDisc = Format(RS!discount, gs_formatAmount) ' Add 20090112
        
        Call hitungTotal
        If statusfix = 1 Then Call kunci(True) Else Call kunci(False)
    Else
        ada = False
    End If
End Sub

Sub BrowseGrid()

    Call kosongBwh
    
    sqlGrid = "select pod.*,(select description from unit_cls uc where uc.unit_cls= pod.unit_cls ) unit_desc, (select description from curr_cls where curr_cls.Curr_cls= pod.Currency_Code) Curr_desc from PurchaseOrder_Detail pod " & _
              "where pod.PO_No = '" & Trim(txtPoNo.Text) & "' order by pod.item_name"
    If rsGrid.State <> adStateClosed Then rsGrid.Close
    rsGrid.Open sqlGrid, Db, adOpenKeyset, adLockOptimistic
    i = 1
    grid.Rows = 1
    With grid
    Do While Not rsGrid.EOF
        .Rows = .Rows + 1
        .Cell(flexcpBackColor, i, 0) = &HFFFFFF
        .TextMatrix(i, ColItem) = Trim(rsGrid("Item_Code"))
        .TextMatrix(i, ColProd) = is_null(Trim(rsGrid("Item_Name")))
        
        .TextMatrix(i, ColDelDate) = Format(Trim(rsGrid("Delivery_date")), "dd MMM yyyy")
        .TextMatrix(i, ColQty) = IS_NOL(rsGrid("Qty"))
        .TextMatrix(i, ColUnit) = is_null(rsGrid("Unit_desc"))
        .TextMatrix(i, ColCurr) = is_null(Trim(rsGrid("Curr_desc")))
        .TextMatrix(i, ColPrice) = Format(IS_NOL(rsGrid("Price")), "##,##0.#0")
        .TextMatrix(i, ColService) = Format(IS_NOL(rsGrid("Price_Service")), "##,##0.#0")
        .TextMatrix(i, ColAmount) = Format(IS_NOL(rsGrid("Amount")), "##,##0.#0")
        '.TextMatrix(i, colRemark) = is_null(rsGrid("remarks"))
        .TextMatrix(i, colSeqNo) = IS_NOL(rsGrid("PoReq_seqno"))
        
        rsGrid.MoveNext
        i = i + 1
    Loop
    End With
    
End Sub
Function IS_NOL(Data)
If IsNull(Data) Then
Data = 0
Exit Function
End If
If Data = "" Then
IS_NOL = 0
Else
IS_NOL = Data
End If

End Function
Sub BrowseAtas()
Dim p As String

    sql = "select * from PurchaseOrder_Master where PO_No = '" & Trim(txtPoNo.Text) & "' and isnull(others_cls,'0') = '1' and period is not null"
    If RS.State <> adStateClosed Then RS.Close
    RS.Open sql, Db, adOpenKeyset, adLockOptimistic
    If Not (RS.BOF And RS.EOF) Then
        p = IIf(IsNull(RS("Period")), "", Left(Trim(RS("Period")), 4) & "-" & Right(Trim(RS("Period")), 2) & "-01")
        Period.Value = Format(p, "MMM yyyy")
        temptgl = Period.Month
        PODate.Value = IIf(IsNull(RS("po_date")), "", Format(Trim(RS("po_date")), "dd MMM yyyy"))
        cboCust.Text = Trim(RS("Supplier_code"))
        txtRev.Text = IIf(IsNull(RS("revise_No")), "", Trim(RS("revise_No")))
        cboPriceCondition.Text = IIf(IsNull(RS("PriceCondition_Cls")), "", Trim(RS("PriceCondition_Cls")))
        txtamount.Text = IIf(IsNull(RS("Amount")), 0, Format(Trim(RS("Amount")), "##,##0.#0"))
        txtPPN.Text = IIf(IsNull(RS("PPN")), 0, Format(Trim(RS("PPN")), "##,##0.#0"))
        txtGrandTotal.Text = IIf(IsNull(RS("Total_Amount")), 0, Format(Trim(RS("Total_Amount")), "##,##0.#0"))
        statusfix = IIf(IsNull(RS("fix_cls")), 0, RS("fix_cls"))
        cboPaymentTerm.Text = is_null(RS!PaymentTerm_Cls)
        CboPacking.Text = is_null(RS!POPacking_Cls)
        cboTransport.Text = is_null(RS!Transportation_Cls)
        cboInsuranceCls.Text = is_null(RS!Insurance_Cls)
        txtMarking(0).Text = is_null(RS!POMarking1)
        txtMarking(1).Text = is_null(RS!POMarking2)
        txtMarking(2).Text = is_null(RS!POMarking3)
        txtMarking(3).Text = is_null(RS!POMarking4)
        txtMarking(4).Text = is_null(RS!POMarking5)
        txtMarking(5).Text = is_null(RS!POMarking6)
        txtremarks = is_null(RS!Remarks)
        TxtDisc = Format(RS!discount, gs_formatAmount) ' Add 20090112
        
        If statusfix = 1 Then Call kunci(True) Else Call kunci(False)
    End If
End Sub
Function GetNo(sql)
Dim rr As New ADODB.Recordset
rr.Open sql, Db, adOpenStatic, adLockReadOnly
If rr.EOF Then Exit Function
GetNo = rr.Fields(0)

End Function
Function seqNo() As Long
Dim sqlseqno As String
Dim rsseqno As New Recordset

    sqlseqno = "select Poreq_SeqNo from PurchaseOrder_Detail order by Poreq_SeqNo desc"
    If rsseqno.State <> adStateClosed Then rsseqno.Close
    rsseqno.Open sqlseqno, Db, adOpenKeyset, adLockOptimistic
    
    If Not (rsseqno.BOF And rsseqno.EOF) Then
        seqNo = rsseqno!POReq_seqno + 1
    Else
        seqNo = 1
    End If
    Set rsseqno = Nothing
End Function

Function cekrecqty(ByVal seqNo As String, ByVal PONO As String) As Double
Dim sqlcek As String, rsCek As New Recordset
    
    cekrecqty = 0
    sqlcek = "select dailyseq_no, sum(qty) recqty from Part_Receipt " & _
             "where PO_No = '" & Trim(PONO) & "' and dailyseq_no = '" & IIf(Trim(seqNo) = "", 0, Trim(seqNo)) & "' " & _
             "group by dailyseq_no"
    If rsCek.State <> adStateClosed Then rsCek.Close
    rsCek.Open sqlcek, Db, adOpenKeyset, adLockOptimistic
    If Not (rsCek.BOF And rsCek.EOF) Then _
        cekrecqty = CDbl(rsCek("recqty"))
        
    Set rsCek = Nothing
End Function

Private Sub cboInsuranceCls_Change()
If cboInsuranceCls.MatchFound Then txtInsurance.Text = cboInsuranceCls.Column(1) Else txtInsurance.Text = ""
End Sub

Private Sub cboItemName_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub cboPacking_Change()
If CboPacking.MatchFound Then TxtPacking.Text = CboPacking.Column(1) Else TxtPacking.Text = ""
End Sub

Private Sub cboPaymentTerm_Change()
If cboPaymentTerm.MatchFound Then txtPaymentTerm.Text = cboPaymentTerm.Column(1) Else txtPaymentTerm.Text = ""
End Sub

Private Sub cbopricecondition_Change()
If cboPriceCondition.MatchFound Then txtPriceCondition.Text = cboPriceCondition.Column(1) Else txtPriceCondition.Text = ""
End Sub

Private Sub cboTransport_Change()
If cboTransport.MatchFound Then TxtTransport.Text = cboTransport.Column(1) Else TxtTransport.Text = ""
End Sub

Private Sub Form_Load()
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
    Call adtocombo
    
    Call Kosong
    combo1.ListIndex = 1
    Header
End Sub

Private Sub cboCust_Click()
Dim ketemu As Boolean

    LblErrMsg = ""
    ketemu = False
    Call kunci(False)

    If cboCust.ListIndex <> -1 Then
        lblcust(0).Text = cboCust.Column(1)
        lblcust(1).Text = cboCust.Column(2)
        countrycls = cboCust.Column(3)
        If combo1.ListIndex = 1 Then    'UPDATE
            Call ClearData
            Call adtocbopono
            For i = 0 To CboPOnO.ListCount - 1
                If txtPoNo.Text = CboPOnO.List(i) Then
                    ketemu = True
                    'cbopono.ListIndex = i
                    Exit For
                End If
            Next i
            If ketemu = False Then txtPoNo.Text = ""
            Call kosongBwh

            
            'Call header
        End If
    Else
        lblcust(0).Text = ""
        lblcust(1).Text = ""
        countrycls = 3
        CboPOnO.clear
        If combo1.ListIndex = 1 Then 'UPDATE
            Call kosongBwh
            'Call header
            txtPoNo.Text = ""
        End If
        If cboCust.Text <> "" Then LblErrMsg.Caption = DisplayMsg(4050) '"Record with this Supplier Code not Exist"
        Exit Sub
    End If
        
    If (countrycls = 1 Or countrycls = 2) Then 'OVERSEAS
        isippn = 0
        txtPPN.Text = 0
        txtGrandTotal.Text = txtamount.Text
        If (txtGrandTotal.Text <> 0) Then txtGrandTotal.Text = Format(txtGrandTotal.Text, "##,##0.#0")
    Else 'DOMESTIC
        Call ppn(PODate.Value)
        txtPPN.Text = CDbl(isippn / 100) * CDbl(txtamount.Text)
        If (txtPPN.Text <> 0) Then txtPPN.Text = Format(txtPPN.Text, "##,##0.#0")
        txtGrandTotal = CDbl(txtPPN.Text) + CDbl(txtamount.Text)
        If (txtGrandTotal.Text <> 0) Then txtGrandTotal.Text = Format(txtGrandTotal.Text, "##,##0.#0")
    End If
End Sub

Private Sub cbocust_LostFocus()
    Call cboCust_Click
End Sub

Private Sub cboCust_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call cboCust_Click
End Sub

Private Sub cbocust_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub period_Change()
    Call period_Click
    temptgl = Period.Month
    'If combo1.ListIndex = 1 Then Call header
End Sub

Private Sub period_Click()
    If Period.Month = 1 And Val(temptgl) = 12 Then Period.Year = Period.Year + 1
    If Period.Month = 12 And Val(temptgl) = 1 Then Period.Year = Period.Year - 1
End Sub

Private Sub Combo1_Click()
Dim ketemu As Boolean

    LblErrMsg = ""
    ketemu = False
    Call kunci(False)
    Call kosongBwh
    'Call header

    If combo1.ListIndex = 0 Then    'CREATE
        Call ClearData
        Command1(0).Caption = "&Create"
        ubah = False
        CboPOnO.locked = True
        txtPoNo.Text = ""
        PODate.Value = Format(Now, "dd MMM yyyy")
        PODate.Enabled = False
        Call PONO(Right(Year(PODate), 2), Format(Month(PODate), "0#"))
        TxtSubAmount.Text = 0
        TxtDisc = 0
        txtamount.Text = 0
        txtPPN.Text = 0
        txtGrandTotal.Text = 0
        cboPriceCondition.ListIndex = -1

    Else    'UPDATE
        If cboCust.Text = "" Then
            'cbopono.clear
            'txtpono.Text = ""
        Else
            Call adtocbopono
        End If
        ubah = True
        Command1(0).Caption = "&Update"
        CboPOnO.locked = False
        txtPoNo.locked = False
        PODate.Enabled = True

        For i = 0 To CboPOnO.ListCount - 1
            If txtPoNo.Text = CboPOnO.List(i) Then
                ketemu = True
                CboPOnO.ListIndex = i
                Exit For
            End If
        Next i
        If ketemu = False Then txtPoNo.Text = ""
    End If
End Sub

Private Sub combo1_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call Combo1_Click
End Sub

Private Sub cbopono_Click()
    LblErrMsg = ""
    txtPoNo.Text = CboPOnO.Text
    'Call header
    Call kosongBwh
    Call BrowseAtas
End Sub

Private Sub cbopono_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call cbopono_Click
End Sub

Private Sub TxtDisc_Change()
    If InStr(1, TxtDisc.Text, ",") = 1 Then TxtDisc.Text = Right(TxtDisc, Len(TxtDisc) - 1)
    If TxtDisc <> "" And TxtSubAmount <> "" And IsNumeric(TxtDisc) And IsNumeric(TxtSubAmount) Then txtamount.Text = Format(CDbl(TxtSubAmount.Text) - CDbl(TxtDisc.Text), "##,##0.#0")
End Sub

Private Sub TxtDisc_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then _
          KeyAscii = 0
    If InStr(1, TxtDisc, ".") > 0 Then If KeyAscii = Asc(".") Then KeyAscii = 0
End Sub

Private Sub TxtDisc_LostFocus()
txtSubTotal = CDbl(TxtSubAmount) - CDbl(TxtDisc)
TxtDisc = Format(TxtDisc, gs_formatAmount)
End Sub

Private Sub txtpono_Change()
Dim ketemu As Boolean

    txtPONo2.Text = txtPoNo.Text
    If combo1.ListIndex = 1 Then
        For i = 0 To CboPOnO.ListCount - 1
            If txtPoNo.Text = CboPOnO.List(i) Then
                ketemu = True
                CboPOnO.ListIndex = i
                Exit For
            End If
        Next
        If ketemu = False Then CboPOnO.ListIndex = -1
    End If
End Sub

Private Sub txtpono_KeyPress(KeyAscii As Integer)
    LblErrMsg = ""
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = Asc("'") Then KeyAscii = 0
    If KeyAscii = 13 Then
        If combo1.ListIndex = 0 Then
            SendKeys vbTab
        Else
            'Call header
            Call kosongBwh
            Call BrowseAtas
        End If
    End If
End Sub

Private Sub PODate_Change()
    LblErrMsg = ""
    If combo1.ListIndex = 0 Then 'CREATE
        Call PONO(Right(Year(PODate), 2), Format(Month(PODate), "0#"))
    Else    'UPDATE
        If cboCust.Text = "" Then
            CboPOnO.clear: txtPoNo.Text = ""
        Else
            Call adtocbopono
        End If
    End If
    If (countrycls = "1" Or countrycls = "2") Then isippn = 0 Else Call ppn(PODate.Value)
End Sub

Private Sub cboCls_Click()
    If UCase(Trim(cboCls.Text)) = "BY CODE" Then
        CboItemCode.Enabled = True
        Call adtocboitem(CboItemCode, "item_code", "item_name", 50, 250, "item_code")
        Call adtocboitem(cboItemName, "item_name", "item_code", 250, 50, "item_name")
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
    cboItemName.Text = ""
    cbounit.ListIndex = -1
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
    If txtQty <> "" And txtprice.Text <> "" And IsNumeric(txtQty) And IsNumeric(txtprice) Then txtSubTotal.Text = Format(CDbl(txtprice.Text) * CDbl(txtQty.Text), "##,##0.#0")
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then _
          KeyAscii = 0
    If InStr(1, txtQty.Text, ".") > 0 Then If KeyAscii = Asc(".") Then KeyAscii = 0
    If (txtQty & Chr(KeyAscii)) > 9999999.99 And KeyAscii <> vbKeyBack Then KeyAscii = 0
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

Private Sub txtPrice_Change()
    If InStr(1, txtprice.Text, ",") = 1 Then txtprice.Text = Right(txtprice, Len(txtprice) - 1)
    If txtQty.Text <> "" And txtprice.Text <> "" And IsNumeric(txtQty) And IsNumeric(txtprice) And TxtService.Text <> "" And IsNumeric(TxtService) Then txtSubTotal.Text = Format((CDbl(txtprice.Text) + CDbl(TxtService)) * CDbl(txtQty.Text), "##,##0.#0")
End Sub

Private Sub txtPrice_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then _
          KeyAscii = 0
    If InStr(1, txtprice.Text, ".") > 0 Then If KeyAscii = Asc(".") Then KeyAscii = 0
    If (txtprice & Chr(KeyAscii)) > 9999999999.99999 And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub txtPrice_LostFocus()
Dim z As Double
    If IsNumeric(txtprice.Text) = False Then txtprice.Text = 0
    If txtprice.Text <> "" Then
        z = CDbl(txtprice.Text)
        If z > 9999999999.99999 Then txtprice.Text = Left(z, 10)
    End If
    txtprice.Text = Format(txtprice.Text, "#,##0.00###")
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim TextGrid As String

Baris = Row
With grid
    TextGrid = grid.Text
    If TextGrid = "S" Then
        If Trim(.TextMatrix(Row, 1)) = "" Then cboCls.ListIndex = 1 Else cboCls.ListIndex = 0
        cboCls.Enabled = False
        CboItemCode.Text = .TextMatrix(Row, ColItem)
        CboItemCode.Enabled = False

        cboItemName.Text = .TextMatrix(Row, ColProd)
        cboItemName.Enabled = False
        
        
        
        
        DelDate = Format(.TextMatrix(Row, ColDelDate), "dd mmm yyyy")
        txtQty.Text = Format(.TextMatrix(Row, ColQty), "#,##0.#0")
        'LblQty = CDbl(.TextMatrix(Row, 8))
        cbounit.ListIndex = -1
        For i = 0 To cbounit.ListCount - 1
            If .TextMatrix(Row, 9) = cbounit.List(i, 0) Then
                cbounit.ListIndex = i
                Exit For
            End If
        Next i
        cbocurr.ListIndex = -1
        For i = 0 To cbocurr.ListCount - 1
            If .TextMatrix(Row, ColCurr) = cbocurr.List(i, 0) Then
                cbocurr.ListIndex = i
                Exit For
            End If
        Next i
        cbounit.Text = .TextMatrix(Row, ColUnit)
        cbocurr.Text = .TextMatrix(Row, ColCurr)
        txtprice.Text = Format(.TextMatrix(Row, ColPrice), "#,##0.00###")
        TxtService.Text = Format(.TextMatrix(Row, ColService), "#,##0.00###")
        txtSubTotal = Format(.TextMatrix(Row, ColAmount), "#,##0.00###")
        txtremarks = .TextMatrix(Row, colRemark)
        lblseqno = .TextMatrix(Row, colSeqNo)
        
        
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

Private Sub cbopricecondition_Click()
    If cboPriceCondition.ListIndex <> -1 Then
        
    Else
        
    End If
End Sub

Private Sub cbopricecondition_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call cbopricecondition_Click
End Sub

Private Sub txtpoday_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyReturn Then _
        KeyAscii = 0
    If KeyAscii = Asc("'") Or KeyAscii = Asc(".") Then KeyAscii = 0
End Sub

Private Sub txtpoday_LostFocus()
    
    
End Sub
Function CekFooter(nilai As Boolean)
CekFooter = False
If nilai = True Then
    If cboPaymentTerm.Text = "" Then
    LblErrMsg = DisplayMsg(8123)
    cboPaymentTerm.SetFocus
    CekFooter = True
    Exit Function
    End If
    If cboPriceCondition.Text = "" Then
        LblErrMsg = DisplayMsg(8129)    'Record with This Price Condition not found !
        cboPriceCondition.SetFocus
        CekFooter = True
        Exit Function
     End If
     If cboPriceCondition.Text = "" Then
        LblErrMsg = DisplayMsg(8129)    'Record with This Price Condition not found !
        cboPriceCondition.SetFocus
        CekFooter = True
        Exit Function
      End If
      
      If cboInsuranceCls.Text = "" Then
         LblErrMsg = "INSURANCE CLS" 'DisplayMsg(4147)    'Record with This Price Condition not found !
         cboInsuranceCls.SetFocus
         CekFooter = True
         Exit Function
       End If
       If CboPacking.Text = "" Then
            LblErrMsg = DisplayMsg(34)         'Record with This Price Condition not found !
            CboPacking.SetFocus
            CekFooter = True
            Exit Function
       End If
                
       If cboPriceCondition.Text = "" Then
         LblErrMsg = DisplayMsg(8129)
         cboPriceCondition.SetFocus
         CekFooter = True
         Exit Function
       End If
        If cboTransport.Text = "" Then
           LblErrMsg = DisplayMsg(8130)
           cboTransport.SetFocus
         CekFooter = True
         Exit Function
        End If
        
End If

If cboPaymentTerm.Text <> "" Then
    If cboPaymentTerm.MatchFound = False Then
    LblErrMsg = DisplayMsg(8050)
    cboPaymentTerm.SetFocus
    CekFooter = True
    Exit Function
    End If
    End If
If cboPriceCondition.Text <> "" Then
    If cboPriceCondition.MatchFound = False Then
        LblErrMsg = DisplayMsg(8051)    'Record with This Price Condition not found !
        cboPriceCondition.SetFocus
        CekFooter = True
    Exit Function
    End If
End If
      
      If cboInsuranceCls.Text <> "" Then
       If cboInsuranceCls.MatchFound = False Then
         LblErrMsg = "Record with This Insurance Clas not found " 'DisplayMsg(4147)    'Record with This Price Condition not found !
         cboInsuranceCls.SetFocus
         CekFooter = True
         Exit Function
       End If
      End If
       If CboPacking.Text <> "" Then
         If CboPacking.MatchFound = False Then
            LblErrMsg = DisplayMsg(4010)         'Record with This Price Condition not found !
            CboPacking.SetFocus
            CekFooter = True
            Exit Function
          End If
       End If
                
       
        If cboTransport.Text <> "" Then
         If cboTransport.MatchFound = False Then
           LblErrMsg = DisplayMsg(8059)
            cboTransport.SetFocus
           CekFooter = True
           Exit Function
        End If
        End If


            
End Function


Private Sub Command1_Click(Index As Integer)
Dim sql1 As String, rs1 As New Recordset
Dim sql2 As String, rs2 As New Recordset
Dim tanya, hapus As Boolean
Dim ubahgrid As Boolean
Dim mnudelete As Integer

    ubahgrid = False
    LblErrMsg = ""
    mnudelete = 0
    
    Select Case Index
    Case 0: 'CREATE / UPDATE
            If hakUpdate(Me.Name) = 0 Then _
                LblErrMsg = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
    
            'HEADER VALIDATION
            If cboCust.Text = "" Then
                cboCust.SetFocus
                LblErrMsg = DisplayMsg(1054) '"Please Select Supplier Code"
                Exit Sub
            ElseIf cboCust.Text <> "" Then
                If cboCust.MatchFound = False Then
                    LblErrMsg = DisplayMsg(4050)    'Record with This Supplier Code Not found !
                    cboCust.SetFocus
                    Exit Sub
                End If
            End If
            If txtPoNo.Text = "" Then
                txtPoNo.SetFocus
                LblErrMsg = DisplayMsg(1046) '"Please Input PO No"
                Exit Sub
            End If
            '----------------------------------------------------------
            
            'FOOTER VALIDATION
            If CekFooter(False) Then Exit Sub
            
            '-------------------------------------------------------------
            
            If combo1.ListIndex = 0 Then    'CREATE
                If ubah = False Then
                    sql = "select * from PurchaseOrder_Master where PO_No = '" & txtPoNo.Text & "'"
                    If RS.State <> adStateClosed Then RS.Close
                    RS.Open sql, Db, adOpenKeyset, adLockOptimistic
                    If Not (RS.BOF And RS.EOF) Then
                        LblErrMsg.Caption = DisplayMsg(1023)
                        txtPoNo.SetFocus
                        Exit Sub
                    Else
                        RS.AddNew
                        RS("po_no") = txtPoNo.Text
                        RS("supplier_code") = cboCust.Text
                    End If
                End If
                RS("period") = Format(Period.Value, "yyyymm")
                RS("po_date") = Format(PODate.Value, "yyyy-mm-dd")
                RS("discount") = CDbl(TxtDisc.Text) ' Add 20090112
                RS("amount") = CDbl(txtamount.Text)
                RS("ppn") = CDbl(txtPPN.Text)
                RS("total_amount") = CDbl(txtGrandTotal.Text)
                RS("others_cls") = "1"
                RS("revise_No") = Trim(txtRev.Text)
                
                    
                
                
                
On Error Resume Next
                RS.update
errHandler:
                If InStr(1, err.Description, "Violation of PRIMARY KEY constraint") > 0 Then
                    Call PONO(Right(Year(PODate), 2), Format(Month(PODate), "0#"))
                    txtPONo2.Text = txtPoNo.Text
                    RS("po_No") = txtPoNo.Text
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
                
                combo1.Text = "Update"
                LblErrMsg.Caption = DisplayMsg(1000)
                ubah = True
            
            Else    'UPDATE
                Dim ketemu As Boolean
                
                If cboCust.Text = "" Then
                    cboCust.SetFocus
                    LblErrMsg = DisplayMsg(1054) '"Please Select Supplier Code"
                    Exit Sub
                ElseIf cboCust.Text <> "" Then
                    If cboCust.MatchFound = False Then
                        LblErrMsg = DisplayMsg(4050)    'Record with This Supplier Code Not found !
                        cboCust.SetFocus
                        Exit Sub
                    End If
                End If
                If txtPoNo.Text = "" Then
                    txtPoNo.SetFocus
                    LblErrMsg = DisplayMsg(1046) '"Please Input PO No"
                    Exit Sub
                End If
                
                If cboCust.Text = "" Then
                    CboPOnO.clear: txtPoNo.Text = ""
                Else
                    Call adtocbopono
                End If
                For i = 0 To CboPOnO.ListCount - 1
                    If txtPoNo.Text = CboPOnO.List(i) Then
                        ketemu = True
                        Exit For
                    End If
                Next
                If ketemu = False Then GoTo here
                Call BrowseGrid
                Call Browse
                Call updateMaster(False)
                If ada = False Then
here:
                    Call kosongBwh
                    
                    LblErrMsg.Caption = DisplayMsg(4015)    'Record with this PO No not found
                    txtPoNo.SetFocus
                    Exit Sub
                End If
            End If

    Case 1: 'SUBMIT
            If hakUpdate(Me.Name) = 0 Then _
                LblErrMsg = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
            
            'HEADER VALIDATION
            If cboCust.Text = "" Then
                cboCust.SetFocus
                LblErrMsg = DisplayMsg(1054) '"Please Select Supplier Code"
                Exit Sub
            ElseIf cboCust.Text <> "" Then
                If cboCust.MatchFound = False Then
                    LblErrMsg = DisplayMsg(4050)    'Record with This Supplier Code Not found !
                    cboCust.SetFocus
                    Exit Sub
                End If
            End If
            If txtPoNo.Text = "" Then
                txtPoNo.SetFocus
                LblErrMsg = DisplayMsg(1046) '"Please Input PO No"
                Exit Sub
            End If
            '-------------------------------------------------------------
            
            
            '-------------------------------------------------------------
                
            If mnudelete <> 1 Then Call hitungTotal
            
            sql = "select * from PurchaseOrder_Master where PO_No = '" & txtPoNo.Text & "' and others_cls = '1' and period is not null"
            If RS.State <> adStateClosed Then RS.Close
            RS.Open sql, Db, adOpenKeyset, adLockOptimistic
            If RS.BOF And RS.EOF Then
                LblErrMsg.Caption = DisplayMsg(4015)    'Record with This PO No not found
                txtPoNo.SetFocus
                Exit Sub
            End If
            RS.Close

            If ubah = True Then
                With grid   'DELETE GRID
                    mnudelete = 0
                    For i = 1 To .Rows - 1
                        If .TextMatrix(i, 0) = "D" Then
                            mnudelete = 1
                            If IsEmpty(tanya) Then tanya = MsgBox("Do You Really Want To Delete This Data ?", vbQuestion & vbYesNo, "Confirmation")
                            If tanya = vbYes Then
                                sql1 = "select * from purchaseOrder_Detail where PO_No = '" & txtPoNo.Text & "' " & _
                                       "and POReq_seqNO= " & .TextMatrix(i, colSeqNo)
                                sql = "Delete from PurchaseOrder_Detail where PO_No = '" & txtPoNo.Text & "'  And POReq_seqNO = " & .TextMatrix(i, colSeqNo)

                                    Db.Execute sql
                                    hapus = True
'                                Set rs1 = Db.Execute(sql1)
'                                If Not (rs1.BOF And rs1.EOF) Then
'                                    lblErrMsg.Caption = DisplayMsg(1204)
'                                    .Row = i
'                                    .Col = 0
'
'                                    .SetFocus
'                                    Exit Sub
'                                Else
'                                    Sql = "Delete from PurchaseOrder_Detail where seq_no = '" & .TextMatrix(i, 16) & "'"
'                                    Db.Execute Sql
'                                    hapus = True
'                                End If
                            Else
                                Exit For
                            End If
                        ElseIf .TextMatrix(i, 0) = "S" Then
                            ubahgrid = True
                        End If
                    Next i
                    If mnudelete = 1 Then
                        If (hapus) Then kosongColGrid: Call BrowseGrid: Browse: LblErrMsg = DisplayMsg(1201)
                        For i = 1 To .Rows - 1
                            .TextMatrix(i, 0) = ""
                        Next i
                        Call hitungTotal
                        Exit Sub
                    End If
                End With
                
                
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
                        LblErrMsg = DisplayMsg(8082)    'Please Input product name
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
                    
                    If txtQty.Text = "" Then
                        txtQty.SetFocus
                        LblErrMsg = DisplayMsg(1012) 'Please Input Quantity
                        Exit Sub
                    ElseIf cbounit.Text = "" Then
                        If cbounit.Enabled = True Then cbounit.SetFocus
                        LblErrMsg = DisplayMsg(1030)
                        Exit Sub
                    ElseIf cbocurr.Text = "" Then
                        cbocurr.SetFocus
                        LblErrMsg = DisplayMsg(1028)    'Please Select Currency
                        Exit Sub
                    ElseIf txtprice.Text = "" Then
                        txtprice.SetFocus
                        LblErrMsg = DisplayMsg(1029) '"Please Input Price"
                        Exit Sub
                    ElseIf cboPaymentTerm.Text = "" Then
                        cboPaymentTerm.SetFocus
                        LblErrMsg = DisplayMsg(8123)
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
                    Dim recqty As Double
                    recqty = cekrecqty(lblseqno, txtPoNo.Text)
                    If CDbl(txtQty.Text) < recqty Then
                        txtQty.SetFocus
                        LblErrMsg = DisplayMsg(4036) & " " & recqty '"Quantity must be higher or equal than "
                        Exit Sub
                    End If
                                                     
                    
                    If grid.Rows > 1 Then
                        If grid.TextMatrix(1, ColCurr) <> cbocurr.Column(1) Then
                            cbocurr.SetFocus
                            LblErrMsg = DisplayMsg(4084)    'Cannot select different currency code
                            Exit Sub
                        End If
                    End If
                                                  
                    If txtprice.Text = 0 Then
                        txtprice.SetFocus
                        LblErrMsg = DisplayMsg(1029) '"Please Input Price"
                        Exit Sub
                    ElseIf CDbl(txtprice.Text) > 9999999999.99999 Then
                        txtprice.SetFocus
                        LblErrMsg = DisplayMsg(4048) & " 9,999,999,999.99999" '"Price must be lower or equal than 9,999,999,999.99999"
                        Exit Sub
                    End If
                    '-------------------------------------------------------
                    If CekFooter(True) Then Exit Sub
                    'INSERT PO DETAIL
                    If ubahgrid = False Then
                        sqlGrid = "select * From PurchaseOrder_Detail "
                        Set rsGrid = Nothing
                        If rsGrid.State <> adStateClosed Then rsGrid.Close
                        rsGrid.Open sqlGrid, Db, adOpenKeyset, adLockOptimistic
                        rsGrid.AddNew
                        rsGrid("SEq_No") = seqNo
                    Else    'UPDATE PO DETAIL
                        sqlGrid = "select * From PurchaseOrder_Detail where PoReq_seqNo = '" & lblseqno & "' "
                        If rsGrid.State <> adStateClosed Then rsGrid.Close
                        rsGrid.Open sqlGrid, Db, adOpenKeyset, adLockOptimistic
                    End If
                    
                    rsGrid("PO_No") = Trim(txtPoNo.Text)
                    rsGrid("Item_Code") = Trim(CboItemCode.Text)
                    rsGrid("item_name") = Trim(cboItemName.Text)
                    rsGrid("Delivery_Date") = Format(DelDate, "yyyy-mm-dd")
                    rsGrid("qty") = CDbl(txtQty.Text)
                    rsGrid!Amount = CDbl(txtSubTotal)
                    rsGrid("unit_cls") = cbounit.Column(0)
                    rsGrid("Currency_Code") = cbocurr.Column(0)
                    rsGrid("Price") = CDbl(txtprice.Text)
                    rsGrid!Price_Service = IS_NOL(CDbl(TxtService.Text))
                    'rsGrid!Remarks = txtRemarks
                    rsGrid("Last_Update") = Now()
                    rsGrid("Last_User") = userLogin
                    rsGrid("PORequest_No") = "PO" 'GetNO("SELECT Max(PORequest_No)+1 FROM PurchaseORder_Detail")
                    rsGrid("PoREq_seqNo") = GetNo("SELECT Max(PoREq_seqNo)+1 FROM PurchaseORder_Detail")
                    
                    
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
                    '
                    Call updateMaster(True)
                    'Call CekPONumber

                    Call BrowseGrid
                    Call Browse
                    LblErrMsg = DisplayMsg(1101)
                    ubahgrid = True
                End If
            

    Case 2: 'CLEAR
            Call Kosong
            combo1.ListIndex = 1
            Call Combo1_Click
            cboCust.SetFocus

    Case 3: 'CANCEL
            If txtPoNo.Text <> "" And cboCust.Text <> "" Then
                For i = 0 To CboPOnO.ListCount - 1
                    If txtPoNo.Text = CboPOnO.List(i) Then
                        ketemu = True
                        Exit For
                    End If
                Next i
                If ketemu = False Then
                    Call kosongBwh
                    'Call header
                    LblErrMsg.Caption = DisplayMsg(4015)    'Record with This PO No not found
                    txtPoNo.SetFocus
                    Exit Sub
                End If
                
                Call BrowseAtas
                Call Browse
            End If
    End Select
End Sub

Private Sub cmdReport_Click()
Dim application As New CRAXDDRT.application
Dim report As New CRAXDDRT.report
Dim rsRpt As New ADODB.Recordset
Dim Rpt As New FrmRpt3
Dim sqlcekdet As String, SqlRpt As String
Dim rscekdet As New Recordset
  
    If combo1.ListIndex = 1 And txtPoNo.Text <> "" And cboCust.Text <> "" Then
        sqlcekdet = "select pom.PO_No from PurchaseOrder_Master pom " & _
                    "inner join PurchaseOrder_Detail pod on pod.PO_No = pom.PO_No " & _
                    "where pom.others_cls = '1' and pom.period is not null " & _
                    "and pom.PO_No = '" & Trim(txtPoNo.Text) & "' and pom.supplier_Code = '" & Trim(cboCust.Text) & "'"
        Set rscekdet = Db.Execute(sqlcekdet)
        If rscekdet.EOF Then LblErrMsg.Caption = DisplayMsg(4006): Me.MousePointer = vbDefault: Exit Sub
        Set rscekdet = Nothing
        
        Me.MousePointer = vbHourglass

'        SqlRpt = "select '1' Sort, rtrim(pom.po_no) po_no, pom.po_date, rtrim(pom.supplier_Code) Supplier_Code, " & _
'                    " trade_name = case when CHARINDEX ( ',' , rtrim(tm.trade_name)) <> 0 then " & _
'                    " ltrim(substring(rtrim(tm.trade_name), CHARINDEX ( ',' , rtrim(tm.trade_name)) + 1, 5)) + " & _
'                    " '.' + left(rtrim(tm.trade_name),CHARINDEX ( ',' , rtrim(tm.trade_name)) - 1) " & _
'                    " Else " & _
'                    " RTrim (tm.trade_name) " & _
'                    " End, "
'            SqlRpt = SqlRpt + " " & _
'                 "rtrim(tm.address1) taddress1, rtrim(tm.address2) taddress2, rtrim(tm.city) tcity, rtrim(tm.postal_code) tpostal_code, " & _
'                 "rtrim(tm.contact_person) contact_person, isnull(tm.telephone,'') Supplierphone, isnull(tm.Fax,'') SupplierFax, pom.POPayment_code, pom.POPayment_Days, pom.POPayment_Terms, " & _
'                 "rtrim(pod.PORequest_No) PORequest_No, pod.POReq_SeqNo, rtrim(pod.item_code) item_code, rtrim(pod.item_name) item_name, " & _
'                 "pod.unit_cls, (select description from unit_cls uc where uc.unit_cls= pod.unit_cls ) unit_desc, isnull(pod.qty,0) Qty, pod.currency_code, (select description from curr_cls where curr_cls.Curr_cls= pod.Currency_Code) Curr_desc ,isnull(pod.price,0) Price, isnull(pod.amount,0) Amount, " & _
'                 "pod.Delivery_Date, pom.PriceCondition_Cls, rtrim(pc.description) PriceCondition, pom.Transportation_Cls, " & _
'                 "rtrim(pom.remarks) Remarks, rtrim(pom.remarks2) Remarks2, rtrim(pom.remarks3) Remarks3, isnull(pom.amount,0) as TAmount, isnull(pom.ppn,0) PPN, isnull(pom.total_amount,0) Total_Amount, " & _
'                 "rtrim(cp.company_name) company_name, rtrim(cp.address1) caddress1, rtrim(cp.address2) caddress2, " & _
'                 "rtrim(cp.Province) cprovince, rtrim(cp.City) ccity, rtrim(cp.postal_code) cpostal_code, rtrim(cp.phone1) cphone1, " & _
'                 "rtrim(cp.phone2) cphone2, rtrim(cp.fax) cfax, rtrim(cp.PO_position) po_position, rtrim(cp.PO_person) po_person, " & _
'                 "rtrim(cp.POAcknowledge_Person) POAcknowledge_Person, rtrim(cp.POAcknowledge_Position) POAcknowledge_Position, " & _
'                 "rtrim(cp.POApproved_Person) POApproved_Person, rtrim(cp.POApproved_Position) POApproved_Position, tm.Trade_Cls, tm.Country_Cls, rtrim(pod.Department_cls) Department_Cls " & _
'                 "from PurchaseOrder_Master pom " & _
'                 "inner join PurchaseOrder_Detail pod on pod.PO_No = pom.PO_No " & _
'                 "left outer join Trade_Master tm on tm.trade_code = pom.supplier_code " & _
'                 "left outer join PriceCondition_Cls pc on pc.PriceCondition_Cls = pom.PriceCondition_Cls " & _
'                 "cross join Company_Profile cp " & _
'                 "where pom.po_no = '" & Trim(txtpono.Text) & "' and pom.others_cls = '1' and pom.period is not null " & _
'                 "order by pod.Item_Name"

' -----
' Perubahan sesuai Format Musahsi
' -----

SqlRpt = " Select POM.Po_No, POM.Po_Date,POM.delivery_Date,PRD.PoRequest_No,PRM.PersonInCharge_Cls,PIC.Description, " & _
            " POM.Supplier_Code,TM.Trade_Name,TM.Contact_Person,TM.Address1,TM.Address2,TM.City,TM.Country, " & _
            " TM.Telephone,Tm.Fax,POM.PaymentTerm_Cls, " & _
            " POD.Item_code,POD.Price,POD.Qty,POD.Amount,isnull(IM.Item_Name,POD.Item_Name), " & _
            " POD.Unit_Cls,U.Description Unit,POD.Currency_Code,C.Description Currency," & _
            "' ' Ref,' ' ShipVia,POM.Remarks comments "
            
'            " isnull((Select ChildRequirement_Qty from requirement_Master Where ChildRequirement_Month=month(POM.delivery_Date)+1 and ChildRequirement_Year=year(POM.delivery_Date) and ChildItem_Code=POD.Item_code),0) F1, " & _
'            " isnull((Select ChildRequirement_Qty from requirement_Master Where ChildRequirement_Month=month(POM.delivery_Date)+2 and ChildRequirement_Year=year(POM.delivery_Date) and ChildItem_Code=POD.Item_code),0) F2 "

SqlRpt = SqlRpt & _
            " From PurchaseOrder_Master POM inner join PurchaseOrder_Detail POD " & _
            " On POM.Po_No=POd.Po_no " & _
            " Inner Join Trade_Master TM on POM.Supplier_Code=TM.Trade_Code " & _
            " left Join Item_Master IM on POD.Item_Code=IM.Item_Code " & _
            " inner Join Unit_Cls U on POD.Unit_Cls=U.Unit_Cls " & _
            " inner Join PORequest_Detail PRD on POD.PORequest_No=PRD.PoRequest_No and POD.PoReq_SeqNo=PRD.PoReq_SeqNo " & _
            " inner Join PoREquest_Master PRM on POD.PORequest_No=PRM.PoRequest_No " & _
            " inner join PersonInCharge_Cls PIC on PRM.PersonInCharge_Cls=PIC.PersonInCharge_Cls " & _
            " inner join curr_cls C on POD.Currency_Code=C.Curr_Cls " & _
            " where pom.po_no = '" & Trim(txtPoNo.Text) & "' and pom.others_cls = '1' and pom.period is null " & _
            " order by pod.PORequest_No, pod.Item_Code, pod.POReq_SeqNo "
            
' -----

        If rsRpt.State <> adStateClosed Then rsRpt.Close
        rsRpt.Open SqlRpt, Db, adOpenDynamic, adLockOptimistic
        
        sqlprint = SqlRpt
        'reportcode = "poparts"
        reportcode = "PoOthers"
        Fbulan = txtPoNo.Text
        printorient = 1
        
        If rsRpt.EOF Then LblErrMsg.Caption = DisplayMsg(4006): Me.MousePointer = vbDefault: Exit Sub
        Set report = application.OpenReport(App.path & "\Reports\rptPOOtherNew.rpt")
        report.Database.Tables(1).SetDataSource rsRpt
        
        Rpt.CRViewer1.ReportSource = report
        Rpt.CRViewer1.ViewReport
        Rpt.CRViewer1.Zoom 1
        Rpt.WindowState = 2
        Rpt.Show 1
        
        Set rsRpt = Nothing
        Me.MousePointer = vbDefault
    End If
End Sub

Private Sub CmdSubMenu_Click()
    Call ClearData
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

Private Sub hscrollbar_Change()
 End Sub

Private Sub hscrollbar_Scroll()
    Call hscrollbar_Change
End Sub

Private Sub updateMaster(Flag As Boolean)
    Dim sQl_Master As String
    Dim rs_Master As New ADODB.Recordset
    
            sQl_Master = "select * from PurchaseOrder_Master where PO_No = '" & txtPoNo.Text & "' and others_cls = '1' and period is not null"
            If rs_Master.State <> adStateClosed Then rs_Master.Close
            rs_Master.Open sQl_Master, Db, adOpenKeyset, adLockOptimistic
            If rs_Master.BOF And rs_Master.EOF Then
                LblErrMsg.Caption = DisplayMsg(4015)    'Record with This PO No not found
                txtPoNo.SetFocus
                rs_Master.Close
                Set rs_Master = Nothing
                Exit Sub
            End If
    rs_Master("period") = Format(Period.Value, "yyyymm")
    rs_Master("revise_No") = Trim(txtRev.Text)
    rs_Master("po_date") = Format(PODate.Value, "yyyy-mm-dd")
    rs_Master("supplier_Code") = Trim(cboCust.Text)
    
    
    If Flag = True Then
        rs_Master("discount") = CDbl(TxtDisc.Text)
        rs_Master("amount") = CDbl(txtamount.Text)
        rs_Master("ppn") = CDbl(txtPPN.Text)
        rs_Master("total_amount") = CDbl(txtGrandTotal.Text)
        rs_Master("others_cls") = "1"
        rs_Master("pricecondition_cls") = is_null(cboPriceCondition.Text)
        rs_Master!PaymentTerm_Cls = is_null(cboPaymentTerm.Text)
        rs_Master!Insurance_Cls = is_null(Trim(cboInsuranceCls.Text))
        rs_Master!Transportation_Cls = is_null(Trim(cboTransport.Text))
        rs_Master!POPacking_Cls = is_null(CboPacking.Text)
        rs_Master!POMarking1 = is_null(Trim(txtMarking(0)))
        rs_Master!POMarking2 = is_null(Trim(txtMarking(1)))
        rs_Master!POMarking3 = is_null(Trim(txtMarking(2)))
        rs_Master!POMarking4 = is_null(Trim(txtMarking(3)))
        rs_Master!POMarking5 = is_null(Trim(txtMarking(4)))
        rs_Master!POMarking6 = is_null(Trim(txtMarking(5)))
        
        
        rs_Master!Remarks = is_null(Trim(txtremarks))
    End If
    rs_Master.update
    rs_Master.Close
    Set rs_Master = Nothing
End Sub
Function is_null(Data)
If IsNull(Data) Then
is_null = ""
Exit Function
End If
If Data = "" Then
is_null = Null
Else
is_null = Trim(Data)
End If

End Function
Private Sub hitungTotal()
            Dim a As Double
            a = 0
            For i = 1 To grid.Rows - 1
                If Trim(grid.TextMatrix(i, 0)) <> "D" Then a = a + IS_NOL(grid.TextMatrix(i, ColAmount))
            Next i
            ' Add 20090112
            TxtSubAmount.Text = a
            If (TxtSubAmount.Text <> 0) Then TxtSubAmount.Text = Format(TxtSubAmount.Text, "##,##0.#0")
            
            txtamount.Text = CDbl(TxtSubAmount) - CDbl(TxtDisc)
            If (txtamount.Text <> 0) Then txtamount.Text = Format(txtamount.Text, "##,##0.#0")
            ' ---
            
            If cboCust.Column(3) = "4" Or cboCust.Column(3) = "6" Then
                txtPPN.Text = CDbl(isippn / 100) * CDbl(txtamount.Text)
            Else
                txtPPN.Text = 0
            End If
            If (txtPPN.Text <> 0) Then txtPPN.Text = Format(txtPPN.Text, "##,##0.#0")
            txtGrandTotal = CDbl(txtPPN.Text) + CDbl(txtamount.Text)
            If (txtGrandTotal.Text <> 0) Then txtGrandTotal.Text = Format(txtGrandTotal.Text, "##,##0.#0")
End Sub

Private Sub txtService_Change()
    If InStr(1, TxtService.Text, ",") = 1 Then TxtService.Text = Right(TxtService, Len(TxtService) - 1)
    If txtQty.Text <> "" And txtprice.Text <> "" And IsNumeric(txtQty) And IsNumeric(txtprice) And TxtService.Text <> "" And IsNumeric(TxtService) Then txtSubTotal.Text = Format((CDbl(txtprice.Text) + CDbl(TxtService)) * CDbl(txtQty.Text), "##,##0.#0")

End Sub

Private Sub txtService_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then _
          KeyAscii = 0
    If InStr(1, TxtService.Text, ".") > 0 Then If KeyAscii = Asc(".") Then KeyAscii = 0
    If (TxtService & Chr(KeyAscii)) > 9999999999.99999 And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub TxtService_LostFocus()
Dim z As Double
    If IsNumeric(TxtService.Text) = False Then TxtService.Text = 0
    If TxtService.Text <> "" Then
        z = CDbl(TxtService.Text)
        If z > 9999999999.99999 Then TxtService.Text = Left(z, 10)
    End If
    TxtService.Text = Format(TxtService.Text, "#,##0.00###")
End Sub

