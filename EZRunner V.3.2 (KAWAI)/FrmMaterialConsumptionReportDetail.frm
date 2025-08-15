VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmMaterialConsumptionReportDetail 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Material Consumption Report Detail"
   ClientHeight    =   10440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   Icon            =   "FrmMaterialConsumptionReportDetail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10440
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar Pb01 
      Height          =   375
      Left            =   480
      TabIndex        =   25
      Top             =   8640
      Width           =   14355
      _ExtentX        =   25321
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
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
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9840
      Width           =   1185
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   390
      TabIndex        =   14
      Top             =   9225
      Width           =   14595
      Begin VB.Label LblErr 
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
         Height          =   255
         Left            =   105
         TabIndex        =   15
         Top             =   195
         Width           =   14370
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1335
      Left            =   360
      TabIndex        =   10
      Top             =   1245
      Width           =   14595
      Begin VB.TextBox lblGroup 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDDFE3&
         BorderStyle     =   0  'None
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
         Left            =   3405
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   855
         Width           =   3510
      End
      Begin VB.TextBox Lblsupp 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDDFE3&
         BorderStyle     =   0  'None
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
         Left            =   3420
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   300
         Width           =   3510
      End
      Begin VB.CommandButton cmdsearch 
         BackColor       =   &H0080FFFF&
         Caption         =   "Sea&rch"
         Height          =   375
         Left            =   12900
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   780
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox lblAddr 
         BackColor       =   &H00FDDFE3&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   8190
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   300
         Width           =   5835
      End
      Begin MSComCtl2.DTPicker Tgl1 
         Height          =   315
         Left            =   8940
         TabIndex        =   2
         Top             =   840
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
         Format          =   293601283
         UpDown          =   -1  'True
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker Tgl2 
         Height          =   315
         Left            =   10950
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   840
         Width           =   1305
         _ExtentX        =   2302
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
         CustomFormat    =   "MMM yyyy"
         Format          =   293601283
         UpDown          =   -1  'True
         CurrentDate     =   37798
      End
      Begin VB.Label Label3 
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
         Index           =   2
         Left            =   10500
         TabIndex        =   22
         Top             =   900
         Width           =   165
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Price Periode"
         Height          =   255
         Index           =   0
         Left            =   7335
         TabIndex        =   20
         Top             =   915
         Width           =   1080
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   255
         Index           =   4
         Left            =   7335
         TabIndex        =   19
         Top             =   330
         Width           =   840
      End
      Begin VB.Label LblPart 
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
         Left            =   285
         TabIndex        =   18
         Top             =   915
         Width           =   855
      End
      Begin VB.Line Line1 
         X1              =   3405
         X2              =   6905
         Y1              =   1140
         Y2              =   1140
      End
      Begin MSForms.ComboBox CboGroupCls 
         Height          =   315
         Left            =   1785
         TabIndex        =   1
         Top             =   795
         Width           =   1515
         VariousPropertyBits=   746604571
         MaxLength       =   35
         DisplayStyle    =   3
         Size            =   "2672;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label LblPart 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Code "
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
         Left            =   285
         TabIndex        =   13
         Top             =   330
         Width           =   1275
      End
      Begin MSForms.ComboBox CboSupplier 
         Height          =   315
         Left            =   1785
         TabIndex        =   0
         Top             =   255
         Width           =   1515
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2672;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Line Line4 
         X1              =   3420
         X2              =   6920
         Y1              =   585
         Y2              =   585
      End
      Begin VB.Line Line2 
         Index           =   1
         X1              =   8175
         X2              =   14055
         Y1              =   585
         Y2              =   585
      End
   End
   Begin VB.CommandButton cmdReport 
      BackColor       =   &H0080FFFF&
      Caption         =   "E&xcel"
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
      Left            =   13785
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9840
      Width           =   1185
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
      Left            =   420
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9855
      Width           =   1140
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   13110
      TabIndex        =   8
      Top             =   240
      Width           =   1845
      _extentx        =   3254
      _extenty        =   714
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   4905
      Left            =   360
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3240
      Width           =   7170
      _cx             =   12647
      _cy             =   8652
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
   Begin MSComDlg.CommonDialog cdlReport 
      Left            =   1380
      Top             =   1845
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VSFlex8Ctl.VSFlexGrid GridSelect 
      Height          =   4905
      Left            =   7740
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3240
      Width           =   7170
      _cx             =   12647
      _cy             =   8652
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
   Begin VB.Label Label1 
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
      Left            =   360
      TabIndex        =   26
      Top             =   8280
      Width           =   60
   End
   Begin VB.Shape Shape2 
      Height          =   615
      Left            =   360
      Top             =   8520
      Width           =   14595
   End
   Begin VB.Label LblPart 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Selected Finish Good"
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
      Left            =   7860
      TabIndex        =   24
      Top             =   2880
      Width           =   1785
   End
   Begin VB.Label LblPart 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Available Finish Good"
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
      Left            =   480
      TabIndex        =   23
      Top             =   2880
      Width           =   1830
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Material Consumption Report Detail"
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
      Left            =   5580
      TabIndex        =   9
      Top             =   315
      Width           =   4125
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00A6D2FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   0
      Left            =   7740
      Top             =   2760
      Width           =   7155
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00A6D2FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   2
      Left            =   360
      Top             =   2760
      Width           =   7155
   End
End
Attribute VB_Name = "FrmMaterialConsumptionReportDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bteColSelect As Byte
Dim bteColItem_Code As Byte
Dim bteColDescription As Byte
Dim bteColGroup As Byte

Dim StrTableTemp As String
Dim adoRs As New ADODB.Recordset
Dim BrsExel As Integer
Dim tgl_sb As String * 2

Sub CekMark()
    Dim Penuh As Boolean
    With grid
        If CboGroupCls = strAll Then
            For i = 1 To .Rows - 1
                If .Cell(flexcpChecked, i, Col) = 2 Then
                    Penuh = False
                    Exit For
                Else
                    Penuh = True
                End If
            Next i
            If Penuh = False Then
                .Cell(flexcpChecked, 0, bteColSelect) = 2
            ElseIf Penuh = True Then
                .Cell(flexcpChecked, 0, bteColSelect) = 1
            End If
        Else
            For i = 1 To .Rows - 1
                If .Cell(flexcpChecked, i, Col) = 2 And .TextMatrix(i, bteColGroup) = CboGroupCls Then
                    Penuh = False
                    Exit For
                Else
                    Penuh = True
                End If
            Next i
            If Penuh = False Then
                .Cell(flexcpChecked, 0, bteColSelect) = 2
            ElseIf Penuh = True Then
                .Cell(flexcpChecked, 0, bteColSelect) = 1
            End If
        End If
    End With
End Sub
Sub HideData()
    Dim X As Integer
    For X = 1 To grid.Rows - 1
        If grid.TextMatrix(X, bteColGroup) = CboGroupCls Then
            grid.RowHidden(X) = False
        Else
            grid.RowHidden(X) = True
        End If
    Next X
End Sub

Sub ShowAllData()
    Dim X As Integer
    For X = 1 To grid.Rows - 1
        grid.RowHidden(X) = False
    Next X
End Sub

Sub SelectGrid()
    Dim X As Integer
    Call HeaderGridSelect
    For X = 1 To grid.Rows - 1
        If grid.Cell(flexcpChecked, X, bteColSelect) = 1 Then
            GridSelect.AddItem ""
            GridSelect.TextMatrix(GridSelect.Rows - 1, bteColItem_Code) = grid.TextMatrix(X, bteColItem_Code)
            GridSelect.TextMatrix(GridSelect.Rows - 1, bteColDescription) = grid.TextMatrix(X, bteColDescription)
        End If
    Next
    
End Sub

Sub adtocbosupplier()
    Dim RssSupplier As New ADODB.Recordset
    Dim sqlcust As String

    sqlcust = "SELECT  rtrim(Trade_Master.trade_Code) supp_code, rtrim(Trade_Master.Trade_Name) supp_name, " & _
        "rtrim(Trade_Master.Address1) address, country_Cls, POPayment_Day From Trade_Master where trade_cls in ('2','3')"
    Set RssSupplier = Db.Execute(sqlcust)

    With cboSupplier
        .clear
        .columnCount = 4
        .ColumnWidths = "80 pt;280 pt; 0 pt; 0 pt; 0 pt"
        .ListWidth = 360
        .ListRows = 15
        .AddItem ""
        .List(0, 0) = strAll
        .List(0, 1) = strAll
        .List(0, 2) = ""
        
        i = 1
        RssSupplier.Requery
        If Not RssSupplier.EOF And Not RssSupplier.BOF Then
            Do Until RssSupplier.EOF
                .AddItem ""
                .List(i, 0) = IIf(IsNull(Trim(RssSupplier!supp_code)), "", Trim(RssSupplier!supp_code))
                .List(i, 1) = IIf(IsNull(Trim(RssSupplier!supp_name)), "", Trim(RssSupplier!supp_name))
                .List(i, 2) = IIf(IsNull(Trim(RssSupplier!Address)), "", Trim(RssSupplier!Address))
                .List(i, 3) = IIf(IsNull(Trim(RssSupplier!country_cls)), "", Trim(RssSupplier!country_cls))
                .List(i, 4) = IIf(IsNull(Val(RssSupplier!POPayment_Day & "")), "", Val(RssSupplier!POPayment_Day & ""))
                i = i + 1
                RssSupplier.MoveNext
            Loop
        End If
        .ListIndex = 0
    End With
    Set RssSupplier = Nothing
    
    
End Sub

Sub adtocboGroup()
Dim RSSGroup As New ADODB.Recordset
    
    
    RSSGroup.Open "select * from group_cls", Db, adOpenKeyset, adLockOptimistic
    With CboGroupCls
        .clear
        .columnCount = 2
        .ColumnWidths = "50 pt;75 pt"
        .ListWidth = 180
        .ListRows = 15
        .AddItem ""
        .List(0, 0) = strAll
        .List(0, 1) = strAll
        
    i = 1
    Do Until RSSGroup.EOF
        .AddItem ""
        .List(i, 0) = Trim(RSSGroup!group_cls)
        .List(i, 1) = Trim(RSSGroup!Description)
        i = i + 1
        RSSGroup.MoveNext
    Loop
    .ListIndex = 0
    End With

End Sub

Private Sub CboGroupCls_Change()
    Call CboGroupCls_Click
End Sub

Private Sub CboGroupCls_Click()

        LblErr.Caption = ""
            
        If CboGroupCls.ListIndex <> -1 Then
            lblgroup.Text = CboGroupCls.Column(1)
        End If
        If CboGroupCls = strAll Then
            Call ShowAllData
        Else
            Call HideData
        End If
        Call CekMark
End Sub

Private Sub CboGroupCls_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call CboGroupCls_Click
End Sub

Private Sub CboGroupCls_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = Asc("'") Then KeyAscii = 0
    Call CboGroupCls_Click
End Sub

Private Sub CboGroupCls_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If Trim(CboGroupCls) = "" Then lblgroup.Text = ""
End Sub

Private Sub CboGroupCls_LostFocus()
    If Trim(CboGroupCls) = "" Then LblErr = "": Exit Sub
    If CboGroupCls.MatchFound = False Then
        lblgroup.Text = ""
        
        LblErr = DisplayMsg("4064")
   
        Exit Sub
    Else
         LblErr = ""
    End If
    
    Call CboGroupCls_Click
End Sub

Private Sub CboSupplier_Change()
    'Call header
End Sub

Private Sub cbosupplier_Click()
        LblErr.Caption = ""
            
        If cboSupplier.ListIndex <> -1 Then
            lblSupp.Text = cboSupplier.Column(1)
            lblAddr.Text = cboSupplier.Column(2)
        End If
End Sub

Private Sub cbosupplier_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then Call cbosupplier_Click
End Sub

Private Sub cboSupplier_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = Asc("'") Then KeyAscii = 0
    Call cbosupplier_Click
End Sub

Private Sub CboSupplier_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If Trim(cboSupplier) = "" Then
            lblSupp.Text = ""
            lblAddr.Text = ""
    End If
End Sub

Private Sub CboSupplier_LostFocus()
    If Trim(cboSupplier) = "" Then LblErr = "": Exit Sub
    If cboSupplier.MatchFound = False Then
        lblSupp.Text = ""
        lblAddr.Text = ""
        LblErr = DisplayMsg("0032")
   
        Exit Sub
    Else
         LblErr = ""
    End If
    
    Call cbosupplier_Click
End Sub

Private Sub cmdClear_Click()
    Call adtocboGroup
    Call adtocbosupplier
    Call Header
    Call ClearData
End Sub

Sub SettingExel()
    On Error GoTo ErExcl
    Dim rsRpt As New ADODB.Recordset
    Dim xlapp As New Excel.application
    Dim Baris As Integer
    Dim RsItem As New ADODB.Recordset
    Dim RsQty As New ADODB.Recordset
    Dim sql As String
    Dim Urut As Integer, strSQL As String, X As Integer
    
    Dim bln As Integer, thn As Integer
    Dim bulan(6) As Integer, tahun(6) As String
    Dim BlnAwal As Date, blnAkhir As Date
    
    Dim ParentCode(500) As String
    Dim IntIndex As Integer, maxCol As Integer
    
    IntIndex = 0
    
    Label1 = "Check available Data ... "
    
    BrsExel = 9
    LblErrMsg = ""
    Label1 = ""
    
    MousePointer = vbHourglass
            
            sql = "select rtrim(company_name) company_name, rtrim(address1) Address1, rtrim(Address2) Address2, rtrim(Province) Province, rtrim(city) City, Rtrim(Postal_Code) Postal_Code, Rtrim(phone1) Phone1, Rtrim(phone2) Phone2, rtrim(fax) Fax  From company_profile "
            If rsRpt.State <> adStateClosed Then rsRpt.Close
            rsRpt.Open sql, Db, adOpenDynamic, adLockOptimistic
            If rsRpt.EOF Then Me.MousePointer = vbDefault: Exit Sub
    Baris = 10
    'HEys
    With xlapp
        .Workbooks.Add
        .Range("a2", "m2").Merge
        .Range("a2") = Trim(rsRpt!company_name)
        .Range("a2").HorizontalAlignment = xlLeft
        .Range("a2").Columns.Font.Name = "Arial"
        .Range("a2").Columns.Font.Size = 10
        .Range("A2").Font.Bold = True
    
        .Range("a3", "m3").Merge
        .Range("a3", "m3") = Trim(rsRpt!address1) & " " & Trim(rsRpt!address2) & " " & Trim(rsRpt!City) & " " & Trim(rsRpt!Province) & " " & Trim(rsRpt!postal_code)
        .Range("a3").HorizontalAlignment = xlLeft
        .Range("a3").Columns.Font.Name = "Arial"
        .Range("a3").Columns.Font.Size = 10
        .Range("A3").Font.Bold = True
        
        .Range("a4", "m4").Merge
        .Range("a4") = "Phone: " & Trim(rsRpt!phone1) & " " & Trim(rsRpt!phone2) & " Fax: " & Trim(rsRpt!fax)
        .Range("a4").HorizontalAlignment = xlLeft
        .Range("a4").Columns.Font.Name = "Arial"
        .Range("a4").Columns.Font.Size = 10
        .Range("A4").Font.Bold = True
        
        .Range("a6", "m6").Merge
        .Range("a6") = Trim(lblJudul)
        .Range("a6").HorizontalAlignment = xlLeft
        .Range("a6").Columns.Font.Name = "Arial"
        .Range("a6").Columns.Font.Size = 9
        .Range("a6").Columns.Font.Bold = True
        
        'header Table
        
        .Range("A8", "A9").Merge
        .Range("A8") = "No"
        .Range("A8").HorizontalAlignment = xlCenter
    
        .Range("B8", "B9").Merge
        .Range("B8") = "Part ID"
        .Range("B8").HorizontalAlignment = xlCenter
        
        .Range("C8", "C9").Merge
        .Range("C8") = "Description"
        .Range("C8").HorizontalAlignment = xlCenter
        
        .Range("D8", "D9").Merge
        .Range("D8") = "Supplier"
        .Range("D8").HorizontalAlignment = xlCenter
        
    Dim brsN, klmN As Integer
    brsN = 8
    klmN = 5
    
    BlnAwal = Tgl1
    For X = 0 To 5
        .Cells(brsN, klmN) = Format(BlnAwal, "MMM YYYY")
        .Range(.Cells(brsN, klmN), .Cells(brsN, klmN + 1)).Merge
        .Cells(brsN + 1, klmN) = "Price "
        .Range(.Cells(brsN + 1, klmN), .Cells(brsN + 1, klmN + 1)).Merge
        BlnAwal = DateAdd("M", 1, BlnAwal)
        klmN = klmN + 2
    Next X
    
    For Gr = 1 To GridSelect.Rows - 1
        
            'If GridSelect.Cell(flexcpChecked, Gr, BteColSelect) = 1 Then
                ParentCode(IntIndex) = GridSelect.TextMatrix(Gr, bteColItem_Code)
                IntIndex = IntIndex + 1
                .Cells(brsN, klmN) = GridSelect.TextMatrix(Gr, bteColItem_Code)
                .Range(.Cells(8, klmN), .Cells(8, klmN + 1)).Font.Bold = True
                .Range(.Cells(8, klmN), .Cells(8, klmN + 1)).HorizontalAlignment = xlCenter
                .Cells(brsN + 1, klmN) = GridSelect.TextMatrix(Gr, bteColDescription)
                .Range(.Cells(9, klmN), .Cells(9, klmN + 1)).Font.Bold = True
                .Range(.Cells(9, klmN), .Cells(9, klmN + 1)).HorizontalAlignment = xlCenter
                klmN = klmN + 1
            'End If
    
    Next Gr
        
    
    Urut = 1
    brsN = 10
    
    strSQL = " Select Item_Code, RTrim(Item_Name) Item_Name, RTrim(Trade_Name) Trade_Name,  " & vbCrLf
                      
    bln = Month(Tgl1.Value) - 1
    thn = Year(Tgl1.Value)
    BlnAwal = Tgl1
    
    For X = 0 To 5
        bln = bln + 1
    
        If bln > 12 Then
            bln = 1: thn = thn + 1
        End If
        bulan(X) = bln
        tahun(X) = thn
        
        blnAkhir = DateAdd("M", 1, BlnAwal)
        
        strSQL = strSQL & "   RTrim(isnull((Select top 1 Description From Price_Master Inner Join Curr_Cls On Price_Master.Currency_Code=Curr_Cls.Curr_Cls " & vbCrLf & _
                          "         Where Item_Code=A.Item_Code And Trade_Code=A.Supplier_Code  And Start_Date<='" & Format(BlnAwal, "YYYYMM01") & "' And End_Date>='" & Format(blnAkhir, "YYYYMM01") & "'  " & vbCrLf & _
                          "         and Price_Cls='01'),'')) Curr" & X & ",    " & vbCrLf & _
                          "      isnull((Select top 1 Price From Price_Master Where Item_Code=A.Item_Code And Trade_Code=A.Supplier_Code " & vbCrLf & _
                          "         And Start_Date<='" & Format(BlnAwal, "YYYYMM01") & "' And End_Date>='" & Format(blnAkhir, "YYYYMM01") & "' and Price_Cls='01'),0) Price" & X & ",    " & vbCrLf
        
        BlnAwal = blnAkhir
        
    Next
    
    For X = 0 To IntIndex - 1
                    strSQL = strSQL & "     (select Sum(Qty) From " & StrTableTemp & vbCrLf & _
                      "         Where Item_Code=a.Item_Code And Parent_Item ='" & Trim(ParentCode(X)) & " ' ) [" & Trim(ParentCode(X)) & " ], " & vbCrLf
    Next X
    
    strSQL = strSQL & "   MakeBuy_Cls  From Item_Master A Inner Join Trade_Master B On A.Supplier_Code=B.Trade_Code " & vbCrLf & _
                        "         Where MakeBuy_Cls='02' " & vbCrLf
                        
    If Trim(cboSupplier) <> strAll Then strSQL = strSQL & " And A.Supplier_Code='" & Trim(cboSupplier) & "' "
    
    strSQL = strSQL & "  Order By Trade_Name, A.Item_Code "
    
    If RsItem.State <> adStateClosed Then RsItem.Close
    RsItem.Open strSQL, Db, 3, 2
    
    Pb01.Max = RsItem.RecordCount
    
    Do While Not RsItem.EOF
        .Cells(brsN, 1) = Urut
        .Cells(brsN, 2) = Trim(RsItem("Item_Code"))
        .Cells(brsN, 3) = RsItem("Item_Name")
        .Cells(brsN, 4) = RsItem("Trade_Name")
        
        For X = 3 To 14 Step 2
            .Cells(brsN, X + 2) = RsItem.Fields(X)
            
            If RsItem.Fields(X) <> "" Then
                .Cells(brsN, X + 3) = RsItem.Fields(X + 1)
            End If
            
            If Trim(RsItem.Fields(X)) <> "IDR" Then
                .Range(.Cells(brsN, X + 3), .Cells(brsN, X + 3)).NumberFormat = gs_formatAmount
            Else
                .Range(.Cells(brsN, X + 3), .Cells(brsN, X + 3)).NumberFormat = gs_formatAmountIDR
            End If
        Next X
        
        For X = 15 To RsItem.Fields.Count - 2
            .Cells(brsN, X + 2) = RsItem.Fields(X)
            maxCol = X
        Next X
    
    brsN = brsN + 1
    Urut = Urut + 1
    
    Label1 = "Transfer Progress ... " '& Round((RsItem.AbsolutePosition / Pb01.Max) * 100, 0) & " %"

    Pb01.Value = Pb01.Value + 1
    RsItem.MoveNext
    
    Loop
    
    RsItem.Close
    
       .Range("B1", "B" & brsN).HorizontalAlignment = xlLeft
       .Range(.Cells(8, 1), .Cells(8, klmN - 1)).VerticalAlignment = xlCenter
    
       .Range(.Cells(8, 1), .Cells(8, klmN - 1)).Borders(xlEdgeTop).LineStyle = xlContinuous
       .Range(.Cells(9, 1), .Cells(9, klmN - 1)).Borders(xlEdgeBottom).LineStyle = xlDouble
       
       .Range(.Cells(8, 1), .Cells(brsN, klmN - 1)).Borders(xlInsideVertical).LineStyle = xlContinuous
       .Range(.Cells(8, 1), .Cells(brsN, klmN - 1)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
       .Range(.Cells(8, 1), .Cells(brsN, klmN - 1)).Borders(xlEdgeBottom).LineStyle = xlContinuous
       .Range(.Cells(8, 1), .Cells(brsN, klmN - 1)).Borders(xlEdgeLeft).LineStyle = xlContinuous
       .Range(.Cells(8, 1), .Cells(brsN, klmN - 1)).Borders(xlEdgeRight).LineStyle = xlContinuous
       
       .Range(.Cells(9, 17), .Cells(brsN, klmN - 1)).NumberFormat = gs_formatQty
       
       ' Format Font
       .Range(.Cells(8, 1), .Cells(9, klmN - 1)).Columns.Font.Size = 9
       .Range(.Cells(10, 1), .Cells(brsN, klmN - 1)).Columns.Font.Size = 9
       
       .Range(.Cells(8, 1), .Cells(9, klmN - 1)).Font.Bold = True
               
       .Range(.Cells(8, 1), .Cells(9, klmN - 1)).HorizontalAlignment = xlCenter
       .Range(.Cells(8, 1), .Cells(9, klmN - 1)).NumberFormat = "MMM YYYY"
    
       For X = 1 To 16 Step 2
           .Range(.Cells(1, 4 + X), .Cells(1, 4 + X)).ColumnWidth = 5
           .Range(.Cells(1, 5 + X), .Cells(1, 5 + X)).ColumnWidth = 10
       Next
       
       .Range(.Cells(8, 17), .Cells(brsN, klmN - 1)).Columns.AutoFit
       .Range("a:d").Columns.AutoFit
       .WindowState = xlMaximized
       .Visible = True
               
    End With
    MousePointer = vbDefault
    
    Label1 = "Transfer Progress Complete ! "
    Exit Sub
    
ErExcl:
    MousePointer = vbDefault
    LblErrMsg = err.number & ":" & err.Description
    Pb01.Value = 0
    
    Exit Sub
    
End Sub

Private Sub cmdReport_Click()
On Error GoTo errHandler
    Me.MousePointer = vbHourglass
    Label1 = ""
    
'=====##Validating
'>> Empty data
    If GridSelect.Rows = 1 Then
    Me.MousePointer = vbNormal
    LblErr = DisplayMsg("8011") 'Please select data first !

    Exit Sub
    End If


'>>Provided data
'      Dim KosongCek As Boolean
'      KosongCek = False
'      For i = 1 To Grid.Rows - 1
'            If Grid.Cell(flexcpChecked, i, Col) = 2 Then
'                KosongCek = True
'
'            Else
'                KosongCek = False
'                Exit For
'            End If
'      Next i
'    If KosongCek = True Then
'    Me.MousePointer = vbNormal
'    LblErr = DisplayMsg("8011") 'Please select data first !
'    Exit Sub
'    End If

'======##End OF Validation

Call DataExel
Call SettingExel
Db.Execute "Drop Table " & StrTableTemp

Me.MousePointer = vbNormal
Exit Sub


    
errHandler:

    Me.MousePointer = vbNormal
     LblErr = "[" & err.number & "] " & err.Description
    
    
End Sub

Sub DataExel()

    Dim RssPar1 As New ADODB.Recordset
    Dim RssPAr2 As New ADODB.Recordset
    Dim RssPAr3 As New ADODB.Recordset
    Dim RssPAr4 As New ADODB.Recordset
    Dim RssPAr5 As New ADODB.Recordset
    Dim RssPar7 As New ADODB.Recordset
    
    Label1 = "Check available Data ... "

    '===###Checking BOM For Count Need of Material
    StrTableTemp = ""
    StrTableTemp = "TempITem" & CStr(Format(Now, "yyyymmdd" & "_" & "hhmmss"))
    
    'Dim NewDB As New ADODB.Connection
    'NewDB.Open Db.ConnectionString
    
    Db.BeginTrans
    
    Db.Execute "Create table " & StrTableTemp & " (Item_Code nvarChar(255),Item_Name nvarChar(255),  " & _
                    "Qty Numeric(18,5),Parent_Item nvarChar(255), Makebuy_Cls Char(2) ) "
    
        For i = 1 To GridSelect.Rows - 1
            'If GridSelect.Cell(flexcpChecked, i, BteColSelect) = 1 Then
                    If RssPar1.State <> adStateClosed Then RssPar1.Close
                    RssPar1.Open "Select (Select makebuy_Cls From Item_Master Where Item_Code = A.Item_Code) MakeBuy_Cls" & _
                                         ", Item_Code,Item_Name,Qty,Parent_ItemCode from Bom_Master A Where Parent_itemCode ='" & GridSelect.TextMatrix(i, bteColItem_Code) & "'", Db, 3, 2
                        Do While Not RssPar1.EOF
                            Insert Trim(RssPar1("Item_Code")), Trim(RssPar1("item_Name")), Trim(RssPar1("Qty")), Trim(GridSelect.TextMatrix(i, bteColItem_Code)), Trim(RssPar1("MakeBuy_Cls"))
                            'Level 2
                            RssPAr2.Open "Select (Select makebuy_Cls From Item_Master Where Item_Code = A.Item_Code) MakeBuy_Cls" & _
                                         ", Item_Code,Item_Name,Qty,Parent_ItemCode from Bom_Master A Where Parent_itemCode ='" & RssPar1("Item_Code") & "'", Db, 3, 2
                    
                    
                                        Do While Not RssPAr2.EOF
                                        Insert Trim(RssPAr2("Item_Code")), Trim(RssPAr2("item_Name")), RssPAr2("Qty") * RssPar1("Qty"), Trim(GridSelect.TextMatrix(i, bteColItem_Code)), Trim(RssPAr2("MakeBuy_Cls"))
                                        'Level 3
                                        RssPAr3.Open "Select (Select makebuy_Cls From Item_Master Where Item_Code = A.Item_Code) MakeBuy_Cls" & _
                                                     ", Item_Code,Item_Name,Qty,Parent_ItemCode from Bom_Master A Where Parent_itemCode ='" & RssPAr2("Item_Code") & "'", Db, 3, 2
                                
                                
                                
                                                Do While Not RssPAr3.EOF
                                                Insert Trim(RssPAr3("Item_Code")), Trim(RssPAr3("item_Name")), RssPAr3("Qty") * RssPAr2("Qty"), Trim(GridSelect.TextMatrix(i, bteColItem_Code)), Trim(RssPAr3("MakeBuy_Cls"))
                                                'Level 4
                                                RssPAr4.Open "Select (Select makebuy_Cls From Item_Master Where Item_Code = A.Item_Code) MakeBuy_Cls" & _
                                                             ", Item_Code,Item_Name,Qty,Parent_ItemCode from Bom_Master A Where Parent_itemCode ='" & RssPAr3("Item_Code") & "'", Db, 3, 2
                                        
                                        
                                                            Do While Not RssPAr4.EOF
                                                            Insert Trim(RssPAr4("Item_Code")), Trim(RssPAr4("item_Name")), RssPAr4("Qty") * RssPAr3("Qty"), Trim(GridSelect.TextMatrix(i, bteColItem_Code)), Trim(RssPAr4("MakeBuy_Cls"))
                                                            'Level 5
                                                            RssPAr5.Open "Select (Select makebuy_Cls From Item_Master Where Item_Code = A.Item_Code) MakeBuy_Cls" & _
                                                                         ", Item_Code,Item_Name,Qty,Parent_ItemCode from Bom_Master A Where Parent_itemCode ='" & RssPAr4("Item_Code") & "'", Db, 3, 2
                                                    
                                                    
                                                    
                                                                        Do While Not RssPAr5.EOF
                                                                        Insert Trim(RssPAr5("Item_Code")), Trim(RssPAr5("item_Name")), RssPAr5("Qty") * RssPAr4("Qty"), Trim(GridSelect.TextMatrix(i, bteColItem_Code)), Trim(RssPAr5("MakeBuy_Cls"))
                                                                        'Level 6
                                                                        RssPar6.Open "Select (Select makebuy_Cls From Item_Master Where Item_Code = A.Item_Code) MakeBuy_Cls" & _
                                                                                     ", Item_Code,Item_Name,Qty,Parent_ItemCode from Bom_Master A Where Parent_itemCode ='" & RssPAr5("Item_Code") & "'", Db, 3, 2
                                                                
                                                                
                                                                
                                                                                    Do While Not RssPar6.EOF
                                                                                    Insert Trim(RssPar6("Item_Code")), Trim(RssPar6("item_Name")), RssPar6("Qty") * RssPAr5("Qty"), Trim(GridSelect.TextMatrix(i, bteColItem_Code)), Trim(RssPar6("MakeBuy_Cls"))
                                                                                    'Level 7
                                                                                    RssPar7.Open "Select (Select makebuy_Cls From Item_Master Where Item_Code = A.Item_Code) MakeBuy_Cls" & _
                                                                                                 ", Item_Code,Item_Name,Qty,Parent_ItemCode from Bom_Master A Where Parent_itemCode ='" & RssPar6("Item_Code") & "'", Db, 3, 2
                                                                            
                                                                            
                                                                            
                                                                                             Do While Not RssPar7.EOF
                                                                                             Insert Trim(RssPar7("Item_Code")), Trim(RssPar7("item_Name")), RssPar7("Qty") * RssPar6("Qty"), Trim(GridSelect.TextMatrix(i, bteColItem_Code)), Trim(RssPar7("MakeBuy_Cls"))
                                                                                            
                                                                                             RssPar7.MoveNext
                                                                                             Loop
                                                                                                        
                                                                            
                                                                            
                                                                            
                                                                                     '<<Level 7
                                                                                    RssPar7.Close
                                                                                    RssPar6.MoveNext
                                                                                    Loop
                                                                                            
                                                                
                                                                
                                                                         '<<Level 6
                                                                        RssPar6.Close
                                                                        RssPAr5.MoveNext
                                                                        Loop
                                                                                                        
                                                    
                                                    
                                                             '<<Level 5
                                                            RssPAr5.Close
                                                            RssPAr4.MoveNext
                                                            Loop
                                                                                
                                        
                                        
                                        
                                                 '<<Level 4
                                                RssPAr4.Close
                                                RssPAr3.MoveNext
                                                Loop
                                
                                
                                
                                    '<< Level 3
                                    RssPAr3.Close
                                    RssPAr2.MoveNext
                                    Loop
                    
                            '<< Level 2
                            RssPAr2.Close
                        RssPar1.MoveNext
                        Loop
                    RssPar1.Close
            'End If
        Next i
        
        
        Db.CommitTrans

    
End Sub


Sub Insert(Item_Code As String, PartName As String, Qty As Double, Parent_Code As String, MakeBuy_Cls As String)
    Db.Execute "Insert Into " & StrTableTemp & " values( '" & Item_Code & "' , '" & PartName & "', " & Qty & ", '" & Parent_Code & "', '" & MakeBuy_Cls & "')"
End Sub
Private Sub cmdSearch_Click()

    Dim RSSCari As New ADODB.Recordset
    Dim StrOPen As String
    Dim kondisi As String
    
    If Trim(cboSupplier) = "" Then LblErr = DisplayMsg("1054"): Exit Sub 'Please Selecet Supplier Code
    
    If Trim(CboGroupCls) = "" Then LblErr = DisplayMsg("8081"): Exit Sub 'Please select Group Cls !
    
    LblErr = ""
    kondisi = ""
    
    Call Header
    
    'If Trim(CboSupplier) = "ALL" And Trim(CboGroupCls) = "ALL" Then
    '    kondisi = " Where FinishgoodPart_Cls='01'"
    '
    'ElseIf Trim(CboSupplier) = "ALL" And Trim(CboGroupCls) <> "ALL" Then
    '    kondisi = "Where group_Cls = '" & Trim(CboGroupCls) & "' And FinishgoodPart_Cls='01'"
    '
    'ElseIf Trim(CboSupplier) <> "ALL" And Trim(CboGroupCls) = "ALL" Then
    '    kondisi = " Where Supplier_Code ='" & Trim(CboSupplier) & "' and FinishgoodPart_Cls='01'"
    '
    'Else
    '    kondisi = " Where Supplier_Code ='" & Trim(CboSupplier) & "' and group_Cls = '" & Trim(CboGroupCls) & "' and FinishgoodPart_Cls='01'" 'real Property
    'End If
    
    If Trim(CboGroupCls) = strAll Then
        kondisi = "Where FinishgoodPart_Cls='01'"
    Else
        kondisi = "Where group_Cls = '" & Trim(CboGroupCls) & "' And FinishgoodPart_Cls='01'"
    End If
    
    StrOPen = ""
    StrOPen = "Select*from Item_Master " & kondisi & " Order By Item_Name"
    
    If RSSCari.State <> adStateClosed Then RSSCari.Close
    
    RSSCari.Open StrOPen, Db, adOpenDynamic, adLockOptimistic
    Do While Not RSSCari.EOF
        With grid
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, bteColItem_Code) = IIf(IsNull(RSSCari("Item_Code")), "", Trim(RSSCari("Item_Code")))
        
            .TextMatrix(.Rows - 1, bteColDescription) = IIf(IsNull(RSSCari("Item_name")), "", Trim(RSSCari("Item_name")))
            
           .ColAlignment(bteColSelect) = flexAlignCenterCenter
           .Cell(flexcpChecked, .Rows - 1, bteColSelect) = flexUnchecked
            
        End With
        RSSCari.MoveNext
    Loop
    If grid.Rows > 1 Then
        grid.Cell(flexcpAlignment, 1, bteColItem_Code, grid.Rows - 1, bteColDescription) = flexAlignLeftCenter
       ' Grid.Cell(flexcpBackColor, 1, bteColDescription, Grid.Rows - 1, bteColDescription) = &HFFFFFF
        grid.Cell(flexcpBackColor, 1, bteColSelect, grid.Rows - 1, bteColSelect) = &HFFFFFF
    
    End If
    RSSCari.Close

End Sub

Private Sub CmdSubMenu_Click()
    DoEvents
    frmMainMenu.Show
    DoEvents
    Unload Me
End Sub
Sub Header()
With grid
    .clear

    .Rows = 1
    .ColS = 4
    
    bteColSelect = 0
    bteColItem_Code = 1
    bteColDescription = 2
    bteColGroup = 3
    
    .ColWidth(bteColSelect) = 300
    .ColWidth(bteColItem_Code) = 2000
    .ColWidth(bteColDescription) = 4250
    .ColWidth(bteColGroup) = 500
    
    .TextMatrix(0, bteColSelect) = ""
    .TextMatrix(0, bteColItem_Code) = "Part No"
    .TextMatrix(0, bteColDescription) = "Part Name"
    
    .ColHidden(bteColGroup) = True
    
    .Cell(flexcpAlignment, 0, 0, 0, .ColS - 1) = flexAlignCenterCenter
    .RowHeight(0) = 250
    .Cell(flexcpChecked, 0, bteColSelect) = flexUnchecked
    .Sort = flexSortCustom
  End With
  
  Call HeaderGridSelect
End Sub

Sub HeaderGridSelect()
With GridSelect
    .clear

    .Rows = 1
    .ColS = 3
    
    bteColSelect = 0
    bteColItem_Code = 1
    bteColDescription = 2
    bteColGroup = 3
    
    .ColWidth(bteColSelect) = 300
    .ColWidth(bteColItem_Code) = 2000
    .ColWidth(bteColDescription) = 4250
    
    .TextMatrix(0, bteColSelect) = ""
    .TextMatrix(0, bteColItem_Code) = "Part No"
    .TextMatrix(0, bteColDescription) = "Part Name"
    
    .Cell(flexcpAlignment, 0, 0, 0, .ColS - 1) = flexAlignCenterCenter
    .RowHeight(0) = 250
    .Cell(flexcpChecked, 0, bteColSelect) = flexUnchecked
    .ColAlignment(bteColItem_Code) = flexAlignLeftCenter
    .Sort = flexSortCustom
    .ColHidden(bteColSelect) = True
    
  End With

End Sub

Private Sub ClearData()
Dim X As Integer
Dim RSSCari As New ADODB.Recordset
Dim StrOPen As String
Dim kondisi As String

StrOPen = ""
StrOPen = "Select * From Item_Master Where FinishGoodPart_Cls='01' Order By Item_Name"

If RSSCari.State <> adStateClosed Then RSSCari.Close

RSSCari.Open StrOPen, Db, adOpenDynamic, adLockOptimistic
Do While Not RSSCari.EOF
    With grid
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, bteColItem_Code) = IIf(IsNull(RSSCari("Item_Code")), "", Trim(RSSCari("Item_Code")))
        .TextMatrix(.Rows - 1, bteColDescription) = IIf(IsNull(RSSCari("Item_name")), "", Trim(RSSCari("Item_name")))
        .TextMatrix(.Rows - 1, bteColGroup) = IIf(IsNull(RSSCari("Group_Cls")), "", Trim(RSSCari("Group_Cls")))

       .ColAlignment(bteColSelect) = flexAlignCenterCenter
       .Cell(flexcpChecked, .Rows - 1, bteColSelect) = flexUnchecked
        
    End With
    RSSCari.MoveNext
Loop
If grid.Rows > 1 Then
    grid.Cell(flexcpAlignment, 1, bteColItem_Code, grid.Rows - 1, bteColDescription) = flexAlignLeftCenter
    grid.Cell(flexcpBackColor, 1, bteColSelect, grid.Rows - 1, bteColSelect) = &HFFFFFF
    grid.Row = grid.Rows - 1
    grid.Col = 0
End If
RSSCari.Close

Label1 = ""
Pb01.Value = 0
End Sub

Private Sub Form_Load()
    Label1 = ""
    Call adtocboGroup
    Call adtocbosupplier
    Call Header
    Call ClearData
    Tgl1 = Format(Now(), "MMM YYYY")
    Tgl2 = Format(DateAdd("M", 5, Now()), "MMM YYYY")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With grid
        If .Col = bteColSelect Then
            If Row = 0 Then
                If .Cell(flexcpChecked, Row, Col) = 1 Then
                    For i = 1 To .Rows - 1
                        If CboGroupCls = strAll Then
                            .Cell(flexcpChecked, i, Col) = 1 ' untuk seluruh item
                        Else
                            If .TextMatrix(i, bteColGroup) = CboGroupCls Then
                                 .Cell(flexcpChecked, i, Col) = 1
                            End If
                        End If
                    Next i
                'flexChecked
                ElseIf .Cell(flexcpChecked, Row, Col) = 2 Then
                    For i = 1 To .Rows - 1
                        If CboGroupCls = strAll Then
                                .Cell(flexcpChecked, i, Col) = 2 'untuk seluruh item
                        Else
                            If .TextMatrix(i, bteColGroup) = CboGroupCls Then
                                .Cell(flexcpChecked, i, Col) = 2
                            End If
                        End If
                    Next i
                'flexUnchecked
                End If
            End If
            
            Call CekMark
        
            Call SelectGrid
        End If
End With
    
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If grid.Col <> bteColSelect Then
        Cancel = True
    Else
    
    End If
    
End Sub

Private Sub grid_Click()
    LblErr = ""
End Sub

Private Sub Tgl1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
    tgl1_Click
End Sub

Private Sub tgl1_Change()
    tgl1_Click
    tgl_sb = Tgl1.Month
End Sub

Private Sub tgl1_Click()
    If Tgl1.Month = 1 And Val(tgl_sb) = 12 Then Tgl1.Year = Tgl1.Year + 1
    If Tgl1.Month = 12 And Val(tgl_sb) = 1 Then Tgl1.Year = Tgl1.Year - 1
    Tgl2 = Format(DateAdd("m", 5, Format(Tgl1, "yyyy-mm-dd")), "MMM YYYY")
End Sub
