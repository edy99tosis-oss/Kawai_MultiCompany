VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmValuationPriceReportDetail 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Valuation Price Report Detail"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15105
   Icon            =   "FrmValuationPriceReportDetail.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   15105
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   3120
      MultiLine       =   -1  'True
      TabIndex        =   29
      Tag             =   "TTFF*/"
      Text            =   "FrmValuationPriceReportDetail.frx":0E42
      Top             =   9960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox Anchor1 
      Height          =   480
      Left            =   1920
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   30
      Top             =   9870
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton btnPriview 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Avg Price Detail"
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
      Left            =   11850
      Style           =   1  'Graphical
      TabIndex        =   27
      Tag             =   "FFTT*/"
      Top             =   10020
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.CommandButton cmdExcel 
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
      Left            =   13730
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "FFTT*/"
      Top             =   10020
      Width           =   1140
   End
   Begin VB.CommandButton cmdReport 
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
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "TTFF*/"
      Top             =   2520
      Width           =   1185
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   600
      Left            =   360
      TabIndex        =   12
      Tag             =   "TFTT*/"
      Top             =   9150
      Width           =   14535
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
         Left            =   135
         TabIndex        =   13
         Tag             =   "TFTF*/"
         Top             =   210
         Width           =   14205
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   2265
      Left            =   360
      TabIndex        =   7
      Tag             =   "TTTF*/"
      Top             =   960
      Width           =   14475
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   300
         Left            =   4320
         TabIndex        =   23
         Tag             =   "TTFF*/"
         Top             =   1095
         Width           =   300
      End
      Begin MSComCtl2.DTPicker DtPeriod 
         Height          =   330
         Left            =   2160
         TabIndex        =   2
         Tag             =   "TTFF*/"
         Top             =   1560
         Width           =   1275
         _ExtentX        =   2249
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
         Format          =   141230083
         UpDown          =   -1  'True
         CurrentDate     =   37860
      End
      Begin MSForms.TextBox TxtProduk 
         Height          =   315
         Left            =   2160
         TabIndex        =   1
         Tag             =   "TTFF*/"
         Top             =   1080
         Width           =   2055
         VariousPropertyBits=   746604571
         MaxLength       =   15
         Size            =   "3625;556"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblPartNumber 
         BackStyle       =   0  'Transparent
         Caption         =   "Part Number"
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
         Left            =   4740
         TabIndex        =   28
         Tag             =   "TTFF*/"
         Top             =   1110
         Width           =   3060
      End
      Begin VB.Line Line8 
         Index           =   2
         X1              =   4740
         X2              =   7860
         Y1              =   1380
         Y2              =   1380
      End
      Begin MSForms.CheckBox CheckBox1 
         Height          =   345
         Left            =   3630
         TabIndex        =   26
         Tag             =   "TTFF*/"
         Top             =   1560
         Width           =   1365
         BackColor       =   16637923
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "2408;609"
         Value           =   "0"
         Caption         =   "Include"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin VB.Label Lbl_Make 
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
         Height          =   225
         Left            =   10200
         TabIndex        =   25
         Tag             =   "TTFF*/"
         Top             =   690
         Width           =   1845
      End
      Begin VB.Line Line8 
         Index           =   1
         X1              =   10170
         X2              =   12930
         Y1              =   945
         Y2              =   945
      End
      Begin VB.Label lbl_cls 
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
         Height          =   225
         Left            =   4080
         TabIndex        =   24
         Tag             =   "TTFF*/"
         Top             =   690
         Width           =   2685
      End
      Begin VB.Line Line6 
         X1              =   4020
         X2              =   6750
         Y1              =   945
         Y2              =   945
      End
      Begin MSForms.ComboBox CboMb 
         Height          =   315
         Left            =   8310
         TabIndex        =   22
         Tag             =   "TTFF*/"
         Top             =   660
         Width           =   1755
         VariousPropertyBits=   746604569
         BackColor       =   16777215
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "3096;556"
         ListRows        =   15
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox CboCls 
         Height          =   315
         Left            =   2160
         TabIndex        =   21
         Tag             =   "TTFF*/"
         Top             =   660
         Width           =   1755
         VariousPropertyBits=   746604569
         BackColor       =   16777215
         MaxLength       =   2
         DisplayStyle    =   3
         Size            =   "3096;556"
         ListRows        =   15
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Make Buy Cls"
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
         Left            =   6990
         TabIndex        =   15
         Tag             =   "TTFF*/"
         Top             =   705
         Width           =   1275
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Finish Good Part Cls"
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
         Left            =   270
         TabIndex        =   11
         Tag             =   "TTFF*/"
         Top             =   690
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Month Period"
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
         Index           =   1
         Left            =   270
         TabIndex        =   10
         Tag             =   "TTFF*/"
         Top             =   1650
         Width           =   1350
      End
      Begin VB.Line Line8 
         Index           =   0
         X1              =   8010
         X2              =   12480
         Y1              =   1380
         Y2              =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Code"
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
         Left            =   270
         TabIndex        =   9
         Tag             =   "TTFF*/"
         Top             =   1170
         Width           =   1155
      End
      Begin VB.Label lblNm 
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
         Height          =   195
         Left            =   8010
         TabIndex        =   8
         Tag             =   "TTFF*/"
         Top             =   1110
         Width           =   4410
      End
      Begin MSForms.ComboBox CboProduk 
         Height          =   375
         Left            =   2220
         TabIndex        =   0
         Tag             =   "TTFF*/"
         Top             =   2580
         Visible         =   0   'False
         Width           =   2085
         VariousPropertyBits=   746604571
         BackColor       =   16777215
         MaxLength       =   15
         DisplayStyle    =   3
         Size            =   "3678;661"
         ListRows        =   15
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "TFFT*/"
      Top             =   10020
      Width           =   1140
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   5160
      Left            =   360
      TabIndex        =   6
      Tag             =   "TTTT*/"
      Top             =   3600
      Width           =   14475
      _cx             =   25532
      _cy             =   9102
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
      Rows            =   10
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
      Editable        =   2
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
      Left            =   12990
      TabIndex        =   16
      Tag             =   "FTTF*/"
      Top             =   315
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   714
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Record(s)"
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
      Left            =   13830
      TabIndex        =   20
      Tag             =   "FFTT*/"
      Top             =   8850
      Width           =   1065
   End
   Begin VB.Label LblRecord 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Record(s)"
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
      Left            =   12810
      TabIndex        =   19
      Tag             =   "FFTT*/"
      Top             =   8850
      Width           =   1065
   End
   Begin VB.Label LblBase 
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
      Left            =   2055
      TabIndex        =   18
      Tag             =   "TTFF*/"
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Base Currency : "
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
      Index           =   3
      Left            =   360
      TabIndex        =   17
      Tag             =   "TTFF*/"
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valuation Price Report Detail"
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
      Left            =   5970
      TabIndex        =   14
      Tag             =   "TTTF*/"
      Top             =   360
      Width           =   3270
   End
End
Attribute VB_Name = "FrmValuationPriceReportDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dateUp As Date

Dim ColDate As Byte, ColCls As Byte, ColDesc As Byte, ColFromLoc As Byte, ColToLoc As Byte, ColPONo As Byte, ColDoNo As Byte, ColLotNo As Byte
Dim ColPremStock As Byte, ColIncomeStock As Byte, ColIncomeOtherStock As Byte, ColOutgoingStock As Byte, ColOutgoingOtherStock As Byte
Dim ColCurrent As Byte, ColCurr As Byte, ColPrice As Byte, ColRate As Byte, ColAvgPrice As Byte, ColAmount As Byte, ColValueInv As Byte

Dim sqlfg As String, sqlmb As String
Dim SqlData As String
Public setTgl, setItem, setTglPrint, setInclude As String

Private Sub btnPriview_Click()
    Dim rsPrive As New ADODB.Recordset
    Dim application As New CRAXDDRT.application
    Dim report As New CRAXDDRT.report
    Dim Rpt As New FrmRpt3
    Dim temp1, temp2, temp3, temp4 As String
    Dim temp5, temp6 As String
    Dim rsCls As New ADODB.Recordset
    Dim sqlProcess As String
    
'    If CboCls = "01" Then
'        temp1 = "0"
'    Else
'        temp1 = ""
'    End If

    If CheckBox1.Value = True Then
        temp3 = "1"
    Else
        temp3 = "0"
    End If
    
    If Left(TxtProduk, 1) = "L" Then
        temp1 = "1"
        temp2 = "3"
    ElseIf Left(TxtProduk, 1) = "E" Then
        temp1 = "0"
        temp2 = "3"
    Else
        temp1 = temp3
        temp2 = "3"
    End If
    
    sql = "select rate_cls from company_profile"
    Set rsCls = Db.Execute(sql)
    
    If Not rsCls.EOF Then temp6 = rsCls!Rate_Cls
 

    If grid.TextMatrix(grid.Row, ColCls) = "P1" Then
     sqlProcess = "   select 1 idx, " & vbCrLf & _
                  "       null part_code, null duty_status," & vbCrLf & _
                  "       description part_name, " & vbCrLf

     sqlProcess = sqlProcess & vbCrLf & _
                  "       isnull(standard_time,0) qtyOut, " & vbCrLf & _
                  "       'minute' unit_desc, " & vbCrLf & _
                  "       null avgprice, " & vbCrLf & _
                  "       null supplier_code,  " & vbCrLf
            
     sqlProcess = sqlProcess & vbCrLf & _
                  "        ( isnull(cost_minute,0)  " & vbCrLf & _
                  "        *  " & vbCrLf & _
                  "        case " & temp6 & " when '0' then  " & vbCrLf & _
                  "           round((isnull((dbo.UF_GetDailyExchangeRate('" & Format(grid.TextMatrix(grid.Row, ColDate), "yyyy-mm-dd") & "',pm.Currency_Code)),0)), " & gi_decimalDigitExchangeRate & ") " & vbCrLf & _
                  "        else  " & vbCrLf & _
                  "           round((isnull((dbo.UF_GetBookExchangeRate (Year(" & Format(grid.TextMatrix(grid.Row, ColDate), "yyyy") & "),Month(" & Format(grid.TextMatrix(grid.Row, ColDate), "mm") & "), pm.Currency_Code)),0)), " & gi_decimalDigitExchangeRate & ") " & vbCrLf & _
                  "        end )process_cost, " & vbCrLf

      sqlProcess = sqlProcess & vbCrLf & _
                   "       null duty,  " & vbCrLf

      sqlProcess = sqlProcess & vbCrLf & _
                   "        isnull(standard_time,0) * isnull(cost_minute,0) * " & vbCrLf & _
                   "       (case " & temp6 & " when '0' then  " & vbCrLf & _
                   "           round((isnull((dbo.UF_GetDailyExchangeRate('" & Format(grid.TextMatrix(grid.Row, ColDate), "yyyy-mm-dd") & "',pm.Currency_Code)),0)), " & gi_decimalDigitExchangeRate & ") " & vbCrLf & _
                   "        else  " & vbCrLf & _
                   "           round((isnull((dbo.UF_GetBookExchangeRate (Year(" & Format(grid.TextMatrix(grid.Row, ColDate), "yyyy") & "),Month(" & Format(grid.TextMatrix(grid.Row, ColDate), "mm") & "), pm.Currency_Code)),0)), " & gi_decimalDigitExchangeRate & ") " & vbCrLf & _
                   "        end ) totalAvg  " & vbCrLf & _
                   "   from process_master pm  " & vbCrLf & _
                   "   left join process_cls pc on pc.process_cls=pm.process_cls  " & vbCrLf & _
                   "   where item_code='" & Trim(TxtProduk) & "'  " & vbCrLf

    ElseIf (grid.TextMatrix(grid.Row, ColCls) = "R" And grid.TextMatrix(grid.Row, ColLotNo) = "Subcon") Then
      sqlProcess = "   select 1 idx, " & vbCrLf & _
                  "       null part_code, null duty_status," & vbCrLf & _
                  "       'Service Cost' part_name, " & vbCrLf

      sqlProcess = sqlProcess & vbCrLf & _
                  "       null qtyOut, " & vbCrLf & _
                  "       null unit_desc, " & vbCrLf & _
                  "       null avgprice, " & vbCrLf & _
                  "       null supplier_code,  " & vbCrLf
            
      sqlProcess = sqlProcess & vbCrLf & _
                  "process_cost = " & vbCrLf & _
                  "--service price " & vbCrLf & _
                  "round( " & vbCrLf & _
                  "((case pr.currency_code when '03' then round((case pr.receipt_cls when 'R' then isnull(pr.price,0) else 0 end), @DigitPriceIDR) else round((case pr.receipt_cls when 'R' then isnull(pr.price,0) else 0 end),@DigitPrice) end) " & vbCrLf & _
                  "* " & vbCrLf & _
                  "--rate  " & vbCrLf & _
                  "round((case 0 when '0' then dbo.UF_GetDailyExchangeRate(pr.Receipt_Date,pr.Currency_Code) else dbo.UF_GetBookExchangeRate (Year(pr.Receipt_Date),Month(pr.Receipt_Date),pr.Currency_Code) end),@DigitExchRate)) " & vbCrLf & _
                  ", @DigitPriceIDR), " & vbCrLf & _
                  "" & vbCrLf & _
                  "null duty, " & vbCrLf & _
                  "totalAvg  = " & vbCrLf & _
                  "--service price " & vbCrLf & _
                  "round( " & vbCrLf & _
                  "((case pr.currency_code when '03' then round((case pr.receipt_cls when 'R' then isnull(pr.price,0) else 0 end), @DigitPriceIDR) else round((case pr.receipt_cls when 'R' then isnull(pr.price,0) else 0 end),@DigitPrice) end) " & vbCrLf & _
                  "* " & vbCrLf & _
                  "--rate " & vbCrLf & _
                  "round((case " & temp6 & " when '0' then dbo.UF_GetDailyExchangeRate(pr.Receipt_Date,pr.Currency_Code) else dbo.UF_GetBookExchangeRate (Year(pr.Receipt_Date),Month(pr.Receipt_Date),pr.Currency_Code) end),@DigitExchRate)) " & vbCrLf & _
                  ", @DigitPriceIDR) " & vbCrLf & _
                  "from part_receipt pr " & vbCrLf & _
                  "where pr.item_code='" & Trim(TxtProduk) & "'  " & vbCrLf & _
                  "and pr.seq_no = '" & grid.TextMatrix(grid.Row, ColDoNo) & "' " & vbCrLf
    Else
     Exit Sub
    End If
    
    
 
    
    Me.MousePointer = vbHourglass
    
    sql = "declare @DigitPriceIDR tinyint " & vbCrLf & _
          "declare @DigitPrice tinyint " & vbCrLf & _
          "declare @DigitExchRate tinyint " & vbCrLf & _
          "declare @CompanyCode char(5)" & vbCrLf & _
          "set @DigitPriceIDR = " & gi_decimalDigitPriceIDR & vbCrLf & _
          "set @DigitPrice = " & gi_decimalDigitPrice & vbCrLf & _
          "set @DigitExchRate = " & gi_decimalDigitExchangeRate & vbCrLf & _
          "" & vbCrLf

    sql = sql + " select * " & vbCrLf & _
            " from " & vbCrLf & _
            " ( " & vbCrLf & _
            "   select 1 idx,item_code,item_name,makeritem_code model, gc.description category, rtrim(uc.description) unit " & vbCrLf & _
            "   from item_master im  " & vbCrLf & _
            "   left join group_cls gc on gc.group_cls=im.group_cls " & vbCrLf & _
            "   left join unit_cls uc on uc.unit_cls=im.unit_cls " & vbCrLf & _
            "   where item_code='" & Trim(TxtProduk) & "' " & vbCrLf & _
            " ) a " & vbCrLf & _
            " full outer join " & vbCrLf & _
            " ( " & vbCrLf & _
            "   select 1 idx, " & vbCrLf

     sql = sql + "       ivp.item_code part_code, duty_status, " & vbCrLf & _
            "       item_name part_name, " & vbCrLf & _
            "       (qtyOut/" & CDbl(grid.TextMatrix(grid.Row, ColIncomeStock)) & ")qtyOut, " & vbCrLf & _
            "       uc.description unit_desc, " & vbCrLf & _
            "       avgprice, " & vbCrLf & _
            "       supplier_code,  " & vbCrLf & _
            "       null Process_cost, " & vbCrLf & _
            "       case " & temp1 & " when 0 then null  " & vbCrLf & _
            "              when 1 then (case duty_status when 3 then isnull(hm.tax,0) else null end) " & vbCrLf & _
            "              end tax,  " & vbCrLf
            
     sql = sql + "       case " & temp1 & " when 0 then ((qtyOut/" & CDbl(grid.TextMatrix(grid.Row, ColIncomeStock)) & ")* avgprice)" & vbCrLf & _
            "              when 1 then (case duty_status when 3 then ((qtyOut/" & CDbl(grid.TextMatrix(grid.Row, ColIncomeStock)) & ")* avgprice)*(1 + (isnull(hm.tax,0)/100)) " & vbCrLf & _
            "                                            else ((qtyOut/" & CDbl(grid.TextMatrix(grid.Row, ColIncomeStock)) & ")* avgprice) end) " & vbCrLf & _
            "             end totalAvg " & vbCrLf
            
     sql = sql + "   from Inventory_PriceDetail ivp  " & vbCrLf & _
            "   left join item_master im on im.item_code=ivp.item_code " & vbCrLf & _
            "   left join unit_cls uc on uc.unit_cls=im.unit_cls  " & vbCrLf & _
            "   left join hs_master hm on hm.hs_code=im.hs_code  " & vbCrLf & _
            "   where do_no='" & grid.TextMatrix(grid.Row, ColDoNo) & "' and duty_status in ('" & temp1 & "','" & temp2 & "')  " & vbCrLf & _
            "   and cls = 'S' " & vbCrLf & _
            "    " & vbCrLf & _
            "   union all " & vbCrLf & _
            "    " & vbCrLf
            
' ------------------------------------
' 20130730 - Add History
' ------------------------------------

    sql = sql & vbCrLf & _
            "   select 1 idx, " & vbCrLf

     sql = sql + "       ivp.item_code part_code, duty_status, " & vbCrLf & _
            "       item_name part_name, " & vbCrLf & _
            "       (qtyOut/" & CDbl(grid.TextMatrix(grid.Row, ColIncomeStock)) & ")qtyOut, " & vbCrLf & _
            "       uc.description unit_desc, " & vbCrLf & _
            "       avgprice, " & vbCrLf & _
            "       supplier_code,  " & vbCrLf & _
            "       null Process_cost, " & vbCrLf & _
            "       case " & temp1 & " when 0 then null  " & vbCrLf & _
            "              when 1 then (case duty_status when 3 then isnull(hm.tax,0) else null end) " & vbCrLf & _
            "              end tax,  " & vbCrLf
            
     sql = sql + "       case " & temp1 & " when 0 then ((qtyOut/" & CDbl(grid.TextMatrix(grid.Row, ColIncomeStock)) & ")* avgprice)" & vbCrLf & _
            "              when 1 then (case duty_status when 3 then ((qtyOut/" & CDbl(grid.TextMatrix(grid.Row, ColIncomeStock)) & ")* avgprice)*(1 + (isnull(hm.tax,0)/100)) " & vbCrLf & _
            "                                            else ((qtyOut/" & CDbl(grid.TextMatrix(grid.Row, ColIncomeStock)) & ")* avgprice) end) " & vbCrLf & _
            "             end totalAvg " & vbCrLf
            
     sql = sql + "   from InventoryPriceDetail_History ivp  " & vbCrLf & _
            "   left join item_master im on im.item_code=ivp.item_code " & vbCrLf & _
            "   left join unit_cls uc on uc.unit_cls=im.unit_cls  " & vbCrLf & _
            "   left join hs_master hm on hm.hs_code=im.hs_code  " & vbCrLf & _
            "   where do_no='" & grid.TextMatrix(grid.Row, ColDoNo) & "' and duty_status in ('" & temp1 & "','" & temp2 & "')  " & vbCrLf & _
            "   and cls = 'S' " & vbCrLf & _
            "    " & vbCrLf & _
            "   union all " & vbCrLf & _
            "    " & vbCrLf
' -----------------------------------

     sql = sql + sqlProcess
            
     sql = sql & vbCrLf & _
            " ) b " & vbCrLf & _
            " on a.idx=b.idx "
                
                
    Set rsPrive = Db.Execute(sql)

    sqlprint = sql
    reportcode = "MaterialConsumptionDetail"
    printorient = 2
    
    Set report = application.OpenReport(App.path & "\Reports\rptMCD.rpt")
    report.Database.Tables(1).SetDataSource rsPrive
    
    If cboCls = "01" Then
        If Left(Trim(TxtProduk), 1) = "L" Then
            temp5 = "Include Duty"
        ElseIf Left(Trim(TxtProduk), 1) = "E" Then
            temp5 = "Exclude Duty"
        End If
    ElseIf cboCls = "02" Then
        If CheckBox1.Value = True Then
            temp5 = "Include Duty"
        Else
            temp5 = "Exclude Duty"
        End If
    End If
    
    report.PaperOrientation = crLandscape
    report.FormulaFields.GetItemByName("01_Date").Text = "'" & Format(grid.TextMatrix(grid.Row, ColDate), "dd mmm yyyy") & "'"
    report.FormulaFields.GetItemByName("02_ItemName").Text = "'" & Trim(TxtProduk) & " / " & lblPartNumber & "'"
    report.FormulaFields.GetItemByName("03_Include").Text = "'" & temp5 & "'"
    report.FormulaFields.GetItemByName("05_TanggalPrint").Text = "'" & Format(Now, "dddd, dd mmmm, yyyy  hh:mm:ss") & "'"
    report.FormulaFields.GetItemByName("80_DigitAmountIDR").Text = gi_decimalDigitAmountIDR
    report.FormulaFields.GetItemByName("81_DigitQtyBOM").Text = gi_decimalDigitQtyBOM
    report.FormulaFields.GetItemByName("82_DigitPriceIDR").Text = gi_decimalDigitPriceIDR
    report.FormulaFields.GetItemByName("83_DigitPercentage").Text = 2
    
    
    setInclude = temp5
    setTgl = Format(grid.TextMatrix(grid.Row, ColDate), "dd mmm yyyy")
    setItem = Trim(TxtProduk) & " / " & lblPartNumber
    setTglPrint = Format(Now, "dddd, dd mmmm, yyyy  hh:mm:ss")
    
    Rpt.CRViewer1.ReportSource = report
    Rpt.CRViewer1.ViewReport
    Rpt.CRViewer1.Zoom 1
    MousePointer = vbDefault
    Rpt.WindowState = 2
    
    Rpt.Show 1
    
    Set rsPrive = Nothing
    Me.MousePointer = vbDefault
    LblErrMsg = ""
End Sub

Private Sub CboCompany_Change()
CboCompany_Click
End Sub
Private Sub CboCompany_Click()
    
End Sub
Private Sub CheckBox1_Click()
    grid.clear
    Header
End Sub

Private Sub Command1_Click()
    Me.MousePointer = vbHourglass
    frm_BrowseItem.getItemCode = TxtProduk.Text
    frm_BrowseItem.Show 1
    
'    CboProduk.Text = frm_BrowseItem.getItemCode
'    CboProduk.SetFocus

    '20130729 - Change to TextBox
    ' ---------------------------------------------------
    TxtProduk.Text = frm_BrowseItem.getItemCode
    TxtProduk.SetFocus
    '---------------------------------------------------
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    With Anchor1
    'Simply provide the RegString to save theposition of form
    'and use it for next show of the forml
    'Seprate the AppName,Section,Key (Same as VB GetSetting) with Comma
    'And Default value (Top and Left) with | Like
below:
    '.RegString ="AnchorCtrl,Positions,FrmAnchDemo,1110|540"
    '.RegString = "AnchorCtrl,Positions," & Me.Name & "0|0"
    '.DoInit ' Has 2 optional arg, that can sets
    'Height and Width of the form when showing the Form
    End With
    
  'Call GetCompanySetup(Me.Name, CboCompany)
  
  If gb_Simulation = True Then Call up_InitSimulation(Me)
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
    DtPeriod.Value = Format(Now, "MMM yyyy")
    dateUp = DtPeriod.Value
    konfigA
    
    Dim RS As New ADODB.Recordset
    RS.Open "select valuationprice_BaseCurrency from Company_Profile", Db, adOpenForwardOnly, adLockReadOnly
    lblBase.Caption = uf_GetCurrencyDescription(Trim$(RS(0) & ""))
    RS.Close
    
    Call SettingColumn
    Call Header
End Sub

Private Sub SettingColumn()
ColDate = 0
ColCls = 1
ColDesc = 2
ColFromLoc = 3
ColToLoc = 4
ColPONo = 5
ColDoNo = 6
ColLotNo = 7
ColPremStock = 8
ColIncomeStock = 9
ColIncomeOtherStock = 10
ColOutgoingStock = 11
ColOutgoingOtherStock = 12
ColCurrent = 13
ColCurr = 14
ColPrice = 15
ColRate = 16
ColAvgPrice = 17
ColAmount = 18
ColValueInv = 19
End Sub

Private Sub konfigA()


cboCls.clear
cboCls.columnCount = 2

'cbocls.AddItem
'cbocls.List(0, 0) = strAll
'cbocls.List(0, 1) = strAll
cboCls.AddItem
cboCls.List(0, 0) = "01"
cboCls.List(0, 1) = "Finish Goods"
cboCls.AddItem
cboCls.List(1, 0) = "02"
cboCls.List(1, 1) = "Parts/WIP/Material"

cboCls.ListWidth = 110
cboCls.ColumnWidths = "20 pt ; 90 pt "
cboCls.ListIndex = 0
cboCls.Text = cboCls.List(0, 0)
'==========================
CboMb.clear
CboMb.columnCount = 2

CboMb.AddItem
CboMb.List(0, 0) = strAll
CboMb.List(0, 1) = strAll
CboMb.AddItem
CboMb.List(1, 0) = "01"
CboMb.List(1, 1) = "Make"
CboMb.AddItem
CboMb.List(2, 0) = "02"
CboMb.List(2, 1) = "Buy"

CboMb.ListWidth = 70
CboMb.ColumnWidths = "20 pt ; 50 pt "
CboMb.ListIndex = 0
CboMb.Text = CboMb.List(0, 0)

    cboCls.ListIndex = 0
    CboMb.ListIndex = 0
    
    DoEvents
    'CariProduk
    Header
    
    TxtProduk = ""
    lblPartNumber.Caption = ""
    lblNm = ""
    
End Sub

Private Sub cbocls_Change()
    If cboCls.Text = "" Then
      LblErrMsg.Caption = ""
      CboProduk.clear
    End If
 
    If cboCls.MatchFound Then
        lbl_cls.Caption = cboCls.List(cboCls.ListIndex, 1)
        LblErrMsg.Caption = ""
        
        Select Case cboCls.ListIndex
        'Case 0: sqlfg = ""
        Case 0:
            sqlfg = " finishgoodpart_cls = '01' "
            CheckBox1.Enabled = False
        Case 1:
            sqlfg = " finishgoodpart_cls = '02' "
            CheckBox1.Enabled = True
        End Select
        
        'CariProduk
    Else
        lbl_cls.Caption = ""
        LblErrMsg.Caption = ""
        If cboCls.Text <> "" Then LblErrMsg = DisplayMsg(29)   'Invalid Finish Good Part Clasification !
    End If
    
    Call Header
End Sub

Private Sub CboMb_Click()
    Lbl_Make.Caption = CboMb.List(CboMb.ListIndex, 1)
    LblErrMsg.Caption = ""

    Select Case CboMb.ListIndex
        Case 0: sqlmb = ""
        Case 1: sqlmb = " makebuy_cls = '01' "
        Case 2: sqlmb = " makebuy_cls = '02' "
    End Select
    'CariProduk
    Call Header
End Sub

Private Sub CboProduk_Change()
    If CboProduk.Text = "" Then LblErrMsg.Caption = ""
    
    If CboProduk.MatchFound Then
     lblPartNumber.Caption = CboProduk.List(CboProduk.ListIndex, 1)
     lblNm.Caption = CboProduk.List(CboProduk.ListIndex, 2)
     LblErrMsg.Caption = ""
    Else
     lblPartNumber.Caption = ""
     lblNm.Caption = ""
     If CboProduk.Text <> "" Then LblErrMsg.Caption = DisplayMsg(4061) 'Product Code not found!
    End If
    
    Call Header
End Sub

Private Sub CboProduk_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = vbKeyReturn Then CboProduk_Change
End Sub

Private Sub Header()
LblRecord = "0"
Dim RS As New ADODB.Recordset
Dim i As Integer

    With grid
        .Rows = 2
        .ColS = 20
        .FixedRows = 2
                        
        .TextMatrix(0, ColDate) = "Date"
        .TextMatrix(1, ColDate) = "Date"
        
        .TextMatrix(0, ColCls) = "Cls"
        .TextMatrix(1, ColCls) = "Cls"
        
        .TextMatrix(0, ColFromLoc) = "From" & vbCrLf & "Location"
        .TextMatrix(1, ColFromLoc) = "From" & vbCrLf & "Location"
        
        .TextMatrix(0, ColToLoc) = "To" & vbCrLf & "Location"
        .TextMatrix(1, ColToLoc) = "To" & vbCrLf & "Location"
        
        .TextMatrix(0, ColDoNo) = "DO No."
        .TextMatrix(1, ColDoNo) = "DO No."
        
        .TextMatrix(0, ColPONo) = "PO No."
        .TextMatrix(1, ColPONo) = "PO No."
        
        .TextMatrix(0, ColLotNo) = "Lot No."
        .TextMatrix(1, ColLotNo) = "Lot No."
        
        .TextMatrix(0, ColDesc) = "Description"
        .TextMatrix(1, ColDesc) = "Description"
        
        .TextMatrix(0, ColPremStock) = "Opening" + vbCrLf + "Balance"
        .TextMatrix(1, ColPremStock) = "Opening" + vbCrLf + "Balance"
        
        .TextMatrix(0, ColIncomeStock) = "Incoming Qty"
        .TextMatrix(0, ColIncomeOtherStock) = "Incoming Qty"
        .TextMatrix(1, ColIncomeStock) = "Incoming"
        .TextMatrix(1, ColIncomeOtherStock) = "Other"
        
        .TextMatrix(0, ColOutgoingStock) = "Outgoing Qty"
        .TextMatrix(0, ColOutgoingOtherStock) = "Outgoing Qty"
        .TextMatrix(1, ColOutgoingStock) = "Outgoing"
        .TextMatrix(1, ColOutgoingOtherStock) = "Other"
        
        .TextMatrix(0, ColCurrent) = "Ending" + vbCrLf + "Balance"
        .TextMatrix(1, ColCurrent) = "Ending" + vbCrLf + "Balance"
        
        .TextMatrix(0, ColCurr) = "Curr"
        .TextMatrix(1, ColCurr) = "Curr"
        
        .TextMatrix(0, ColPrice) = "Price"
        .TextMatrix(1, ColPrice) = "Price"
        
        .TextMatrix(0, ColRate) = "Rate"
        .TextMatrix(1, ColRate) = "Rate"
        
        .TextMatrix(0, ColAvgPrice) = "Average" + vbCrLf + "Price"
        .TextMatrix(1, ColAvgPrice) = "Average" + vbCrLf + "Price"
        
        .TextMatrix(0, ColAmount) = "Amount"
        .TextMatrix(1, ColAmount) = "Amount"
        
        .TextMatrix(0, ColValueInv) = "Value of" + vbCrLf + "Inventory"
        .TextMatrix(1, ColValueInv) = "Value of" + vbCrLf + "Inventory"
        
        .FixedRows = 2
        .MergeCells = flexMergeFixedOnly
        .FrozenCols = 3
    
        .ColWidth(ColDate) = 1300
        .ColWidth(ColCls) = 500
        .ColWidth(ColDesc) = 2000
        .ColWidth(ColFromLoc) = 800
        .ColWidth(ColToLoc) = 800
        .ColWidth(ColPONo) = 2000
        .ColWidth(ColDoNo) = 1300
        .ColWidth(ColLotNo) = 1600
        .ColWidth(ColPremStock) = 1500
        .ColWidth(ColIncomeStock) = 1500
        .ColWidth(ColIncomeOtherStock) = 1500
        .ColWidth(ColOutgoingStock) = 1500
        .ColWidth(ColOutgoingOtherStock) = 1500
        .ColWidth(ColCurrent) = 1500
        .ColWidth(ColCurr) = 500
        .ColWidth(ColPrice) = 1500
        .ColWidth(ColRate) = 1500
        .ColWidth(ColAvgPrice) = 1500
        .ColWidth(ColAmount) = 2000
        .ColWidth(ColValueInv) = 2000
        
        .Cell(flexcpAlignment, 0, ColDate, 1, ColValueInv) = flexAlignCenterCenter
        .ColAlignment(ColFromLoc) = flexAlignLeftCenter
        .ColAlignment(ColToLoc) = flexAlignLeftCenter
        .ColAlignment(ColPONo) = flexAlignLeftCenter
        .ColAlignment(ColDoNo) = flexAlignLeftCenter
        .ColAlignment(ColLotNo) = flexAlignLeftCenter
        .ColAlignment(ColDate) = flexAlignCenterCenter
    
        .MergeRow(0) = True
        .MergeCol(ColDate) = True
        .MergeCol(ColCls) = True
        .MergeCol(ColDesc) = True
        .MergeCol(ColFromLoc) = True
        .MergeCol(ColToLoc) = True
        .MergeCol(ColLotNo) = True
        .MergeCol(ColPONo) = True
        .MergeCol(ColDoNo) = True
        
        .MergeCol(ColPremStock) = True
        .MergeCol(ColCurrent) = True
        .MergeCol(ColCurr) = True
        .MergeCol(ColPrice) = True
        .MergeCol(ColRate) = True
        .MergeCol(ColAvgPrice) = True
        .MergeCol(ColAmount) = True
        .MergeCol(ColValueInv) = True
        .MergeCells = flexMergeFixedOnly
        .FrozenCols = 2
        
    End With

End Sub

Private Sub CmdExcel_Click()
Dim xlapp As New Excel.application

If grid.Rows > 2 Then

        LblErrMsg = ""
        Me.MousePointer = vbHourglass
        
        Dim Idx As Integer

        Dim xlColDate As String
        Dim xlColCls As String
        Dim xlColDesc As String
        Dim xlColFromLoc As String
        Dim xlColToLoc As String
        Dim xlColPONo As String
        Dim xlColDoNo As String
        Dim xlColLotNo As String
        Dim xlColPremStock As String
        Dim xlColIncomeStock As String
        Dim xlColIncomeOtherStock As String
        Dim xlColOutgoingStock As String
        Dim xlColOutgoingOtherStock As String
        Dim xlColCurrent As String
        Dim xlColCurr As String
        Dim xlColPrice As String
        Dim xlColRate As String
        Dim xlColAvgPrice As String
        Dim xlColAmount As String
        Dim xlColValueInv As String
                        
         xlColDate = "a"
         xlColCls = "b"
         xlColDesc = "c"
         xlColFromLoc = "d"
         xlColToLoc = "e"
         xlColPONo = "f"
         xlColDoNo = "g"
         xlColLotNo = "h"
         xlColPremStock = "i"
         xlColIncomeStock = "j"
         xlColIncomeOtherStock = "k"
         xlColOutgoingStock = "l"
         xlColOutgoingOtherStock = "m"
         xlColCurrent = "n"
         xlColCurr = "o"
         xlColPrice = "p"
         xlColRate = "q"
         xlColAvgPrice = "r"
         xlColAmount = "s"
         xlColValueInv = "t"

        With xlapp
            .Workbooks.Add
     
            .Range(xlColDate & "2", xlColValueInv & "2").Merge
            .Range(xlColDate & "2") = "VALUATION PRICE REPORT DETAIL"
            .Range(xlColDate & "2").horizontalAlignment = xlCenter
            .Range(xlColDate & "2").Font.Bold = True
                 
            .Range(xlColDate & "3:" & xlColValueInv & "3").Merge
            .Range(xlColDate & "3") = "Base Currency : " + lblBase
            .Range(xlColDate & "3").horizontalAlignment = xlCenter
            .Range(xlColDate & "3").Font.Bold = False
     
            .Range(xlColDate & "4") = "Finish Good Part Cls " + " : " + cboCls.Text + " / " + lbl_cls.Caption
            .Range(xlColDate & "5") = "Make/Buy Cls " + " : " + CboMb.Text + " / " + Lbl_Make.Caption
            .Range(xlColDate & "6").NumberFormat = "@"
            .Range(xlColDate & "6") = "Product Code " + " : " + Trim(TxtProduk) + " / " + lblPartNumber.Caption + " / " + lblNm.Caption
            .Range(xlColDate & "7") = "Period " + " : " + Format(DtPeriod.Value, "MMM yyyy")
            .Range(xlColRate & "7", xlColValueInv & "7").Merge
            .Range(xlColRate & "7") = "Issued Date : " + Format(Now, "dd MMM yyyy  hh:MM:ss")
            .Range(xlColDate & "4", xlColIncomeStock & "4").Merge
            .Range(xlColDate & "5", xlColIncomeStock & "5").Merge
            .Range(xlColDate & "6", xlColIncomeStock & "6").Merge
            .Range(xlColDate & "7", xlColIncomeStock & "7").Merge
         
            .Range(xlColDate & "9", xlColDate & "10").Merge
            .Range(xlColCls & "9", xlColCls & "10").Merge
            .Range(xlColDesc & "9", xlColDesc & "10").Merge
            .Range(xlColFromLoc & "9", xlColFromLoc & "10").Merge
            .Range(xlColToLoc & "9", xlColToLoc & "10").Merge
            .Range(xlColPONo & "9", xlColPONo & "10").Merge
            .Range(xlColDoNo & "9", xlColDoNo & "10").Merge
            .Range(xlColLotNo & "9", xlColLotNo & "10").Merge
            .Range(xlColPremStock & "9", xlColPremStock & "10").Merge
            .Range(xlColIncomeStock & "9", xlColIncomeOtherStock & "9").Merge
            .Range(xlColOutgoingStock & "9", xlColOutgoingOtherStock & "9").Merge
            .Range(xlColCurrent & "9", xlColCurrent & "10").Merge
            .Range(xlColCurr & "9", xlColCurr & "10").Merge
            .Range(xlColPrice & "9", xlColPrice & "10").Merge
            .Range(xlColRate & "9", xlColRate & "10").Merge
            .Range(xlColAvgPrice & "9", xlColAvgPrice & "10").Merge
            .Range(xlColAmount & "9", xlColAmount & "10").Merge
            .Range(xlColValueInv & "9", xlColValueInv & "10").Merge
                         
            .Range(xlColDate & "9") = "Date"
            .Range(xlColCls & "9") = "Cls"
            .Range(xlColDesc & "9") = "Description"
            .Range(xlColFromLoc & "9") = "From Loc."
            .Range(xlColToLoc & "9") = "To Loc."
            .Range(xlColPONo & "9") = "PO No."
            .Range(xlColDoNo & "9") = "DO No."
            .Range(xlColLotNo & "9") = "Lot No."
            .Range(xlColPremStock & "9") = "Opening" & Chr(10) & "Balance"
            .Range(xlColIncomeStock & "9") = "Incoming Qty"
            .Range(xlColIncomeStock & "10") = "Incoming"
            .Range(xlColIncomeOtherStock & "10") = "Other"
            .Range(xlColOutgoingStock & "9") = "Outgoing Qty"
            .Range(xlColOutgoingStock & "10") = "Outgoing"
            .Range(xlColOutgoingOtherStock & "10") = "Other"
            .Range(xlColCurrent & "9") = "Ending"
            .Range(xlColCurr & "9") = "Curr"
            .Range(xlColPrice & "9") = "Price"
            .Range(xlColRate & "9") = "Rate"
            .Range(xlColAvgPrice & "9") = "Average" & Chr(10) & "Price"
            .Range(xlColAmount & "9") = "Amount"
            .Range(xlColValueInv & "9") = "Value of" & Chr(10) & "Inventory"
                        
            .Range(xlColDate & "9", xlColValueInv & "10").horizontalAlignment = xlCenter
            .Range(xlColDate & "9", xlColValueInv & "10").verticalAlignment = xlCenter
            
            Idx = 10
            
            '#Fill Grid
            For i = 2 To grid.Rows - 1
                Idx = Idx + 1
                .Range(xlColDate & Idx) = grid.TextMatrix(i, ColDate)
                .Range(xlColCls & Idx) = grid.TextMatrix(i, ColCls)
                .Range(xlColDesc & Idx) = grid.TextMatrix(i, ColDesc)
                .Range(xlColFromLoc & Idx) = grid.TextMatrix(i, ColFromLoc)
                .Range(xlColToLoc & Idx) = grid.TextMatrix(i, ColToLoc)
                .Range(xlColPONo & Idx) = grid.TextMatrix(i, ColPONo)
                .Range(xlColDoNo & Idx) = grid.TextMatrix(i, ColDoNo)
                .Range(xlColLotNo & Idx) = grid.TextMatrix(i, ColLotNo)
                .Range(xlColPremStock & Idx) = grid.TextMatrix(i, ColPremStock)
                .Range(xlColIncomeStock & Idx) = grid.TextMatrix(i, ColIncomeStock)
                .Range(xlColIncomeOtherStock & Idx) = grid.TextMatrix(i, ColIncomeOtherStock)
                .Range(xlColOutgoingStock & Idx) = grid.TextMatrix(i, ColOutgoingStock)
                .Range(xlColOutgoingOtherStock & Idx) = grid.TextMatrix(i, ColOutgoingOtherStock)
                .Range(xlColCurrent & Idx) = grid.TextMatrix(i, ColCurrent)
                .Range(xlColCurr & Idx) = grid.TextMatrix(i, ColCurr)
                .Range(xlColPrice & Idx) = grid.TextMatrix(i, ColPrice)
                
                If grid.TextMatrix(i, ColCurr) = "IDR" Then
                 .Range(xlColPrice & Idx).NumberFormat = gs_formatPriceIDR
                Else
                 .Range(xlColPrice & Idx).NumberFormat = gs_formatPrice
                End If
                
                .Range(xlColRate & Idx) = grid.TextMatrix(i, ColRate)
                .Range(xlColAvgPrice & Idx) = grid.TextMatrix(i, ColAvgPrice)
                .Range(xlColAmount & Idx) = grid.TextMatrix(i, ColAmount)
                .Range(xlColValueInv & Idx) = grid.TextMatrix(i, ColValueInv)
            Next


            '#Run Macro
            .Range(xlColDate & "2:" & xlColAmount & "2").Select
            With .Selection.Font
                .Size = 18
            End With
            .Columns(xlColDate & ":" & xlColDate).columnWidth = 11.57
            .Columns(xlColCls & ":" & xlColCls).EntireColumn.AutoFit
            .Columns(xlColDesc & ":" & xlColDesc).columnWidth = 19.43
            .Columns(xlColFromLoc & ":" & xlColFromLoc).columnWidth = 11.43
            .Columns(xlColToLoc & ":" & xlColToLoc).columnWidth = 11
            .Columns(xlColPONo & ":" & xlColPONo).columnWidth = 18.86
            .Columns(xlColDoNo & ":" & xlColDoNo).columnWidth = 12.14
            .Columns(xlColLotNo & ":" & xlColLotNo).columnWidth = 12.86
            .Columns(xlColPremStock & ":" & xlColPremStock).columnWidth = 12.86
            .Columns(xlColIncomeStock & ":" & xlColIncomeStock).columnWidth = 12.86
            .Columns(xlColIncomeOtherStock & ":" & xlColIncomeOtherStock).columnWidth = 12.86
            .Columns(xlColOutgoingStock & ":" & xlColOutgoingStock).columnWidth = 12.86
            .Columns(xlColOutgoingOtherStock & ":" & xlColOutgoingOtherStock).columnWidth = 12.86
            .Columns(xlColCurrent & ":" & xlColCurrent).columnWidth = 12.86
            .Columns(xlColCurr & ":" & xlColCurr).columnWidth = 5
            .Columns(xlColPrice & ":" & xlColPrice).columnWidth = 12.86
            .Columns(xlColRate & ":" & xlColRate).columnWidth = 12.86
            .Columns(xlColAvgPrice & ":" & xlColAvgPrice).columnWidth = 12.86
            .Columns(xlColAmount & ":" & xlColAmount).columnWidth = 15
            .Columns(xlColValueInv & ":" & xlColValueInv).columnWidth = 15
            
            .Columns(xlColDoNo & ":" & xlColDoNo).Select
            With .Selection
                .horizontalAlignment = xlLeft
            End With
            .Columns(xlColPONo & ":" & xlColPONo).Select
            With .Selection
                .horizontalAlignment = xlLeft
            End With
            .Columns(xlColToLoc & ":" & xlColToLoc).Select
            With .Selection
                .horizontalAlignment = xlLeft
            End With
            .Columns(xlColFromLoc & ":" & xlColFromLoc).Select
            With .Selection
                .horizontalAlignment = xlLeft
            End With
            .Columns(xlColCls & ":" & xlColCls).Select
            With .Selection
                .horizontalAlignment = xlCenter
            End With
            .Columns(xlColDate & ":" & xlColDate).Select
            With .Selection
                .horizontalAlignment = xlCenter
            End With
            .Range(xlColDate & "9:" & xlColValueInv & "10").Select

            With .Selection
                .horizontalAlignment = xlCenter
                .verticalAlignment = xlCenter
                With .Interior
                 .Pattern = xlSolid
                 .PatternColorIndex = xlAutomatic
                 .color = 15773696
                 .TintAndShade = 0
                 .PatternTintAndShade = 0
                End With
            End With
            
            .Range(xlColDate & "1:" & xlColValueInv & Idx).Select
            .Selection.Font.Name = "Arial"
            
            .Range(xlColDate & "3:" & xlColValueInv & Idx).Select
            .Selection.Font.Size = 10
            
            .Range(xlColDate & "9:" & xlColValueInv & Idx).Select
            
            .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
            .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
            With .Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            With .Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            With .Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            With .Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            With .Selection.Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            With .Selection.Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With

            .Range(xlColDate & "4:" & xlColDesc & "7").Select
              With .Selection
                  .horizontalAlignment = xlLeft
                  .verticalAlignment = xlCenter
              End With
              
              .Range(xlColPremStock & "11:" & xlColValueInv & Idx).Select
              With .Selection
                  .horizontalAlignment = xlRight
                  .verticalAlignment = xlCenter
              End With

              .Range(xlColCurr & "11:" & xlColCurr & Idx).Select
              With .Selection
                  .horizontalAlignment = xlCenter
                  .verticalAlignment = xlCenter
              End With

              .Range(xlColDesc & "11:" & xlColLotNo & Idx).Select
              With .Selection
                  .horizontalAlignment = xlLeft
                  .verticalAlignment = xlCenter
              End With

              .Range(xlColPremStock & "11:" & xlColCurrent & Idx - 2).Select
              With .Selection
                  .NumberFormat = gs_formatQty
              End With

              .Range(xlColRate & "11:" & xlColRate & Idx - 3).Select
              With .Selection
                  .NumberFormat = gs_formatExchangeRate
              End With

              .Range(xlColAvgPrice & "11:" & xlColAvgPrice & Idx - 3).Select
              With .Selection
                  .NumberFormat = gs_formatPriceIDR
              End With
                            
              .Range(xlColAmount & "11:" & xlColAmount & Idx - 3).Select
              With .Selection
                  .NumberFormat = gs_formatAmountIDR
              End With
              
              .Range(xlColValueInv & "11:" & xlColValueInv & Idx - 3).Select
              With .Selection
                  .NumberFormat = gs_formatAmountIDR
              End With
              
              .Range(xlColPremStock & Idx - 2 & ":" & xlColCurrent & Idx - 1).Select
              With .Selection
                  .NumberFormat = gs_formatAmountIDR
              End With
              
              .Range(xlColCurrent & Idx & ":" & xlColCurrent & Idx).Select
              With .Selection
                  .NumberFormat = gs_formatPriceIDR
              End With
'
'                .Range(xlColPremStock & "11:" & xlColAmount & Idx).Select
'              With .Selection
'                  .NumberFormat = gs_formatPriceIDR
'              End With
'
            'Issued Date
            .Range(xlColRate & "7").horizontalAlignment = xlHAlignRight
            
            .ActiveSheet.PageSetup.PaperSize = xlPaperA4
            .ActiveSheet.PageSetup.Orientation = 2
            .WindowState = xlMaximized
            .ActiveWindow.Zoom = 80
            .Visible = True
            End With
        
        Me.MousePointer = vbDefault
        
End If

End Sub

Private Sub cmdReport_Click()
Dim tempExc As String

'If CboProduk.ListIndex < 0 Then
'    LblErrMsg = DisplayMsg(4061) '"Please Input Product Code!"
'    Exit Sub
'End If

If lblPartNumber.Caption = "Part Number" Or lblPartNumber = "" Then
    LblErrMsg = DisplayMsg(4061) '"Please Input Product Code!"
    Exit Sub
End If

7 LblErrMsg = ""

'lblErrMsg = up_ValidateDateRange(Format(DtPeriod, "yyyy-MM-dd"), False)

LblRecord = Format(0, "#,##0")

If LblErrMsg <> "" Then
    Exit Sub
End If

If CheckBox1.Value = True And (cboCls = "ALL" Or cboCls = "02") Then
    tempExc = " and duty_status in('1','3') "
ElseIf CheckBox1.Value = False And (cboCls = "ALL" Or cboCls = "02") Then
    tempExc = " and duty_status in('0','3') "
Else
    tempExc = ""
End If

Screen.MousePointer = vbHourglass

    SqlData = " DECLARE  @CompanyCode AS CHAR(5) declare @Period char(10)    " & _
              vbLf & "declare @ItemCode as char(15) " & _
              vbLf & "declare @EndMonthDate char(10) " & _
              vbLf & "declare @Year char(4) " & _
              vbLf & "declare @Month varchar(2) " & _
              vbLf & "declare @YearMonth varchar(6) " & _
              vbLf & "declare @RateCls char(1) " & _
              vbLf & "declare @digitexchrate int "
              
    SqlData = SqlData & _
              vbLf & " " & _
              vbLf & "Set @ItemCode='" & Trim(TxtProduk.Text) & "' " & _
              vbLf & "Set @Period ='" & Format(DtPeriod.Value, "yyyy-MM-dd") & "' " & _
              vbLf & "Set @Year = year(@Period) " & _
              vbLf & "Set @Month = Month(@Period) " & _
              vbLf & "Set @YearMonth = cast(@year as varchar) + right(cast((@month + 100) as varchar),2)  " & _
              vbLf & "Set @EndMonthDate = convert(char(10),dateadd(day,-1,dateadd(month, 1,cast(@Year as char(4)) + '-' +  cast(@Month as varchar(2)) + '-01')),120) " & _
              vbLf & "Set @RateCls = (select top 1 rate_cls from company_profile) -- If 0 then DailyExchangeRate else BookKeepingExchangeRate " & _
              vbLf & "Set @digitexchrate = " & gi_decimalDigitExchangeRate
              
    SqlData = SqlData & _
              vbLf & "" & _
              vbLf & "SELECT * FROM " & _
              vbLf & "(" & _
              vbLf & "SELECT ipd.tgl, ipd.cls,  " & _
              vbLf & "description = case ipd.cls " & _
              vbLf & "              when 'OP' then 'Opening' " & _
              vbLf & "              when 'R' then 'Receipt' " & _
              vbLf & "              when 'R1' then 'Return' " & _
              vbLf & "              when 'S' then 'Consumption' " & _
              vbLf & "              when 'RJ' then 'Reject' " & _
              vbLf & "              when 'D' then 'Delivery Order' " & _
              vbLf & "              when 'L' then 'Loss' " & _
              vbLf & "              when 'P1' then 'Production' " & _
              vbLf & "              when 'Inv' then 'Inventory Different' " & _
              vbLf & "              else '' end, " & _
              vbLf & "ipd.from_loc, ipd.to_loc, ipd.po_no, ipd.do_no, ipd.lot_no, " & _
              vbLf & "ipd.QtyOP, ipd.QtyIn, ipd.QtyInOt, ipd.QtyOut, ipd.QtyOutOt, ipd.QtyEnding, isnull(cc.description,'') curr, " & _
              vbLf & "ipd.Price, rate = round((case @ratecls when '0' then dbo.UF_GetDailyExchangeRate(ipd.tgl,ipd.Currency_Code) else dbo.UF_GetBookExchangeRate (Year(ipd.tgl),Month(ipd.tgl),ipd.Currency_Code) end),@digitexchrate), " & _
              vbLf & "ipd.AvgPrice, ipd.Amount, ipd.AmountEnding " & _
              vbLf & "-- TAMBAHAN KOLOM BARU UNTUK SORT 2013-10-16 " & _
              vbLf & ",idxGrp = 0 " & _
              vbLf & ",idxLkl = ipd.idx, ipd.Seq_No  "
                    
    SqlData = SqlData & _
              vbLf & "" & _
              vbLf & "From Inventory_PriceDetail ipd " & _
              vbLf & "Left Join Curr_Cls cc " & _
              vbLf & "On ipd.currency_code = cc.curr_cls " & _
              vbLf & "where rtrim(ipd.item_code) = rtrim(@itemCode) " & _
              vbLf & "and rtrim(ipd.period) = @yearmonth " & tempExc & ""
              
    SqlData = SqlData & _
              vbLf & "" & _
              vbLf & "Union ALL " & _
              vbLf & "/*###### Ending #######*/ " & _
              vbLf & "SELECT tgl = @EndMonthDate, cls = 'End', " & _
              vbLf & "description = 'Ending', from_loc = '', to_loc = '', po_no = '', do_no = '', lot_no = '', " & _
              vbLf & "QtyOP = 0, QtyIn = 0, QtyInOt = 0, QtyOut = 0, QtyOutOt = 0, QtyEnding = ip.Current_Stock, " & _
              vbLf & "curr = '', Price = 0, rate = 0, AvgPrice = ip.Current_Price, Amount = 0, " & _
              vbLf & "AmountEnding = (select isnull(amountending,0) from Inventory_PriceDetail where month(tgl) = @month and year(tgl) = @year and item_code = @itemcode and cls = 'Inv' " & tempExc & ") " & _
              vbLf & "-- TAMBAHAN KOLOM BARU UNTUK SORT 2013-10-16 " & _
              vbLf & ",idxGrp = 1 " & _
              vbLf & ",idxLkl = 0, '99'SeqNo " & _
              vbLf & "From  Inventory_Price ip " & _
              vbLf & "where rtrim(ip.item_code) = rtrim(@itemCode) " & _
              vbLf & "and inventory_year = @year " & _
              vbLf & "and inventory_month = @month " & tempExc & ""
              
    SqlData = SqlData & _
              vbLf & "" & _
              vbLf & "Union ALL " & _
              vbLf & "/*###### Total Qty Vertical #######*/ " & _
              vbLf & "SELECT tgl = @EndMonthDate, cls = 'TQV', " & _
              vbLf & "description = '', from_loc = '', to_loc = '', po_no = '', do_no = '', lot_no = 'Total Qty', " & _
              vbLf & "TotQtyOP = ip.premonth_stock, TotQtyIn = ip.incoming_stock, TotQtyInOt = ip.incomingother_stock, " & _
              vbLf & "TotQtyOut = ip.outgoing_stock, TotQtyOutOt = ip.outgoingother_stock, " & _
              vbLf & "TotQtyEnding = ip.current_stock, " & _
              vbLf & "curr = '', Price = 0, rate = 0, AvgPrice = 0, Amount = 0, " & _
              vbLf & "AmountEnding = 0 " & _
              vbLf & ",idxGrp = 2 " & _
              vbLf & ",idxLkl = 0, '99'SeqNo " & _
              vbLf & "From  Inventory_Price ip " & _
              vbLf & "where rtrim(ip.item_code) = rtrim(@itemCode) " & _
              vbLf & "and inventory_year = @year " & _
              vbLf & "and inventory_month = @month " & tempExc & ""
              
    SqlData = SqlData & _
              vbLf & "" & _
              vbLf & "Union ALL " & _
              vbLf & "/*###### Total Amount Vertical #######*/ " & _
              vbLf & "SELECT tgl = @EndMonthDate, cls = 'TAV', " & _
              vbLf & "description = '', from_loc = '', to_loc = '', po_no = '', do_no = '', lot_no = 'Total Amount', " & _
              vbLf & "TotAmountOP = ip.premonth_stock * ip.premonth_price, " & _
              vbLf & "TotAmountIn = ip.incoming_stock * ip.incoming_price, " & _
              vbLf & "TotAmountInOt = ip.incomingother_stock * ip.incomingother_price, " & _
              vbLf & "TotAmountOut = ip.outgoing_stock * ip.outgoing_price, " & _
              vbLf & "TotAmountOutOt = ip.outgoingother_stock * ip.outgoingother_price, " & _
              vbLf & "TotAmountEnding = (select isnull(amountending,0) from Inventory_PriceDetail where  month(tgl) = @month and year(tgl) = @year and item_code = @itemcode and cls = 'Inv' " & tempExc & "), " & _
              vbLf & "curr = '', Price = 0, rate = 0, AvgPrice = 0, Amount = 0, " & _
              vbLf & "AmountEnding = 0 " & _
              vbLf & ",idxGrp = 3 " & _
              vbLf & ",idxLkl = 0, '99'SeqNo " & _
              vbLf & "From  Inventory_Price ip " & _
              vbLf & "where rtrim(ip.item_code) = rtrim(@itemCode) " & _
              vbLf & "and inventory_year = @year " & _
              vbLf & "and inventory_month = @month " & tempExc & ""
              
    SqlData = SqlData & _
              vbLf & " Union ALL " & _
              vbLf & "/*###### Avg Price #######*/ " & _
              vbLf & "SELECT tgl = @EndMonthDate, cls = 'AVG', " & _
              vbLf & "description = '', from_loc = '', to_loc = '', po_no = '', do_no = '', lot_no = 'Average Price', " & _
              vbLf & "QtyOP = 0, QtyIn = 0, QtyInOt = 0, QtyOut = 0, QtyOutOt = 0, QtyEnding = ip.Current_Price, " & _
              vbLf & "curr = '', Price = 0, rate = 0, AvgPrice = 0, Amount = 0, " & _
              vbLf & "AmountEnding = 0 " & _
              vbLf & ",idxGrp = 4 " & _
              vbLf & ",idxLkl = 0, '99'SeqNo " & _
              vbLf & "From  Inventory_Price ip " & _
              vbLf & "where rtrim(ip.item_code) = rtrim(@itemCode) " & _
              vbLf & "and inventory_year = @year " & _
              vbLf & "and inventory_month = @month " & tempExc & ""

    ' ###########################
    ' 20130730 - Get Data From History
    ' ###########################
    
    SqlData = SqlData & _
              vbLf & "" & _
              vbLf & " --- History ---" & _
              vbLf & " UNION ALL" & _
              vbLf & "SELECT ipd.tgl, ipd.cls,  " & _
              vbLf & "description = case ipd.cls " & _
              vbLf & "              when 'OP' then 'Opening' " & _
              vbLf & "              when 'R' then 'Receipt' " & _
              vbLf & "              when 'R1' then 'Return' " & _
              vbLf & "              when 'S' then 'Consumption' " & _
              vbLf & "              when 'RJ' then 'Reject' " & _
              vbLf & "              when 'D' then 'Delivery Order' " & _
              vbLf & "              when 'L' then 'Loss' " & _
              vbLf & "              when 'P1' then 'Production' " & _
              vbLf & "              when 'Inv' then 'Inventory Different' " & _
              vbLf & "              else '' end, " & _
              vbLf & "ipd.from_loc, ipd.to_loc, ipd.po_no, ipd.do_no, ipd.lot_no, " & _
              vbLf & "ipd.QtyOP, ipd.QtyIn, ipd.QtyInOt, ipd.QtyOut, ipd.QtyOutOt, ipd.QtyEnding, isnull(cc.description,'') curr, " & _
              vbLf & "ipd.Price, rate = round((case @ratecls when '0' then dbo.UF_GetDailyExchangeRate(ipd.tgl,ipd.Currency_Code) else dbo.UF_GetBookExchangeRate (Year(ipd.tgl),Month(ipd.tgl),ipd.Currency_Code) end),@digitexchrate), " & _
              vbLf & "ipd.AvgPrice, ipd.Amount, ipd.AmountEnding " & _
              vbLf & ",idxGrp = 0 " & _
              vbLf & ",idxLkl = ipd.idx, '99' SeqNo "
                    
    SqlData = SqlData & _
              vbLf & "" & _
              vbLf & "From InventoryPriceDetail_History ipd " & _
              vbLf & "Left Join Curr_Cls cc " & _
              vbLf & "On ipd.currency_code = cc.curr_cls " & _
              vbLf & "where rtrim(ipd.item_code) = rtrim(@itemCode) " & _
              vbLf & "and ipd.period = CONVERT(CHAR(6),CAST(@period AS datetime) ,112) " & tempExc & ""
              
    SqlData = SqlData & _
              vbLf & "" & _
              vbLf & "Union ALL " & _
              vbLf & "/*###### Ending #######*/ " & _
              vbLf & "SELECT tgl = @EndMonthDate, cls = 'End', " & _
              vbLf & "description = 'Ending', from_loc = '', to_loc = '', po_no = '', do_no = '', lot_no = '', " & _
              vbLf & "QtyOP = 0, QtyIn = 0, QtyInOt = 0, QtyOut = 0, QtyOutOt = 0, QtyEnding = ip.Current_Stock, " & _
              vbLf & "curr = '', Price = 0, rate = 0, AvgPrice = ip.Current_Price, Amount = 0, " & _
              vbLf & "AmountEnding = (select isnull(amountending,0) from InventoryPriceDetail_History where month(tgl) = @month and year(tgl) = @year and item_code = @itemcode and cls = 'Inv' " & tempExc & " and coalesce(amountending,0)<>0 )  " & _
              vbLf & ",idxGrp = 1 " & _
              vbLf & ",idxLkl = 0, '99'SeqNo " & _
              vbLf & "From  InventoryPrice_History ip " & _
              vbLf & "where rtrim(ip.item_code) = rtrim(@itemCode) " & _
              vbLf & "and inventory_year = @year " & _
              vbLf & "and inventory_month = @month " & tempExc & ""
              
    SqlData = SqlData & _
              vbLf & "" & _
              vbLf & "Union ALL " & _
              vbLf & "/*###### Total Qty Vertical #######*/ " & _
              vbLf & "SELECT tgl = @EndMonthDate, cls = 'TQV', " & _
              vbLf & "description = '', from_loc = '', to_loc = '', po_no = '', do_no = '', lot_no = 'Total Qty', " & _
              vbLf & "TotQtyOP = ip.premonth_stock, TotQtyIn = ip.incoming_stock, TotQtyInOt = ip.incomingother_stock, " & _
              vbLf & "TotQtyOut = ip.outgoing_stock, TotQtyOutOt = ip.outgoingother_stock, " & _
              vbLf & "TotQtyEnding = ip.current_stock, " & _
              vbLf & "curr = '', Price = 0, rate = 0, AvgPrice = 0, Amount = 0, " & _
              vbLf & "AmountEnding = 0 " & _
              vbLf & ",idxGrp = 2 " & _
              vbLf & ",idxLkl = 0, '99'SeqNo " & _
              vbLf & "From  InventoryPrice_History ip " & _
              vbLf & "where rtrim(ip.item_code) = rtrim(@itemCode) " & _
              vbLf & "and inventory_year = @year " & _
              vbLf & "and inventory_month = @month " & tempExc & ""
              
    SqlData = SqlData & _
              vbLf & "" & _
              vbLf & "Union ALL " & _
              vbLf & "/*###### Total Amount Vertical #######*/ " & _
              vbLf & "SELECT tgl = @EndMonthDate, cls = 'TAV', " & _
              vbLf & "description = '', from_loc = '', to_loc = '', po_no = '', do_no = '', lot_no = 'Total Amount', " & _
              vbLf & "TotAmountOP = ip.premonth_stock * ip.premonth_price, " & _
              vbLf & "TotAmountIn = ip.incoming_stock * ip.incoming_price, " & _
              vbLf & "TotAmountInOt = ip.incomingother_stock * ip.incomingother_price, " & _
              vbLf & "TotAmountOut = ip.outgoing_stock * ip.outgoing_price, " & _
              vbLf & "TotAmountOutOt = ip.outgoingother_stock * ip.outgoingother_price, " & _
              vbLf & "TotAmountEnding = (select isnull(amountending,0) from InventoryPriceDetail_History where month(tgl) = @month and year(tgl) = @year and item_code = @itemcode and cls = 'Inv' " & tempExc & " and coalesce(amountending,0)<>0  ), " & _
              vbLf & "curr = '', Price = 0, rate = 0, AvgPrice = 0, Amount = 0, " & _
              vbLf & "AmountEnding = 0 " & _
              vbLf & ",idxGrp = 3 " & _
              vbLf & ",idxLkl = 0, '99'SeqNo " & _
              vbLf & "From  InventoryPrice_History ip " & _
              vbLf & "where rtrim(ip.item_code) = rtrim(@itemCode) " & _
              vbLf & "and inventory_year = @year " & _
              vbLf & "and inventory_month = @month " & tempExc & ""
              
    SqlData = SqlData & _
              vbLf & " Union ALL " & _
              vbLf & "/*###### Avg Price #######*/ " & _
              vbLf & "SELECT tgl = @EndMonthDate, cls = 'AVG', " & _
              vbLf & "description = '', from_loc = '', to_loc = '', po_no = '', do_no = '', lot_no = 'Average Price', " & _
              vbLf & "QtyOP = 0, QtyIn = 0, QtyInOt = 0, QtyOut = 0, QtyOutOt = 0, QtyEnding = ip.Current_Price, " & _
              vbLf & "curr = '', Price = 0, rate = 0, AvgPrice = 0, Amount = 0, " & _
              vbLf & "AmountEnding = 0 " & _
              vbLf & ",idxGrp = 4 " & _
              vbLf & ",idxLkl = 0, '99'SeqNo " & _
              vbLf & "From  InventoryPrice_History ip " & _
              vbLf & "where rtrim(ip.item_code) = rtrim(@itemCode) " & _
              vbLf & "and inventory_year = @year " & _
              vbLf & "and inventory_month = @month " & tempExc & "" & _
              vbLf & "" & _
              "     )Q " & _
              vbLf & "ORDER BY q.tgl,q.idxgrp, q.idxlkl, Q.Seq_No "
              
    ' ###########################
    
    Dim RS As New ADODB.Recordset
    If RS.State <> adStateClosed Then RS.Close
    RS.CursorLocation = adUseClient
    RS.Open SqlData, Db, adOpenKeyset, adLockOptimistic
            Text1 = SqlData
'            Text1.Visible = True
    Header
    
    Dim i As Integer
    i = 1

    While RS.EOF = False
        i = i + 1
        grid.AddItem ""
        
        If Trim(RS!Cls) = "AVG" Then
          GoTo AVG
        ElseIf Trim(RS!Cls) = "TQV" Or Trim(RS!Cls) = "TAV" Then
          GoTo TV
        End If
        
        grid.TextMatrix(i, ColDate) = Format(RS!Tgl, "dd MMM yyyy")
        grid.TextMatrix(i, ColCls) = Trim(RS!Cls)
        grid.TextMatrix(i, ColDesc) = Trim(RS!Description)
        grid.TextMatrix(i, ColFromLoc) = Trim(RS!from_loc)
        grid.TextMatrix(i, ColToLoc) = Trim(RS!to_loc)
        grid.TextMatrix(i, ColPONo) = Trim(RS!po_no)
        
        If Trim(RS!Cls) = "D" Or Trim(RS!Cls) = "S" Or Trim(RS!Cls) = "P1" Or Trim(RS!Lot_no) = "Subcon" Then 'Delivery, Production, Receipt Subcon
         grid.TextMatrix(i, ColDoNo) = Trim(RS!do_no)
        End If
        
        If Trim(RS!Cls) = "R" And Trim(RS!Lot_no) <> "Subcon" Then 'Receipt NonSubcon
         grid.TextMatrix(i, ColCurr) = Trim(RS!Curr)
        
         If Trim(RS!Curr) = "IDR" Then
          grid.TextMatrix(i, ColPrice) = Format(RS!Price, gs_formatPriceIDR)
         Else
          grid.TextMatrix(i, ColPrice) = Format(RS!Price, gs_formatPrice)
         End If
        
         grid.TextMatrix(i, ColRate) = Format(RS!rate, gs_formatExchangeRate)
        End If
        
        grid.TextMatrix(i, ColAvgPrice) = Format(RS!AvgPrice, gs_formatPriceIDR)
        grid.TextMatrix(i, ColAmount) = Format(RS!Amount, gs_formatAmountIDR)
        grid.TextMatrix(i, ColValueInv) = Format(RS!AmountEnding, gs_formatAmountIDR)
        
TV:     'Total Vertical
        If Trim(RS!Cls) = "TAV" Then
         grid.TextMatrix(i, ColPremStock) = Format(RS!QtyOP, gs_formatAmountIDR)
         grid.TextMatrix(i, ColIncomeStock) = Format(RS!QtyIn, gs_formatAmountIDR)
         grid.TextMatrix(i, ColIncomeOtherStock) = Format(RS!QtyInOt, gs_formatAmountIDR)
         grid.TextMatrix(i, ColOutgoingStock) = Format(RS!qtyOut, gs_formatAmountIDR)
         grid.TextMatrix(i, ColOutgoingOtherStock) = Format(RS!QtyOutOt, gs_formatAmountIDR)
        Else
         grid.TextMatrix(i, ColPremStock) = Format(RS!QtyOP, gs_formatQty)
         grid.TextMatrix(i, ColIncomeStock) = Format(RS!QtyIn, gs_formatQty)
         grid.TextMatrix(i, ColIncomeOtherStock) = Format(RS!QtyInOt, gs_formatQty)
         grid.TextMatrix(i, ColOutgoingStock) = Format(RS!qtyOut, gs_formatQty)
         grid.TextMatrix(i, ColOutgoingOtherStock) = Format(RS!QtyOutOt, gs_formatQty)
        End If
AVG:    'Average Price
        grid.TextMatrix(i, ColLotNo) = Trim(RS!Lot_no)
        If Trim(RS!Cls) = "TAV" Then
         grid.TextMatrix(i, ColCurrent) = Format(RS!QtyEnding, gs_formatAmountIDR)
        ElseIf Trim(RS!Cls) = "TQV" Then
         grid.TextMatrix(i, ColCurrent) = Format(RS!QtyEnding, gs_formatQty)
        ElseIf Trim(RS!Cls) = "AVG" Then
         grid.TextMatrix(i, ColCurrent) = Format(RS!QtyEnding, gs_formatPriceIDR)
        Else
         grid.TextMatrix(i, ColCurrent) = Format(RS!QtyEnding, gs_formatQty)
        End If
       
        RS.MoveNext
    Wend

    If RS.RecordCount >= 3 Then
     LblRecord = Format(RS.RecordCount - 3, "#,##0")
     
     grid.Cell(flexcpBackColor, i - 2, ColDate, i, ColValueInv) = &HFFFFC0
     grid.Cell(flexcpFontBold, i - 2, ColDate, i, ColValueInv) = True
     grid.Cell(flexcpBackColor, i - 2, ColDate, i, ColLotNo) = &HE0E0E0
     grid.Cell(flexcpBackColor, i - 2, ColCurr, i, ColValueInv) = &HE0E0E0
     grid.Cell(flexcpBackColor, i, ColDate, i, ColOutgoingOtherStock) = &HE0E0E0
    Else
     LblRecord = Format("0", "#,##0")
    End If
    
    If RS.State <> adStateClosed Then RS.Close
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdSubMenu_Click()
    DoEvents
    frmMainMenu.Show
    DoEvents
    Unload Me
End Sub

Private Sub DtPeriod_change()
If Format(DtPeriod.Value, "MM") < Format(dateUp, "MM") And Val(Format(DtPeriod.Value, "MM")) = 1 And Val(Format(dateUp, "MM")) = 12 Then _
            DtPeriod.Year = DtPeriod.Year + 1: GoTo pass
    If Format(DtPeriod.Value, "MM") > Format(dateUp, "MM") And Val(Format(DtPeriod.Value, "MM")) = 12 And Val(Format(dateUp, "MM")) = 1 Then _
            DtPeriod.Year = DtPeriod.Year - 1
pass:
    dateUp = Format(DtPeriod.Value, "dd MMM yyyy")
    
Call Header
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

Sub CariProduk()
Dim i As Long
lblPartNumber = ""
lblNm = ""
    With CboProduk
        .clear
        .columnCount = 3
        .ColumnWidths = "90pt;180pt;200pt"
        .ListWidth = 470
        .ListRows = 15
    
    SqlData = " select item_code, makeritem_code, item_name from item_master " & vbCrLf & _
                  " --Where use_endday >= convert(char(8), getdate(), 112) " & vbCrLf & _
                  " Order by Item_Code"
    
'    If sqlfg <> "" And sqlmb <> "" Then
'        sqldata = sqldata + " where " + sqlfg + " and " + sqlmb + " and use_endday >= convert(char(8), getdate(), 112) "
'    Else
'        If sqlfg <> "" Then
'            sqldata = sqldata + " where " + sqlfg + " and use_endday >= convert(char(8), getdate(), 112) "
'        Else
'            If sqlmb <> "" Then
'                sqldata = sqldata + " where " + sqlmb + " and use_endday >= convert(char(8), getdate(), 112) "
'            End If
'        End If
'    End If
    Dim RS As New ADODB.Recordset
    RS.Open SqlData, Db, 1, 3
    i = 0
    While Not RS.EOF
       .AddItem ""
       .List(i, 0) = Trim$(RS!Item_Code & "")
       .List(i, 1) = Trim$(RS!MakerItem_Code & "")
       .List(i, 2) = Trim$(RS!item_name & "")
       RS.MoveNext
       i = i + 1
    Wend
    RS.Close
    End With
End Sub


Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Cancel = True
End Sub

Private Sub TxtProduk_Change()
    Dim sqlitem As String
    Dim RsItem As New ADODB.Recordset
    
    sqlitem = "Select * From Item_Master " & vbCrLf & _
        "   Where Item_Code='" & Trim(Replace(TxtProduk, "'", "''") & "") & "' "
        
    If RsItem.State <> adStateClosed Then RsItem.Close
    
    Set RsItem = Db.Execute(sqlitem)
    
    If Not RsItem.EOF Then
        lblPartNumber.Caption = Trim(RsItem("MakerItem_Code") & "")
        lblNm.Caption = Trim(RsItem("Item_Name") & "")
        cboCls = Trim(RsItem("FinishGoodPart_Cls") & "")
        CboMb = Trim(RsItem("MakeBuy_Cls") & "")
    Else
        lblPartNumber.Caption = ""
        lblNm.Caption = ""
        cboCls = "01"
        CboMb = strAll
    End If
    
End Sub
