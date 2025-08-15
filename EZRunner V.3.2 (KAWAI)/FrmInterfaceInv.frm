VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmInterfaceInv 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Inventory InterfaceInventory Interface"
   ClientHeight    =   10680
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15120
   Icon            =   "FrmInterfaceInv.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10680
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   975
      Left            =   300
      TabIndex        =   8
      Top             =   1260
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
         Height          =   435
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   300
         Width           =   1125
      End
      Begin MSComCtl2.DTPicker InvFrom 
         Height          =   315
         Left            =   2040
         TabIndex        =   12
         Tag             =   "TTFF*/"
         Top             =   360
         Width           =   1350
         _ExtentX        =   2381
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
         Format          =   151453699
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker InvTo 
         Height          =   315
         Left            =   3780
         TabIndex        =   14
         Tag             =   "TTFF*/"
         Top             =   2160
         Visible         =   0   'False
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
         Format          =   151453699
         CurrentDate     =   37798
      End
      Begin VB.Label LblStatus 
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
         Height          =   195
         Left            =   13695
         TabIndex        =   19
         Top             =   645
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Line Line1 
         Visible         =   0   'False
         X1              =   3540
         X2              =   7800
         Y1              =   -480
         Y2              =   -480
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
         TabIndex        =   15
         Top             =   2220
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inventory Period"
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
         Left            =   420
         TabIndex        =   13
         Tag             =   "TTFF*/"
         Top             =   420
         Width           =   1440
      End
      Begin VB.Label LblPlant 
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
         Left            =   3600
         TabIndex        =   11
         Top             =   -720
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plant Location"
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
         TabIndex        =   10
         Top             =   -720
         Visible         =   0   'False
         Width           =   1185
      End
      Begin MSForms.ComboBox cboPlant 
         Height          =   345
         Left            =   1860
         TabIndex        =   9
         Top             =   -840
         Visible         =   0   'False
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   600
      Left            =   270
      TabIndex        =   6
      Top             =   9300
      Width           =   14640
      Begin MSComctlLib.ProgressBar PBar 
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   180
         Visible         =   0   'False
         Width           =   14475
         _ExtentX        =   25532
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
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
         Height          =   285
         Left            =   7215
         TabIndex        =   7
         Top             =   180
         Width           =   75
      End
   End
   Begin VB.CommandButton Cmd_Clear 
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
      Left            =   12570
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   10020
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
      Left            =   270
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   10020
      Width           =   1125
   End
   Begin VB.CommandButton Cmd_save 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Export"
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
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   10020
      Width           =   1125
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   6525
      Left            =   300
      TabIndex        =   3
      Top             =   2340
      Width           =   14640
      _cx             =   25823
      _cy             =   11509
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
      Left            =   13080
      TabIndex        =   5
      Top             =   360
      Width           =   1845
      _extentx        =   3254
      _extenty        =   714
   End
   Begin MSForms.Label LblRecord 
      Height          =   255
      Left            =   11535
      TabIndex        =   18
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Inventory Interface"
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
      Left            =   300
      TabIndex        =   4
      Top             =   360
      Width           =   14610
   End
End
Attribute VB_Name = "FrmInterfaceInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sql As String
Dim RS As New Recordset
Dim ubah As Boolean, hapus As Boolean, gavalid As Boolean, ubahedate As Boolean
Dim SDate, EDate, sdateawal, edateakhir
Dim i As Integer

Dim bteColSelect As Byte
Dim bteColCompany As Byte
Dim bteColCalYear As Byte
Dim bteColCalMonth As Byte
Dim bteColPlant As Byte
Dim bteColLocation As Byte
Dim bteColMaterial As Byte
Dim bteColProfitCenter As Byte
Dim bteColQuantity As Byte
Dim bteColBaseUOM As Byte
Dim bteColAmount As Byte
Dim bteColCurrency As Byte
Dim bteColSourceSystem As Byte

Sub Header()

    Dim X As Integer
    
    LblErrMsg = ""
    LblRecord = "0 Record(s)"
    
    bteColSelect = 0
    bteColCompany = 1
    bteColCalYear = 2
    bteColCalMonth = 3
    bteColPlant = 4
    bteColLocation = 5
    bteColMaterial = 6
    bteColProfitCenter = 7
    bteColQuantity = 8
    bteColBaseUOM = 9
    bteColAmount = 10
    bteColCurrency = 11
    bteColSourceSystem = 12
  
    With grid
        .clear
        
        .Rows = 1
        .ColS = 13
        
        .TextMatrix(0, bteColSelect) = "S"
        .TextMatrix(0, bteColCompany) = "Company"
        .TextMatrix(0, bteColCalYear) = "Calendar Year"
        .TextMatrix(0, bteColCalMonth) = "Calendar Month"
        .TextMatrix(0, bteColPlant) = "Plant"
        .TextMatrix(0, bteColLocation) = "Storage Location"
        .TextMatrix(0, bteColMaterial) = "Material"
        .TextMatrix(0, bteColProfitCenter) = "Profit Center"
        .TextMatrix(0, bteColQuantity) = "Quantity"
        .TextMatrix(0, bteColBaseUOM) = "Base UOM"
        .TextMatrix(0, bteColAmount) = "Amount"
        .TextMatrix(0, bteColCurrency) = "Currency"
        .TextMatrix(0, bteColSourceSystem) = "SourceSystem"
        
        .ColWidth(bteColSelect) = 0
        .ColWidth(bteColCompany) = 1000
        .ColWidth(bteColCalYear) = 1300
        .ColWidth(bteColCalMonth) = 1500
        .ColWidth(bteColPlant) = 850
        .ColWidth(bteColLocation) = 1500
        .ColWidth(bteColMaterial) = 1600
        .ColWidth(bteColProfitCenter) = 1500
        .ColWidth(bteColQuantity) = 1250
        .ColWidth(bteColBaseUOM) = 1000
        .ColWidth(bteColAmount) = 850
        .ColWidth(bteColCurrency) = 850
        .ColWidth(bteColSourceSystem) = 1250
        
        .ColAlignment(bteColSelect) = flexAlignLeftCenter
        .ColAlignment(bteColCompany) = flexAlignCenterCenter
        .ColAlignment(bteColCalYear) = flexAlignCenterCenter
        .ColAlignment(bteColCalMonth) = flexAlignCenterCenter
        .ColAlignment(bteColPlant) = flexAlignCenterCenter
        .ColAlignment(bteColLocation) = flexAlignCenterCenter
        .ColAlignment(bteColMaterial) = flexAlignLeftCenter
        .ColAlignment(bteColProfitCenter) = flexAlignCenterCenter
        .ColAlignment(bteColQuantity) = flexAlignRightCenter
        .ColAlignment(bteColBaseUOM) = flexAlignCenterCenter
        .ColAlignment(bteColAmount) = flexAlignRightCenter
        .ColAlignment(bteColCurrency) = flexAlignCenterCenter
        .ColAlignment(bteColSourceSystem) = flexAlignCenterCenter
        
        .ColHidden(0) = True
        
        .EditMaxLength = 1
    End With

End Sub

Function fc_WriteIniFile(strSectionHeader As String, strVariableName As String, strValue As String, strFileName As String) As Integer
    fc_WriteIniFile = WritePrivateProfileString(strSectionHeader, strVariableName, strValue, strFileName)
End Function

Private Sub cboplant_Change()
    Call cboplant_Click
End Sub

Private Sub cboplant_Click()
    If cboPlant.ListIndex < 0 Then
        LblPlant = ""
    Else
        LblPlant = cboPlant.Column(1)
    End If
    Call Header
End Sub

Private Sub Cmd_Save_Click()
    Dim adoStream As ADODB.Stream
    Dim adoStreamOut As ADODB.Stream
    Dim fs
    Dim a
    
    Dim XData As Integer
    Dim IFPart As String
    Dim ListOfData As String
    Dim PbMax As Integer
    Dim CLoop As Long
    
    On Error GoTo ErrExport
    
    LblErrMsg = ""
    
    IFPart = App.path & "\IFData" & "\IF_Inv_" & Format(InvFrom, "yyyyMM") & "_on_" & Format(InvFrom, "yyyyMMdd")
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(IFPart & ".tsv", True)
    
    XData = 0
    
    If grid.Rows <= 1 Then
        LblErrMsg = DisplayMsg("0013")
        Exit Sub
    End If
    
    PbMax = grid.Rows - 1
    
    PBar.Visible = True
    PBar.Max = PbMax
    
    Do While XData <= grid.Rows - 1
        
        ListOfData = grid.TextMatrix(XData, 1) & vbTab & _
                            grid.TextMatrix(XData, 2) & vbTab & _
                            grid.TextMatrix(XData, 3) & vbTab & _
                            grid.TextMatrix(XData, 4) & vbTab & _
                            grid.TextMatrix(XData, 5) & vbTab & _
                            grid.TextMatrix(XData, 6) & vbTab & _
                            grid.TextMatrix(XData, 7) & vbTab & _
                            grid.TextMatrix(XData, 8) & vbTab & _
                            grid.TextMatrix(XData, 9) & vbTab & _
                            grid.TextMatrix(XData, 10) & vbTab & _
                            grid.TextMatrix(XData, 11) & vbTab & _
                            grid.TextMatrix(XData, 12)
                            
        a.WriteLine (ListOfData)
        
        PBar.Value = XData
        XData = XData + 1
        
    Loop
    
    a.Close
    
    Set adoStream = New ADODB.Stream

    adoStream.Charset = "ASCII"
    adoStream.Open
    adoStream.LoadFromFile IFPart & ".tsv"
    
    adoStream.Position = 0
    Set adoStreamOut = New ADODB.Stream
    adoStreamOut.Charset = "UTF-8"
    adoStreamOut.Open
    adoStreamOut.WriteText adoStream.ReadText
    adoStreamOut.SaveToFile App.path & "\IFData" & "\520SPD03.txt", adSaveCreateOverWrite
    
    PBar.Visible = False
    
    Kill (IFPart & ".tsv")
    
    LblErrMsg = "Export Inventory Data Success !"
    
    Exit Sub

ErrExport:
    LblErrMsg = "[" & err.number & "]-" & err.Description
    err.clear
    
End Sub

Private Sub cmdSearch_Click()

    Dim RsSearch As New ADODB.Recordset
    Dim StrSearch As String
    Dim CData As Integer
    Dim XData As Integer

    LblErrMsg = ""
    
    On Error GoTo ErrSearch
    Me.MousePointer = vbHourglass
    
    Call Header
    
    sql = "SELECT * FROM InterfaceInv_Valuation WHERE "
    
    
    StrSearch = "  DECLARE @StartDate DATETIME  " & vbCrLf & _
                            "  DECLARE @EndDate DATETIME  " & vbCrLf & _
                            "  DECLARE @YearPeriod VARCHAR(4)  " & vbCrLf & _
                            "  DECLARE @MonthPeriod VARCHAR(2)  " & vbCrLf & _
                            "  DECLARE @CMonth Numeric(2)  " & vbCrLf & _
                            "  " & vbCrLf & _
                            " declare @last_closing datetime " & vbCrLf & _
                            " set @last_closing = ( " & vbCrLf & _
                            "   select top 1 cast(cast(inventory_year as varchar) + right('00' + cast(inventory_month as varchar), 2) + '01' as datetime) " & vbCrLf & _
                            "   from inventory_control order by cast(cast(inventory_year as varchar) + right('00' + cast(inventory_month as varchar), 2) + '01' as datetime) desc " & vbCrLf & _
                            "   ) " & vbCrLf & _
                            "    "

    StrSearch = StrSearch + "  SET @StartDate = '" & Format(InvFrom, "yyyy-MM-dd") & "'  " & vbCrLf & _
                            "  " & vbCrLf & _
                            "  SET @YearPeriod = '" & Format(InvFrom, "yyyy") & "' " & vbCrLf & _
                            "  SET @MonthPeriod = '" & Format(InvFrom, "MM") & "' " & vbCrLf & _
                            "  SET @CMonth=@MonthPeriod " & vbCrLf & _
                            "   " & vbCrLf & _
                            "  -- ############################  " & vbCrLf & _
                            "  -- Inventory " & vbCrLf & _
                            "  -- ############################  " & vbCrLf & _
                            "  " & vbCrLf & _
                            " SELECT '' S, 'C520' Company, @YearPeriod [Calendar Year], @MonthPeriod [Calendar Month], " & vbCrLf & _
                            "   'C521' Plant, 'C521' StorageLocation, Item_Code Material, " & vbCrLf
    
    StrSearch = StrSearch + "   '5200000110' ProfitCenter, Stock Quantity, 'PC' BaseUOM, 0 Amount, " & vbCrLf & _
                            "   'USD' Currency, 'KI' SourceSystem " & vbCrLf & _
                            "       FROM " & vbCrLf & _
                            "        " & vbCrLf & _
                            "           (SELECT Item_Code,  " & vbCrLf & _
                            "                   CASE WHEN DATEDIFF(M,@last_closing,@StartDate) < 0 Then  " & vbCrLf & _
                            "                               Coalesce((SELECT  SUM(COALESCE(Inventory,[Current]))  " & vbCrLf & _
                            "                                       FROM Stock_History SH WHERE SH.Item_code=SM.Item_Code " & vbCrLf & _
                            "                                           AND SH.Stock_Year=@YearPeriod AND SH.Stock_Month=@CMonth),0) " & vbCrLf & _
                            "                            WHEN DATEDIFF(M,@last_closing,@StartDate) = 0 Then " & vbCrLf & _
                            "                               SUM(COALESCE(LM_Inventory,LM_Current))  " & vbCrLf
    
    StrSearch = StrSearch + "                            WHEN DATEDIFF(M,@last_closing,@StartDate) = 1 Then " & vbCrLf & _
                            "                               SUM(COALESCE(TM_Inventory,TM_Current))  " & vbCrLf & _
                            "                            Else " & vbCrLf & _
                            "                               SUM(COALESCE(NM_Inventory,NM_Current))  " & vbCrLf & _
                            "                   END Stock                                " & vbCrLf & _
                            "               FROM Stock_Master SM " & vbCrLf & _
                            "                   WHERE Warehouse_Code IN  " & vbCrLf & _
                            "                       (SELECT WH_Code FROM dbo.WareHouse_Master " & vbCrLf & _
                            "                           WHERE StockControl_Cls='01') " & vbCrLf & _
                            "                   GROUP BY Item_Code " & vbCrLf & _
                            "           ) Stk  " & vbCrLf
    
    StrSearch = StrSearch + "        ORDER BY Stk.Item_Code " & vbCrLf & _
                            "  "
    
    If RsSearch.State <> adStateClosed Then RsSearch.Close
    
    Set RsSearch = Db.Execute(StrSearch)
    
    If RsSearch.EOF Then
        LblErrMsg = DisplayMsg("0013")
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    
    XData = 0
    
    Do While Not RsSearch.EOF
        grid.AddItem ""
        For XData = 0 To RsSearch.Fields.Count - 1
            If XData = bteColAmount Then
                grid.TextMatrix(grid.Rows - 1, XData) = Format(RsSearch.Fields(XData), "#,##0.00")
            ElseIf XData = bteColQuantity Then
                grid.TextMatrix(grid.Rows - 1, XData) = Format(RsSearch.Fields(XData), "#,##0.000")
            Else
                grid.TextMatrix(grid.Rows - 1, XData) = Trim(RsSearch.Fields(XData)) & ""
        End If
        Next XData
        RsSearch.MoveNext
    Loop
    
    LblRecord = Format(grid.Rows - 1, "#,##0") & " Record(s)"
    
    Me.MousePointer = vbDefault
    Exit Sub
    
ErrSearch:

    LblErrMsg = "[" & err.number & "]-" & err.Description
    err.clear
    Me.MousePointer = vbDefault
    
End Sub

Private Sub Command1_Click()
    MsgBox "satu" & vbTab & "dua"
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
    Call Kosong
    Call Header
    
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
End Sub

Sub Kosong()
    Dim RsPlant As New ADODB.Recordset
    Dim strSQL As String
    Dim X As Integer
    
    strSQL = "Select Trade_Code Plant_Code, Trade_Name Plant_Name" & vbCrLf & _
                  " From Trade_Master " & vbCrLf & _
                  "        WHERE Trade_Cls=1 " & vbCrLf & _
                  "            ORDER BY Trade_Code " & vbCrLf
                
    If RsPlant.State <> adStateClosed Then RsPlant.Close
    
    Set RsPlant = Db.Execute(strSQL)
    
    cboPlant.clear
    cboPlant.ListWidth = 350
    cboPlant.columnCount = 2
    cboPlant.ColumnWidths = "100 pt;250 pt"
    
    X = 0
    Do While Not RsPlant.EOF
        cboPlant.AddItem ""
        cboPlant.List(X, 0) = Trim(RsPlant("Plant_Code") & "")
        cboPlant.List(X, 1) = Trim(RsPlant("Plant_Name") & "")
        RsPlant.MoveNext
        X = X + 1
    Loop

    InvFrom = Format(Now(), "yyyy-MMM-") & "01"
    InvTo = DateAdd("m", 1, InvFrom) - 1

End Sub

Private Sub Cmd_SubMenu_Click()
    Unload Me
    frmMainMenu.Show
End Sub

Sub Browse()

End Sub

Private Sub cmd_clear_Click()
    Kosong
    Header
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    LblErrMsg.Caption = ErrMsg
End If
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 If grid.Col <> bteColSelect Then Cancel = True
End Sub

Private Sub Grid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
  If grid.Col = bteColSelect Then
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then
      KeyAscii = 0
    End If
    If KeyAscii = Asc(".") Then KeyAscii = 0
  End If
End Sub

Private Sub InvFrom_Change()
    Call Header
End Sub


Private Sub InvTo_Change()
    Call Header
End Sub

