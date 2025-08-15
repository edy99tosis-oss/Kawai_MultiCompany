VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form FrmInterface_InvoiceSerial 
   BackColor       =   &H00FDDFE3&
   Caption         =   "I/F SAP - Invoice Serial No"
   ClientHeight    =   10680
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15120
   Icon            =   "FrmInterfaceSAP_InvoiceSerial.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10680
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
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
      TabIndex        =   16
      Tag             =   "FFTT*/"
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
      TabIndex        =   15
      Tag             =   "TFFT*/"
      Top             =   10020
      Width           =   1125
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
      TabIndex        =   14
      Tag             =   "FFTT*/"
      Top             =   10020
      Width           =   1125
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   600
      Left            =   270
      TabIndex        =   11
      Tag             =   "TFTT*/"
      Top             =   9300
      Width           =   14640
      Begin MSComctlLib.ProgressBar PBar 
         Height          =   375
         Left            =   90
         TabIndex        =   12
         Tag             =   "TTTF*/"
         Top             =   180
         Visible         =   0   'False
         Width           =   14460
         _ExtentX        =   25506
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
         Height          =   240
         Left            =   90
         TabIndex        =   13
         Tag             =   "TTTF*/"
         Top             =   180
         Width           =   14325
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1875
      Left            =   300
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
      Begin MSComCtl2.DTPicker DelFrom 
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         Tag             =   "TTFF*/"
         Top             =   840
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
         Format          =   141230083
         CurrentDate     =   37798
      End
      Begin MSComCtl2.DTPicker DelTo 
         Height          =   315
         Left            =   3780
         TabIndex        =   3
         Tag             =   "TTFF*/"
         Top             =   840
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
         Format          =   141230083
         CurrentDate     =   37798
      End
      Begin MSForms.ComboBox cboCust 
         Height          =   345
         Left            =   1680
         TabIndex        =   10
         Tag             =   "TTFF*/"
         Top             =   300
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
         Caption         =   "Customer Code"
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
         Width           =   1350
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
         Top             =   360
         Width           =   60
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delivery From"
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
         Width           =   1215
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
         Top             =   900
         Width           =   210
      End
      Begin VB.Line Line1 
         X1              =   3390
         X2              =   7650
         Y1              =   630
         Y2              =   630
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
   Begin VB.PictureBox Anchor1 
      Height          =   480
      Left            =   840
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   21
      Top             =   240
      Visible         =   0   'False
      Width           =   1200
   End
   Begin InetCtlsObjects.Inet Inetftp 
      Left            =   2400
      Top             =   270
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   5910
      Left            =   285
      TabIndex        =   17
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
      Left            =   13080
      TabIndex        =   18
      Tag             =   "FTTF*/"
      Top             =   360
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   714
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "I/F SAP - Invoice Serial No"
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
      TabIndex        =   20
      Tag             =   "TTTF*/"
      Top             =   390
      Width           =   14610
   End
   Begin MSForms.Label LblRecord 
      Height          =   255
      Left            =   11580
      TabIndex        =   19
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
Attribute VB_Name = "FrmInterface_InvoiceSerial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private m_GettingDir As Boolean
Dim LocalPath As String

Dim sql As String
Dim RS As New Recordset
Dim ubah As Boolean, hapus As Boolean, gavalid As Boolean, ubahedate As Boolean
Dim SDate, EDate, sdateawal, edateakhir
Dim i As Long

Dim StrFilePath As String

Dim bteColSelect As Byte
Dim bteColMaterial As Byte
Dim bteColInvoiceNo As Byte
Dim bteColDeliveryDate As Byte
Dim bteColSerialNo As Byte
Dim bteColSourceSystem As Byte
Dim bteColStorageLoc As Byte
Dim bteColPlaceofDes As Byte
Dim bteColFinalPlaceofDes As Byte
Dim bteColPartsClass As Byte
Dim bteColQty As Byte
Dim bteColAmount As Byte
Dim bteColCurrency As Byte

Sub Header()
    
    Dim X As Integer
    
    LblErrMsg = ""
    LblRecord = "0 Record(s)"
    
    bteColSelect = 0
    bteColMaterial = 1
    bteColInvoiceNo = 2
    bteColDeliveryDate = 3
    bteColSerialNo = 4
    bteColSourceSystem = 5
    bteColStorageLoc = 6
    bteColPlaceofDes = 7
    bteColFinalPlaceofDes = 8
    bteColPartsClass = 9
    bteColQty = 10
    bteColAmount = 11
    bteColCurrency = 12
    
    With grid
        .clear
        
        .Rows = 1
        .ColS = 13
        
        .Cell(flexcpChecked, 0, bteColSelect) = flexUnchecked
        .TextMatrix(0, bteColMaterial) = "Material"
        .TextMatrix(0, bteColInvoiceNo) = "Invoice No"
        .TextMatrix(0, bteColDeliveryDate) = "Delivery Date"
        .TextMatrix(0, bteColSerialNo) = "Serial Number"
        .TextMatrix(0, bteColSourceSystem) = "Source System"
        .TextMatrix(0, bteColStorageLoc) = "Storage Location"
        .TextMatrix(0, bteColPlaceofDes) = "Place of Destination"
        .TextMatrix(0, bteColFinalPlaceofDes) = "Final Place of Destination"
        .TextMatrix(0, bteColPartsClass) = "Parts Classification"
        .TextMatrix(0, bteColQty) = "Qty"
        .TextMatrix(0, bteColAmount) = "Amount"
        .TextMatrix(0, bteColCurrency) = "Currency"
        
        For X = 0 To 12
            .ColAlignment(X) = flexAlignCenterCenter
        Next
        
        .ColWidth(bteColSelect) = 300
        .ColWidth(bteColMaterial) = 1500
        .ColWidth(bteColInvoiceNo) = 1500
        .ColWidth(bteColDeliveryDate) = 1500
        .ColWidth(bteColSerialNo) = 1500
        .ColWidth(bteColSourceSystem) = 1500
        .ColWidth(bteColStorageLoc) = 1600
        .ColWidth(bteColPlaceofDes) = 1800
        .ColWidth(bteColFinalPlaceofDes) = 2100
        .ColWidth(bteColPartsClass) = 1800
        .ColWidth(bteColQty) = 1000
        .ColWidth(bteColAmount) = 1500
        .ColWidth(bteColCurrency) = 1000
        
        .EditMaxLength = 1
        
    End With
End Sub

Sub gridLoad()
    Dim RsSearch As New ADODB.Recordset
    Dim StrSearch As String
    Dim CData As Integer
    Dim XData As Integer
    Dim JmlTran As Integer

    LblErrMsg = ""
    
    On Error GoTo ErrSearch
    
    
    If cboCust.MatchFound = False Then
        LblErrMsg.Caption = DisplayMsg("4072")
        Exit Sub
    End If
    Me.MousePointer = vbHourglass
    Call Header
    


StrSearch = "  SELECT " & vbCrLf & _
                        "   Material=Coalesce(D.SAP_Item_Code,B.ITEM_Code), " & vbCrLf & _
                        "   Invoice_No=b.Invoice_No, " & vbCrLf & _
                        "   DO_Date=H.Packing_Date, " & vbCrLf & _
                        "   C.Serial_No, " & vbCrLf & _
                        "   'D68681' SourceSystem, " & vbCrLf & _
                        "   'BP0A' StorageLocation, " & vbCrLf & _
                        "    ISNULL(Final_Destination_CLS,'')Final_Destination, " & vbCrLf & _
                        "    ISNULL(K.PlaceOfDestination_Cls,'1')FinalPlace_Destination,  " & vbCrLf & _
                        "    NoCommercial_Cls, " & vbCrLf & _
                        "    CASE WHEN C.Serial_No <> '' THEN 1 ELSE B.Qty END Qty, "

StrSearch = StrSearch + "    ROUND(B.Price * (CASE WHEN C.Serial_No <> '' THEN 1 ELSE B.Qty END),2) Amount, " & vbCrLf & _
                        "    'USD' Currency,      " & vbCrLf & _
                        "    ISNULL(E.InterfaceDel_Cls,'0') InterfaceDel_Cls,ISNULL(E.Cust_Code,'') " & vbCrLf & _
                        "  " & vbCrLf & _
                        "  FROM  Packing_Detail A " & vbCrLf & _
                        "  LEFT JOIN Invoice_Detail B ON A.Packing_No=B.Packing_No AND A.PackingSeq_No=B.PackingSeq_No " & vbCrLf & _
                        "  LEFT JOIN Item_Master D ON B.Item_Code=D.Item_Code  " & vbCrLf & _
                        "  lEFT JOIN Invoice_Master E ON E.Invoice_No=B.Invoice_No " & vbCrLf & _
                        "  LEFT JOIN Delivery_Order F ON F.PO_No=A.Order_No AND F.Seq_No=A.Order_SeqNo " & vbCrLf & _
                        "  LEFT JOIN Packing_Master H ON A.Packing_No=H.Packing_No " & vbCrLf & _
                        "  LEFT JOIN OrderEntry_Detail K ON A.Order_No=K.PO_No AND A.Item_Code=K.Item_Code AND A.Order_SeqNo=K.Seq_No AND A.SerialNoFrom=K.SerialNoFrom AND A.SerialNoTo=K.SerialNoto " & vbCrLf & _
                        "  Outer Apply (Select * From Delivery_Order where PO_No=A.Order_No AND Seq_No=A.Order_SeqNo and Item_Code=a.Item_Code)G " & vbCrLf
StrSearch = StrSearch + "  Outer Apply (Select * From Serial_Detail where Item_Code=a.Item_Code and Serial_No between A.SerialNoFrom and A.SerialNoTo)C " & vbCrLf & _
                        "  Outer Apply (Select NoCommercial_Cls=Case when NoCommercial_Cls='0' then '1' else '2' end  From OrderEntry_Master where PO_No=A.Order_No)J " & vbCrLf & _
                        "  WHERE B.Currency_Code='02' and H.Packing_Date >= '" & Format(DelFrom.Value, "yyyy-MM-dd") & "' AND H.Packing_Date <= '" & Format(DelTo.Value, "yyyy-MM-dd") & "'  " & vbCrLf
                        
    If cbointeface.ListIndex = 1 Then
        StrSearch = StrSearch + "AND ISNULL(InterfaceDel_Cls,'') = '' " & vbCrLf
    Else
        StrSearch = StrSearch + "AND ISNULL(InterfaceDel_Cls,'') = '1' " & vbCrLf
    End If
    
    If cboCust.ListIndex > 0 Then
        StrSearch = StrSearch + "AND ISNULL(E.Cust_Code,'') = '" & Trim(cboCust.Text) & "' " & vbCrLf
    End If

                        
    StrSearch = StrSearch + "  ORDER BY H.Packing_Date,Invoice_No,Coalesce(D.SAP_Item_Code,B.ITEM_Code)  "


    
    If RsSearch.State <> adStateClosed Then RsSearch.Close
    
    Set RsSearch = Db.Execute(StrSearch)
    
    If RsSearch.EOF Then
        LblErrMsg = DisplayMsg("0013")
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    
    XData = 0
    JmlTran = 0
    i = 1
    
    Do While Not RsSearch.EOF
        grid.AddItem ""
'        If RsSearch.Fields(2) = "X" Then
            If RsSearch("InterfaceDel_Cls") = "1" Then
                grid.Cell(flexcpChecked, grid.Rows - 1, 0) = Checked
            Else
                grid.Cell(flexcpChecked, grid.Rows - 1, 0) = flexUnchecked
            End If
            
            JmlTran = JmlTran + 1
'        End If
        
        For XData = 0 To RsSearch.Fields.Count - 3
            If XData = 10 Then
                grid.TextMatrix(grid.Rows - 1, XData + 1) = Format(RsSearch.Fields(XData), "#,##0.00")
            ElseIf XData = 2 Then
                grid.TextMatrix(grid.Rows - 1, XData + 1) = Format(RsSearch.Fields(XData), "yyyy/MM/dd")
            Else
                grid.TextMatrix(grid.Rows - 1, XData + 1) = Trim(RsSearch.Fields(XData)) & ""
            End If
        Next XData
        
        grid.Cell(flexcpAlignment, i, bteColMaterial, i, bteColCurrency) = flexAlignLeftCenter
        grid.Cell(flexcpAlignment, i, bteColQty, i, bteColAmount) = flexAlignRightCenter
        
        i = i + 1
        RsSearch.MoveNext
    Loop
     
    LblRecord = Format(JmlTran, "#,##0") & " Record(s)"
    
    Me.MousePointer = vbDefault
    Exit Sub
    
ErrSearch:

    LblErrMsg = "[" & err.number & "]-" & err.Description
    err.clear
    Me.MousePointer = vbDefault
End Sub

Function fc_WriteIniFile(strSectionHeader As String, strVariableName As String, strValue As String, strFileName As String) As Integer
    fc_WriteIniFile = WritePrivateProfileString(strSectionHeader, strVariableName, strValue, strFileName)
End Function

Private Sub CboCust_Change()
    Call cboCust_Click
End Sub

Private Sub cboCust_Click()
    If cboCust.ListIndex < 0 Then
        LblCustomer = ""
    Else
        LblCustomer = cboCust.Column(1)
    End If
    Call Header
End Sub

Private Sub cmd_clear_Click()
    Kosong
    Header
End Sub

Private Sub Cmd_Save_Click()

'On Error GoTo ErrHandler

    Dim fHndl As Integer
    Dim PlaceDestination As String
    Dim FinalDestination As String
    Dim InvoiceNo As String
    Dim PartClass As String
    Dim OldInvoiceNo As String
    Dim Y As Double
    Dim Q As Double
    Dim XData As Integer
    Dim PbMax As Integer
    Dim strSQL As String
    Dim headerCsv As Boolean
    Dim thn As String
    
    LblErrMsg = ""
    StrFilePath = ""
    
    If grid.Rows <= 1 Then
        LblErrMsg = DisplayMsg("0013")
        Exit Sub
    End If
    
    XData = 1
    headerCsv = True
    
    PbMax = grid.Rows - 1
    PBar.Visible = True
    PBar.Max = PbMax
    
    Dim dbUpd As New Connection
    dbUpd.ConnectionString = Db.ConnectionString
    dbUpd.Open
    dbUpd.BeginTrans
        
    Do While XData <= grid.Rows - 1
        If grid.Cell(flexcpChecked, XData, 0) = flexChecked Then
        
            If headerCsv = True Then
                
                StrFilePath = App.path & "\IFData" & "\KI3" & Format(Now, "yyyyMMddhhmmss") & ".csv"
                fHndl = FreeFile
                
                Open StrFilePath For Output As fHndl
                              
                headerCsv = False
                
                strSQL = "Update Invoice_Master" & vbCrLf & _
                         "Set InterfaceDel_Cls='1',InterfaceDel_Date=Getdate(),InterfaceDel_User='" & userLogin & "'" & vbCrLf & _
                         "Where Invoice_No='" & Trim(InvoiceNo) & "'"
                Db.Execute strSQL
                
            End If
                        
            InvoiceNo = grid.TextMatrix(XData, bteColInvoiceNo)
            
            InvoiceNo = Trim(grid.TextMatrix(XData, bteColInvoiceNo))
            PlaceDestination = IIf(Trim(grid.TextMatrix(XData, bteColPlaceofDes)) = "JAPAN", "1", "2")
            FinalDestination = IIf(Trim(grid.TextMatrix(XData, bteColFinalPlaceofDes)) = "JAPAN", "1", "2")
            
            If Trim(grid.TextMatrix(XData, bteColSerialNo)) <> "" Then
                PartClass = ""
            ElseIf Trim(grid.TextMatrix(XData, bteColSerialNo)) = "" Then
                PartClass = Trim(grid.TextMatrix(XData, bteColPartsClass))
            End If
            
            Print #fHndl, Trim(grid.TextMatrix(XData, bteColMaterial)) & "," & _
                          InvoiceNo & "," & _
                          Format(grid.TextMatrix(XData, bteColDeliveryDate), "yyyyMMdd") & "," & _
                          Trim(grid.TextMatrix(XData, bteColSerialNo)) & "," & _
                          Trim(grid.TextMatrix(XData, bteColSourceSystem)) & "," & _
                          Trim(grid.TextMatrix(XData, bteColStorageLoc)) & "," & _
                          Trim(PlaceDestination) & "," & _
                          Trim(FinalDestination) & "," & _
                          PartClass & "," & _
                          Trim(grid.TextMatrix(XData, bteColQty)) & "," & _
                          Format(grid.TextMatrix(XData, bteColAmount), "###0.00") & "," & _
                          Trim(grid.TextMatrix(XData, bteColCurrency))
        
'            Next Y
        End If
        PBar.Value = XData
        XData = XData + 1
        
'        OldInvoiceNo = Grid.TextMatrix(XData, bteColInvoiceNo)
    Loop
    
    Header
    gridLoad
    
    If headerCsv = False Then
'        Close #fHndl
'
'        UploadFTP
'
'        dbUpd.CommitTrans
'
'        LblErrMsg = "Export Invoice Serial Data Success !"


  Close #fHndl
        Sleep 1000
        
        Shell App.path & "\FTPuploadDL.bat", vbHide
        Sleep 10000
        
        FileCopy StrFilePath, App.path & "\IFDATA\BACKUP\" & Dir(StrFilePath)
        Sleep 1000
        
        Kill StrFilePath
        Sleep 1000
        LblErrMsg = "Export Invoice Serial Data Success !"
    Else
        LblErrMsg = "Please check data that you want to export !"
    End If
    
    PBar.Visible = False
    
    Exit Sub

errHandler:
    LblErrMsg = "[" & err.number & "]-" & err.Description
    dbUpd.RollbackTrans
    PBar.Visible = False
    err.clear
End Sub

Private Sub cmdSearch_Click()
    gridLoad
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
    If ErrMsg = "" Then
        Unload Me
    Else
        LblErrMsg.Caption = ErrMsg
    End If
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim i As Double
Dim Q As Double
If Col = 0 And Row = 0 Then

    For Q = 1 To grid.Rows - 1
        If grid.Cell(flexcpChecked, 0, 0) = flexChecked Then
           grid.Cell(flexcpChecked, Q, 0) = flexChecked
        ElseIf grid.Cell(flexcpChecked, 0, 0) = flexUnchecked Then
            grid.Cell(flexcpChecked, Q, 0) = flexUnchecked
        End If
    Next Q
    
End If

    Dim Invoice_No As String
    Invoice_No = Trim(grid.TextMatrix(Row, bteColInvoiceNo))
    
    For i = 1 To grid.Rows - 1
        
        If Invoice_No = Trim(grid.TextMatrix(i, bteColInvoiceNo)) Then
            If i <> Row Then
                If grid.Cell(flexcpChecked, i, 0) = flexUnchecked Then
                    grid.Cell(flexcpChecked, i, 0) = flexChecked
                Else
                    grid.Cell(flexcpChecked, i, 0) = flexUnchecked
                End If
            End If
            Invoice_No = Trim(grid.TextMatrix(i, bteColInvoiceNo))
        End If
        
    Next i
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

Private Sub Command1_Click()
    MsgBox "satu" & vbTab & "dua"
End Sub

Private Sub Form_Load()
    If gb_Simulation = True Then Call up_InitSimulation(Me)
    Call Kosong
    Call Header
    
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
    
    With Anchor1
     ' .RegString = "AnchorCtrl,Positions," & Me.Name & "0|0"
     ' .DoInit
    End With
End Sub

Sub Kosong()
    Dim RsCust As New ADODB.Recordset
    Dim strSQL As String
    Dim X As Integer
    
    strSQL = "Select Trade_Code Cust_Code, Trade_Name Cust_Name" & vbCrLf & _
                  " From Trade_Master " & vbCrLf & _
                  "        Where Trade_Cls=2 AND left(trade_Code,1)='C' " & vbCrLf & _
                  "            ORDER BY Trade_Code " & vbCrLf
                
    If RsCust.State <> adStateClosed Then RsCust.Close
    
    Set RsCust = Db.Execute(strSQL)
    
    cboCust.clear
    cboCust.ListWidth = 350
    cboCust.columnCount = 2
    cboCust.ColumnWidths = "100 pt;250 pt"
    
    cboCust.AddItem ""
    cboCust.List(0, 0) = "ALL"
    cboCust.List(0, 1) = "ALL"
    
    X = 1
    Do While Not RsCust.EOF
        cboCust.AddItem ""
        cboCust.List(X, 0) = Trim(RsCust("Cust_Code") & "")
        cboCust.List(X, 1) = Trim(RsCust("Cust_Name") & "")
        RsCust.MoveNext
        X = X + 1
    Loop
    cboCust.ListIndex = 0
    
    
    DelFrom = Format(Now(), "yyyy-MMM-") & "01"
    DelTo = DateAdd("m", 1, DelFrom) - 1
    
    
    With cbointeface
        .clear
        .AddItem "Yes"
        .AddItem "No"
        
        .ListIndex = 1
    End With

End Sub

Private Sub UploadFTP()
Dim host_name As String
Dim RsFtp As New ADODB.Recordset
Dim strSQL As String
Dim Host As String
Dim user As String
Dim Pwd As String
Dim Folder As String

    Enabled = False
    MousePointer = vbHourglass
    
    LblErrMsg.Caption = "Working"
    
    strSQL = "Select * From ftp_Setting"
    If RsFtp.State <> adStateClosed Then RS.Close
    
    Set RsFtp = Db.Execute(strSQL)
    
    If Not RsFtp.EOF Then
    
        Host = IIf(IsNull(RsFtp("Host3")), "", Trim(RsFtp("host3")))
        user = IIf(IsNull(RsFtp("user3")), "", Trim(RsFtp("user3")))
        Pwd = IIf(IsNull(RsFtp("pwd3")), "", Trim(RsFtp("pwd3")))
        Folder = IIf(IsNull(RsFtp("Folder3")), "", Trim(RsFtp("Folder3")))

    End If

    
        
    DoEvents

    ' You must set the URL before the user name and
    ' password. Otherwise the control cannot verify
    ' the user name and password and you get the error:
    '
    '       Unable to connect to remote host
    host_name = Host
    If LCase$(Left$(host_name, 6)) <> "ftp://" Then _
        host_name = "ftp://" & host_name
    Inetftp.URL = host_name

    Inetftp.userName = user
    Inetftp.Password = Pwd
'    Folder = Folder & "520FID02.TXT"
    ' Do not include the host name here. That will make
    ' the control try to use its default user name and
    ' password and you'll get the error again.
'    inetFTP.Execute , "Put " & _
'        txtLocalFile.Text & " " & txtRemoteFile.Text
        
    Inetftp.Execute , "Put " & _
         StrFilePath & " " & Folder

End Sub

Private Sub Cmd_SubMenu_Click()
    Unload Me
    frmMainMenu.Show
End Sub

Private Sub DelFrom_Change()
    Call Header
End Sub


Private Sub DelTo_Change()
    Call Header
End Sub

