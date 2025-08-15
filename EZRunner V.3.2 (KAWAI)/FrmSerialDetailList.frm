VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmSerialDetailList 
   BackColor       =   &H00FDDFE3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Finished Goods Stock Report"
   ClientHeight    =   10305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9930
   Icon            =   "FrmSerialDetailList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10305
   ScaleWidth      =   9930
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   1905
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   9720
      Begin MSComCtl2.DTPicker dtEnd 
         Height          =   315
         Left            =   4200
         TabIndex        =   17
         Top             =   1260
         Width           =   1575
         _ExtentX        =   2778
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
         CurrentDate     =   43377
      End
      Begin MSComCtl2.DTPicker dtStart 
         Height          =   315
         Left            =   1980
         TabIndex        =   16
         Top             =   1260
         Width           =   1605
         _ExtentX        =   2831
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
         CurrentDate     =   43377
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FFFF&
         Caption         =   "Search"
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
         Left            =   7620
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   630
         Width           =   1035
      End
      Begin VB.Label Label3 
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
         Height          =   255
         Left            =   3690
         TabIndex        =   19
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Production Date"
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
         Left            =   300
         TabIndex        =   18
         Top             =   1290
         Width           =   1605
      End
      Begin MSForms.ComboBox cbostatus 
         Height          =   315
         Left            =   1995
         TabIndex        =   11
         Top             =   750
         Width           =   2205
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "3889;556"
         ListRows        =   15
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         Value           =   "AAAAAA"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delivery"
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
         Left            =   300
         TabIndex        =   10
         Top             =   780
         Width           =   720
      End
      Begin VB.Line Line1 
         X1              =   4335
         X2              =   8640
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Label lblNm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Name"
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
         Left            =   4335
         TabIndex        =   9
         Top             =   300
         Width           =   1200
      End
      Begin MSForms.ComboBox cbo 
         Height          =   315
         Left            =   1995
         TabIndex        =   8
         Top             =   240
         Width           =   2205
         VariousPropertyBits=   746604571
         MaxLength       =   25
         DisplayStyle    =   3
         Size            =   "3889;556"
         ListRows        =   15
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         Value           =   "AAAAAA"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
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
         Left            =   300
         TabIndex        =   7
         Top             =   300
         Width           =   1155
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   615
      Left            =   120
      TabIndex        =   14
      Top             =   8580
      Width           =   9675
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
         Height          =   360
         Left            =   120
         TabIndex        =   15
         Top             =   180
         Width           =   9435
      End
   End
   Begin VB.CommandButton CmdExcel 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Excel"
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
      Left            =   7500
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   9570
      Width           =   1035
   End
   Begin VB.CommandButton CmdSubmit 
      BackColor       =   &H0080FFFF&
      Caption         =   "Submit"
      Enabled         =   0   'False
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
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9570
      Width           =   1035
   End
   Begin VB.CommandButton Cmd_Back 
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
      Index           =   8
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9570
      Width           =   1125
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   465
      Left            =   7920
      TabIndex        =   0
      Top             =   120
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   820
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   4365
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3900
      Width           =   9705
      _cx             =   17119
      _cy             =   7699
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
   Begin VB.TextBox TxtSeqNo 
      Height          =   225
      Left            =   180
      TabIndex        =   6
      Top             =   8850
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Finished Goods Stock Report"
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
      Left            =   180
      TabIndex        =   1
      Top             =   900
      Width           =   9570
   End
End
Attribute VB_Name = "FrmSerialDetailList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bteColSelect As Byte
Dim bteColItem As Byte
Dim bteColItemName As Byte
Dim bteColItemQty As Byte
Dim bteColSerial As Byte
Dim bteColResult As Byte
Dim bteColDelivery As Byte
Dim bteColDStatus As Byte
Dim bteColDStatusUpdate As Byte
Dim bteColHDelivery As Byte
Dim bteColHDeliveryUpdate As Byte
Dim bteColHItem As Byte
Dim bteColResultDate As Byte
Dim bteColDate As Byte

Private Sub cbo_Change()
    Call cbo_Click
    Call Header
End Sub

Private Sub cbo_Click()
    If cbo.ListIndex < 0 Then
        lblNm = ""
    Else
        lblNm = cbo.Column(1)
    End If
End Sub

Private Sub cbostatus_Change()
    Call cboStatus_Click
    Call Header
End Sub

Private Sub cboStatus_Click()
    If cboStatus.ListIndex < 0 Then
        cboStatus.ListIndex = 0
    End If
End Sub

Private Sub Cmd_Back_Click(Index As Integer)
    DoEvents
    frmMainMenu.Show
    DoEvents
    Unload Me
End Sub

Sub SetCombo()

    Dim RsItem As New ADODB.Recordset
    Dim sqlMaster As String
    Dim i As Long
    
    sqlMaster = "Select Item_Code, Item_Name" & vbCrLf & _
                      " From Item_Master" & vbCrLf & _
                      "      WHERE COALESCE(FinishGoodPart_Cls,'')='01' " & vbCrLf & _
                      "         Order by Item_Code"
                      
    If RsItem.State <> adStateClosed Then RsItem.Close
    
    Set RsItem = Db.Execute(sqlMaster)
                          
    ' Combo Item Code
    ' -----------------------
    cbo.clear
    cbo.columnCount = 2
    
    cbo.AddItem ""
    cbo.List(0, 0) = strAll
    cbo.List(0, 1) = strAll
    
    i = 1
    Do While Not RsItem.EOF
        cbo.AddItem ""
        cbo.List(i, 0) = Trim(RsItem("Item_Code") & "")
        cbo.List(i, 1) = Trim(RsItem("Item_Name") & "")
        RsItem.MoveNext
        i = i + 1
    Loop
                
    cbo.ListWidth = 300
    cbo.ColumnWidths = "100 pt; 200 pt"
    cbo.ListIndex = 0
                
    ' Combo Delivery Status
    ' ---------------------------
    cboStatus.clear
    cboStatus.columnCount = 1
    
    cboStatus.AddItem ""
    cboStatus.List(0, 0) = "No"
    
    cboStatus.AddItem ""
    cboStatus.List(1, 0) = "Yes"
    
    cboStatus.AddItem ""
    cboStatus.List(2, 0) = strAll
    
    cboStatus.ListIndex = 0
    
End Sub

Private Sub CmdExcel_Click()
    Dim xlapp As New Excel.application

    Dim i As Long
    Dim j As Long
    Dim xstr As String
    
    If grid.Rows > 1 Then
    
            LblErrMsg = ""
            Me.MousePointer = vbHourglass
    
            Dim Idx As Integer
    
            With xlapp
                .Workbooks.Add
                
                .Range("A1:E1").Merge
                .Range("A1", "E1").Columns.Font.Name = "Arial"
                .Range("A1", "E1").Columns.Font.Size = "14"
                .Range("A1") = "Finished Goods Stock Report"
                .Range("A1").horizontalAlignment = xlCenter
                .Range("A1").Font.Bold = True
                           
                
                .Range("A3") = "Product Code :"
                .Range("B3") = " " & Trim(cbo)
                .Range("A3").horizontalAlignment = xlLeft
                .Range("A3").Font.Bold = True
                
                .Range("C3") = Trim(lblNm)
                .Range("C3").horizontalAlignment = xlLeft
                
                .Range("A4") = "Delivery :"
                .Range("B4") = " " & Trim(cboStatus)
                .Range("A4").horizontalAlignment = xlLeft
                .Range("A4").Font.Bold = True
                               
                .Range("A5") = "Start Date :"
                .Range("B5") = " " & Format(DtStart, "dd-MMM-yyyy")
                .Range("A5", "B5").horizontalAlignment = xlLeft
                .Range("A5").Font.Bold = True
                
                .Range("C5") = "End Date :"
                .Range("D5") = " " & Format(DtEnd, "dd-MMM-yyyy")
                .Range("C5", "D5").horizontalAlignment = xlLeft
                .Range("C5").Font.Bold = True
                
                'Header
                
                .Range("A7") = "Product Code"
                .Range("B7") = "Product Name"
                .Range("C7") = "Qty Stock"
                .Range("D7") = "Serial Number"
                
                If cboStatus <> "No" Then
                    .Range("E7") = "Delivery No"
                    .Range("F7") = "Result Date"
                Else
                    .Range("E7") = "Result Date"
                    
                    
                End If
                
                j = 7
                
                For i = 1 To grid.Rows - 1
                    j = j + 1
                    LblErrMsg = " Transfering data ... (record " & i & ")"
                    DoEvents
                    If Trim(grid.TextMatrix(i, bteColItem)) <> "" Then
                        ' Detail Data Summary
                        .Range("A" & Trim(Str(j))) = Trim(grid.TextMatrix(i, bteColItem))
                        .Range("B" & Trim(Str(j))) = Trim(grid.TextMatrix(i, bteColItemName))
                        .Range("C" & Trim(Str(j))) = Trim(grid.TextMatrix(i, bteColItemQty))
'                       i = i + 1
                    End If
                    
                    'Detail serial
                     .Range("D" & Trim(Str(j))) = Trim("'" & grid.TextMatrix(i, bteColSerial))
                     .Range("E" & Trim(Str(j))) = Trim("'" & grid.TextMatrix(i, bteColResultDate))
                     
                     
                    If cboStatus <> "No" Then
                        .Range("E" & Trim(Str(j))) = Trim(grid.TextMatrix(i, bteColHDelivery))
                        .Range("F" & Trim(Str(j))) = Trim(grid.TextMatrix(i, bteColResultDate))
                    End If
                    
                Next i
                
                 LblErrMsg = " Transfering data complete. "
                .Visible = True
                .Columns("A:F").Columns.AutoFit
                .WindowState = xlMaximized
                .ActiveWindow.Zoom = 80
                
                .Range("A7", "F7").Columns.Font.Bold = True
                .Range("A7", "F7").horizontalAlignment = xlCenter
            
                End With
    
            Me.MousePointer = vbDefault
    Else
        LblErrMsg = "No Data to display."
    End If

End Sub

Private Sub CmdSubmit_Click()

    Dim strSQL As String
    Dim X As Long
    
    On Error GoTo ErrSubmit
    
    LblErrMsg = ""
    
    If MsgBox("Are you sure want to update this data ? ", vbYesNo + vbDefaultButton2 + vbQuestion, "Confirmation") = vbNo Then
        Exit Sub
    End If
    
    Me.MousePointer = vbHourglass
    
    For X = 1 To grid.Rows - 1
        If grid.TextMatrix(X, bteColDStatus) <> grid.TextMatrix(X, bteColDStatusUpdate) Then
            
            ' Update Delivery of Serial Number
            strSQL = "Update Serial_Detail" & vbCrLf & _
                        "   Set DO_No='" & Trim(grid.TextMatrix(X, bteColHDeliveryUpdate)) & "' , DO_SeqNo='0' " & vbCrLf & _
                        "       Where Item_Code='" & Trim(grid.TextMatrix(X, bteColHItem)) & "' " & vbCrLf & _
                        "           AND Serial_No='" & Trim(grid.TextMatrix(X, bteColSerial)) & "' " & vbCrLf
                        
            Db.Execute strSQL
        End If
    Next X
    
    Call BrowseGrid
    
    LblErrMsg = DisplayMsg("1000")
    
    Me.MousePointer = vbDefault
    
    Exit Sub

ErrSubmit:
    Me.MousePointer = vbDefault
    LblErrMsg = "[" & err.number & "]-" & err.Description
    err.clear

End Sub

Private Sub Command1_Click()

    Call BrowseGrid
    End Sub

Private Sub dtEnd_Change()
Call Header
End Sub

Private Sub dtStart_Change()
Call Header
End Sub

Private Sub Form_Load()

    Call SetCombo
    Call Header

End Sub

Private Sub Header()

    bteColItem = 0
    bteColItemName = 1
    bteColItemQty = 2
    bteColSerial = 3
    bteColResult = 4
    bteColDelivery = 5
    bteColResultDate = 6
    bteColDStatus = 7
    bteColDStatusUpdate = 8
    bteColHDelivery = 9
    bteColHDeliveryUpdate = 10
    bteColHItem = 11
    
    
    With grid
        .ColS = 12
        .Rows = 1
        
        .TextMatrix(0, bteColItem) = "Product Code"
        .TextMatrix(0, bteColItemName) = "Product Name"
        .TextMatrix(0, bteColItemQty) = "Stock"
        .TextMatrix(0, bteColSerial) = "Serial Number"
        .TextMatrix(0, bteColResult) = "Result"
        .TextMatrix(0, bteColDelivery) = "Delivery"
        .TextMatrix(0, bteColResultDate) = "Result Date"
    
         '.ColWidth(bteColselect) = 300
         .ColWidth(bteColItem) = 1750
         .ColWidth(bteColItemName) = 3250
         .ColWidth(bteColItemQty) = 1250
         .ColWidth(bteColSerial) = 1500
         .ColWidth(bteColResult) = 750
         .ColWidth(bteColDelivery) = 1000
         .ColWidth(bteColResultDate) = 1400
    
         .ColAlignment(bteColItem) = flexAlignLeftCenter
         .ColAlignment(bteColItemName) = flexAlignLeftCenter
         .ColAlignment(bteColItemQty) = flexAlignCenterCenter
         .ColAlignment(bteColSerial) = flexAlignCenterCenter
         .ColAlignment(bteColResult) = flexAlignCenterCenter
         .ColAlignment(bteColDelivery) = flexAlignCenterCenter
         .ColAlignment(bteColResultDate) = flexAlignLeftCenter
              
        .ColHidden(bteColResult) = True
        .ColHidden(bteColDStatus) = True
        .ColHidden(bteColDStatusUpdate) = True
        .ColHidden(bteColHDelivery) = True
        .ColHidden(bteColHDeliveryUpdate) = True
        .ColHidden(bteColHItem) = True
        
    End With
End Sub

Private Sub BrowseGrid()
    Dim i As Long
    Dim strSQL As String
    Dim RsDet As New ADODB.Recordset
  
  
    LblErrMsg = ""
    Me.MousePointer = vbHourglass
    On Error GoTo ErrBrowse
    
    Call Header
                          
'    strSQL = " SELECT * FROM   " & vbCrLf & _
'                  "   (   " & vbCrLf & _
'                  "       SELECT SD.Item_Code, Item_Name, Serial_No, '' SerialTo, 0 Jml,   " & vbCrLf & _
'                  "                Rtrim(COALESCE(Result_No,'')) Result_SeqNo,    " & vbCrLf & _
'                  "                RTrim(COALESCE(SD.DO_No,'')) Delivery_No,   " & vbCrLf & _
'                  "                (  " & vbCrLf & _
'                  "                  select  Receipt_Date from Part_Receipt where  Receipt_Cls='p1' and ProductionResult_Cls='1'   " & vbCrLf & _
'                  "                          and DailySeq_No=SD.Product_No and Serial_No between SerialNoFrom and SerialNoTo                               " & vbCrLf & _
'                  "                 )Result_date,  " & vbCrLf & _
'                  "               0 Header , SD.product_No      " & vbCrLf & _
'                  "        FROM Serial_Detail SD               LEFT JOIN Item_Master IM ON IM.Item_Code=SD.Item_Code    "
'
'strSQL = strSQL + "        WHERE COALESCE(SD.Result_No,'')<>''   " & vbCrLf & _
'                  "    " & vbCrLf & _
'                  "               AND SD.Item_Code='" & Trim(cbo) & "' " & vbCrLf & _
'                  "      " & vbCrLf & _
'                  "       UNION ALL   " & vbCrLf & _
'                  "      " & vbCrLf & _
'                  "       SELECT SD.Item_Code, Item_Name,    " & vbCrLf & _
'                  "                Min(Serial_No) SerialFrom,  MAX(Serial_No) SerialTo,  COUNT(Serial_No) Jml,   " & vbCrLf & _
'                  "                '' Result_SeqNo, '' Delivery_No,  " & vbCrLf & _
'                  "                (                  select  Receipt_Date from Part_Receipt where  Receipt_Cls='p1' and ProductionResult_Cls='1'   " & vbCrLf & _
'                  "                          and DailySeq_No=SD.Product_No and Serial_No between SerialNoFrom and SerialNoTo                               "
'
'strSQL = strSQL + "                 ) Result_date,  " & vbCrLf & _
'                  "                 1 Header, SD.Product_No  " & vbCrLf & _
'                  "        FROM Serial_Detail SD     " & vbCrLf & _
'                  "            LEFT JOIN Item_Master IM ON IM.Item_Code=SD.Item_Code    " & vbCrLf & _
'                  "        WHERE COALESCE(SD.Result_No,'')<>''   " & vbCrLf & _
'                  "    " & vbCrLf & _
'                  "               AND SD.Item_Code='" & Trim(cbo) & "' " & vbCrLf & _
'                  "        GROUP BY SD.Item_Code, Item_Name , Serial_No, SD.Product_No " & vbCrLf & _
'                  "   ) DetailList     ORDER BY Item_Code, Header desc, Serial_No  " & vbCrLf & _
'                  "    " & vbCrLf & _
'                  "  "


strSQL = " SELECT  Item_Code,Item_Name,Serial_No,SerialTo,Jml,Result_SeqNo,Result_date,Header,Delivery_No=Coalesce((Select MAX(DO_No) From Delivery_Order where item_Code=Item_Code and  Serial_No>=SerialNoFrom  AND Serial_No<=SerialNoto),'') FROM    " & vbCrLf & _
                  "    (    " & vbCrLf & _
                  "        SELECT SD.Item_Code, Item_Name, Serial_No, '' SerialTo, 0 Jml,    " & vbCrLf & _
                  "                 Rtrim(COALESCE(Result_No,'')) Result_SeqNo,     " & vbCrLf & _
                  "                 RTrim(COALESCE(SD.DO_No,'')) Delivery_No,    " & vbCrLf & _
                  "                 (   select distinct NULLIF(MAX(COALESCE(Receipt_Date,'9999-12-31')),'9999-12-31') from Part_Receipt where  Receipt_Cls='p1' and ProductionResult_Cls='1'    " & vbCrLf & _
                  "                           and DailySeq_No=(CAST(SD.Product_No AS INT)) and Serial_No between SerialNoFrom and SerialNoTo                                " & vbCrLf & _
                  "                  )Result_date,   " & vbCrLf & _
                  "                0 Header     " & vbCrLf & _
                  "         FROM Serial_Detail SD               " & vbCrLf & _
                  "             LEFT JOIN Item_Master IM ON IM.Item_Code=SD.Item_Code             "

strSQL = strSQL + "             WHERE COALESCE(SD.Result_No,'')<>''    " & vbCrLf & _
                  "      " & vbCrLf & _
                  IIf(cbo.Text = "ALL", "", "                AND SD.Item_Code='" & Trim(cbo) & "' ") & vbCrLf & _
                  "        " & vbCrLf & _
                  "        UNION ALL    " & vbCrLf & _
                  "  " & vbCrLf & _
                  "         " & vbCrLf & _
                  "     Select Item_Code,Item_Name,serialFrom=min(serialFrom),SerialTo=max(serialTo),Jml=count(Jml),Result_SeqNo,Delivery_No,Result_Date=Max(Result_Date),Header From    " & vbCrLf & _
                  "     (      " & vbCrLf & _
                  "        SELECT SD.Item_Code, Item_Name,     " & vbCrLf & _
                  "                 Min(Serial_No) SerialFrom,  MAX(Serial_No) SerialTo,  COUNT(Serial_No) Jml,    "

strSQL = strSQL + "                 '' Result_SeqNo, '' Delivery_No,   " & vbCrLf & _
                  "                   (   select distinct NULLIF(MAX(COALESCE(Receipt_Date,'9999-12-31')),'9999-12-31') from Part_Receipt where  Receipt_Cls='p1' and ProductionResult_Cls='1'    " & vbCrLf & _
                  "                           and DailySeq_No=(CAST(SD.Product_No AS INT)) and Serial_No between SerialNoFrom and SerialNoTo                                " & vbCrLf & _
                  "                  ) Result_date,   " & vbCrLf & _
                  "                  1 Header " & vbCrLf & _
                  "         FROM Serial_Detail SD      " & vbCrLf & _
                  "             LEFT JOIN Item_Master IM ON IM.Item_Code=SD.Item_Code     " & vbCrLf & _
                  "         WHERE COALESCE(SD.Result_No,'')<>''    " & vbCrLf & _
                  "      " & vbCrLf & _
                  IIf(cbo.Text = "ALL", "", "                AND SD.Item_Code='" & Trim(cbo) & "' ") & vbCrLf & _
                  "        GROUP BY SD.Item_Code, Item_Name,sd.Product_No,Serial_No  "

strSQL = strSQL + "      )A where Result_date>='" & Format(DtStart.Value, "yyyy-MM-dd") & "' " & vbCrLf & _
                  "          and Result_date<='" & Format(DtEnd.Value, "yyyy-MM-dd") & "' " & vbCrLf & _
                  "       group by Item_Code,Item_Name,Result_SeqNo,Delivery_No,Header " & vbCrLf & _
                  "  " & vbCrLf & _
                  "  " & vbCrLf & _
                  "    ) DetailList   " & vbCrLf & _
                  "         where Result_date>='" & Format(DtStart.Value, "yyyy-MM-dd") & "'  " & vbCrLf & _
                  "         and Result_date<='" & Format(DtEnd.Value, "yyyy-MM-dd") & "'  " & vbCrLf & _
                  "     " & vbCrLf & _
                  "        ORDER BY Item_Code, Header desc, Serial_No "


    Set RsDet = Db.Execute(strSQL)
        
    If RsDet.EOF Then
        LblErrMsg = DisplayMsg("0013")
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    
    i = 1
    With grid
        Do While Not RsDet.EOF
            .Rows = .Rows + 1
            
            If RsDet("Header") = 1 Then
                
                ' Header & Summary By Item
                ' ---------------------------------
                .TextMatrix(i, bteColItem) = Trim(RsDet("Item_Code") & "")
                .TextMatrix(i, bteColItemName) = Trim(RsDet("Item_Name") & "")
                .TextMatrix(i, bteColItemQty) = Trim(RsDet("Jml"))
'                .TextMatrix(i, bteColResultDate) = Format(Trim(RsDet!Result_Date), "dd MMM yyyy")
                
            
            Else
                ' Detail Serial No
                ' ------------------------
'                .TextMatrix(i, bteColItem) = Trim(RsDet("Item_Code") & "")
'                .TextMatrix(i, bteColItemName) = Trim(RsDet("Item_Name") & "")
                .TextMatrix(i, bteColSerial) = Trim(RsDet("Serial_No") & "")
                .TextMatrix(i, bteColHItem) = Trim(RsDet("Item_Code") & "")
                .TextMatrix(i, bteColResultDate) = Format(Trim(RsDet!Result_Date), "dd MMM yyyy")
                
                ' Result Status
                If RsDet.Fields("Result_SeqNo") = "" Then
                   .Cell(flexcpChecked, i, bteColResult) = flexUnchecked
                Else
                   .Cell(flexcpChecked, i, bteColResult) = flexChecked
                End If
                
                ' Delivery Status
                If RsDet.Fields("Delivery_No") = "" Then
                   .Cell(flexcpChecked, i, bteColDelivery) = flexUnchecked
                   .TextMatrix(i, bteColDStatus) = 0
                   .TextMatrix(i, bteColDStatusUpdate) = 0
                Else
                   .Cell(flexcpChecked, i, bteColDelivery) = flexChecked
                   .TextMatrix(i, bteColDStatus) = 1
                   .TextMatrix(i, bteColDStatusUpdate) = 1
                End If
                
                .TextMatrix(i, bteColHDelivery) = RsDet.Fields("Delivery_No")
                .TextMatrix(i, bteColHDeliveryUpdate) = RsDet.Fields("Delivery_No")
            
            End If
            
            RsDet.MoveNext
            i = i + 1
        Loop
        
    End With
Me.MousePointer = vbDefault
    Exit Sub
    
ErrBrowse:

    LblErrMsg = "[" & err.number & "]-" & err.Description
    err.clear
    Me.MousePointer = vbDefault
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If grid.Cell(flexcpChecked, Row, Col) = flexUnchecked Then
        grid.TextMatrix(Row, bteColDStatusUpdate) = 0
        grid.TextMatrix(Row, bteColHDeliveryUpdate) = ""
    Else
        grid.TextMatrix(Row, bteColDStatusUpdate) = 1
        grid.TextMatrix(Row, bteColHDeliveryUpdate) = "Manual Adjust"
    End If
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> bteColDelivery Then Cancel = True
End Sub

