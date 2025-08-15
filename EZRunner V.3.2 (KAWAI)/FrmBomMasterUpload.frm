VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmBomMasterUpload 
   BackColor       =   &H00FDDFE3&
   Caption         =   "BOM Master (Upload)"
   ClientHeight    =   10875
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15120
   Icon            =   "FrmBomMasterUpload.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10875
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CDExcel 
      Left            =   9960
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FDDFE3&
      Height          =   780
      Left            =   600
      TabIndex        =   11
      Top             =   2040
      Width           =   13905
      Begin VB.TextBox txtUpload 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   315
         TabIndex        =   12
         Top             =   270
         Width           =   7935
      End
      Begin MSForms.CommandButton cmdTemplate 
         Height          =   375
         Left            =   10560
         TabIndex        =   15
         Top             =   240
         Width           =   1215
         BackColor       =   8454143
         Caption         =   "Template"
         Size            =   "2143;661"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdUpload 
         Height          =   375
         Left            =   9720
         TabIndex        =   14
         Top             =   240
         Width           =   495
         BackColor       =   14737632
         Caption         =   "..."
         Size            =   "873;661"
         FontName        =   "Verdana"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.Label Label1 
         Height          =   255
         Left            =   8520
         TabIndex        =   13
         Top             =   315
         Width           =   1215
         BackColor       =   16637923
         Caption         =   "Upload File"
         Size            =   "2143;450"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   780
      Left            =   600
      TabIndex        =   8
      Top             =   1080
      Width           =   13905
      Begin VB.CommandButton cmdBrowser 
         Caption         =   "..."
         Height          =   300
         Index           =   0
         Left            =   3360
         TabIndex        =   21
         Top             =   300
         Width           =   300
      End
      Begin VB.TextBox LblParent 
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
         Left            =   3765
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   330
         Width           =   4485
      End
      Begin MSForms.ComboBox cboParent 
         Height          =   315
         Left            =   1665
         TabIndex        =   1
         Top             =   300
         Width           =   1500
         VariousPropertyBits=   612386843
         MaxLength       =   10
         DisplayStyle    =   3
         Size            =   "2646;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   3765
         X2              =   8295
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label LblCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Parent Code"
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
         Left            =   315
         TabIndex        =   10
         Top             =   315
         Width           =   1065
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   675
      Left            =   671
      TabIndex        =   6
      Top             =   9330
      Width           =   13905
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
         Left            =   120
         TabIndex        =   7
         Top             =   210
         Width           =   13650
      End
   End
   Begin VB.CommandButton CmdClear 
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
      Left            =   12236
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   10185
      Width           =   1140
   End
   Begin VB.CommandButton CmdSubmit 
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
      Left            =   13436
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   10185
      Width           =   1140
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
      Left            =   671
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   10185
      Width           =   1140
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   420
      Left            =   12705
      TabIndex        =   2
      Top             =   120
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   741
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   4305
      Left            =   600
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   3000
      Width           =   13905
      _cx             =   24527
      _cy             =   7594
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
      AllowUserResizing=   3
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
   Begin MSForms.Label Label5 
      Height          =   255
      Left            =   1320
      TabIndex        =   19
      Top             =   8880
      Width           =   2655
      BackColor       =   16637923
      Caption         =   "Invalid Format End Date"
      Size            =   "4683;450"
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label4 
      Height          =   255
      Left            =   1320
      TabIndex        =   18
      Top             =   8400
      Width           =   2655
      BackColor       =   16637923
      Caption         =   "Invalid Format Start Date"
      Size            =   "4683;450"
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label3 
      Height          =   255
      Left            =   1320
      TabIndex        =   17
      Top             =   7920
      Width           =   2655
      BackColor       =   16637923
      Caption         =   "Invalid Qty Format"
      Size            =   "4683;450"
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Left            =   1320
      TabIndex        =   16
      Top             =   7500
      Width           =   2655
      BackColor       =   16637923
      Caption         =   "Invalid Product Code"
      Size            =   "4683;450"
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00C0C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C000&
      FillColor       =   &H00C0C000&
      Height          =   300
      Left            =   671
      Top             =   8870
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080FFFF&
      FillColor       =   &H0080FFFF&
      Height          =   300
      Left            =   675
      Top             =   8355
      Width           =   300
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      Height          =   300
      Left            =   675
      Top             =   7875
      Width           =   300
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080FF80&
      FillColor       =   &H0080FF80&
      Height          =   300
      Left            =   675
      Top             =   7440
      Width           =   300
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "BOM Master (Upload)"
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
      Left            =   555
      TabIndex        =   0
      Top             =   360
      Width           =   13905
   End
End
Attribute VB_Name = "FrmBomMasterUpload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i As Integer
Dim simpan As Boolean, ubah As Boolean, hapus As Boolean 'Status Ubah/Hapus
Dim parentItem As String 'Parent Item
Dim gridTglAwal As String, gridTglAkhir As String 'Klik Grid
Dim sama As Boolean 'cek Parent
Dim nilKosong As Boolean
Dim kondisi As String
Dim tglAwal As String, tglAkhir As String
Dim tglSesdh As String, tglSeblm As String
Dim l_prod_code As Double, l_qty_format As Double, l_Start_Date As Double, l_End_Date As Double

Dim bteColProdCode As Byte
Dim bteColPartNo As Byte
Dim bteColDesc As Byte
Dim bteColQtyR As Byte
Dim bteColQty As Byte
Dim bteColUnit As Byte
Dim bteColUnitDesc As Byte
Dim bteColDateStart As Byte
Dim bteColDateEnd As Byte
Dim bteColRevision As Byte
Dim bteColInvalidMsg As Byte

Dim ls_PathExcel As String

'====================================================================================================================================================================
' 1. Functions (Start)
'====================================================================================================================================================================

Private Function uf_ValidateInput() As Boolean
    If cboParent = "" Then
        cboParent.SetFocus
        LblErrMsg = "Please Select Parent Code !"
        uf_ValidateInput = False
        Exit Function
    End If
    uf_ValidateInput = True
End Function

Private Sub up_GridHeader()

    Dim i As Integer
    
    bteColProdCode = 0
    bteColPartNo = 1
    bteColDesc = 2
    bteColQtyR = 3
    bteColQty = 4
    bteColUnit = 5
    bteColUnitDesc = 6
    bteColDateStart = 7
    bteColDateEnd = 8
    bteColRevision = 9
    bteColInvalidMsg = 10
    
    With grid
        .clear
        .ColS = 11
        .Rows = 1
        
        .TextMatrix(0, bteColProdCode) = "Product Code"
        .TextMatrix(0, bteColPartNo) = "Part Number"
        .TextMatrix(0, bteColDesc) = "Description"
        .TextMatrix(0, bteColQtyR) = "R Qty"
        .TextMatrix(0, bteColQty) = "Qty"
        .TextMatrix(0, bteColUnit) = "Unit"
        .TextMatrix(0, bteColUnitDesc) = "Desc"
        .TextMatrix(0, bteColDateStart) = "Start Date"
        .TextMatrix(0, bteColDateEnd) = "End Date"
        .TextMatrix(0, bteColRevision) = "Revision"
        .TextMatrix(0, bteColInvalidMsg) = "Invalid Message"
        
'        .ColWidth(bteColSelect) = 300 'besar kolom
        .ColWidth(bteColProdCode) = 2200
        .ColWidth(bteColPartNo) = 2200
        .ColWidth(bteColDesc) = 3000
        .ColWidth(bteColQtyR) = 1300
        .ColWidth(bteColQty) = 1300
        .ColWidth(bteColUnit) = 500
        .ColWidth(bteColUnitDesc) = 600
        .ColWidth(bteColDateStart) = 1500
        .ColWidth(bteColDateEnd) = 1500
        .ColWidth(bteColRevision) = 1500
        .ColWidth(bteColInvalidMsg) = 4500
        
        .ColHidden(bteColQtyR) = True
        
'        .ColAlignment(bteColSelect) = flexAlignCenterCenter
        .ColAlignment(bteColProdCode) = flexAlignLeftCenter
        .ColAlignment(bteColPartNo) = flexAlignLeftCenter
        .ColAlignment(bteColDesc) = flexAlignLeftCenter
        .ColAlignment(bteColQtyR) = flexAlignRightCenter
        .ColAlignment(bteColQty) = flexAlignRightCenter
        .ColAlignment(bteColUnit) = flexAlignCenterCenter
        .ColAlignment(bteColUnitDesc) = flexAlignLeftCenter
        .ColAlignment(bteColDateStart) = flexAlignCenterCenter
        .ColAlignment(bteColDateEnd) = flexAlignCenterCenter
        .ColAlignment(bteColRevision) = flexAlignLeftCenter
        .ColAlignment(bteColInvalidMsg) = flexAlignLeftCenter
        
        .EditMaxLength = 1
    End With

End Sub

Private Sub up_FillCombo()

    Dim rscbo As New ADODB.Recordset 'Isi Combo

    '******** Isi Combo *********
   sql = "select Item_Code as a,MakerItem_Code b,Item_Name c,Unit_Cls " & _
           "from Item_Master " & _
           "where use_endday >= convert(char(8), getdate(), 112) " & _
           "order by Item_Code"
    Set rscbo = Db.Execute(sql)
    
    '**** Isi combo Parent dan Item Code
    cboParent.clear
    cboParent.columnCount = 3
    cboParent.TextColumn = 1
    
    
    i = 0
    Do While Not (rscbo.EOF)
        cboParent.AddItem ""
        cboParent.List(i, 0) = Trim(rscbo("a"))
        cboParent.List(i, 1) = Trim(rscbo("b"))
        cboParent.List(i, 2) = Trim(rscbo("c"))
        
        
        i = i + 1
        rscbo.MoveNext
    Loop
    cboParent.ListWidth = 480
    cboParent.ColumnWidths = "120 pt;120 pt;240 pt"
    cboParent.ListRows = 15
    
    
    '************************
    Set rscbo = Nothing
    
End Sub

Private Sub cboParent_Change()
Dim rsUnit As New ADODB.Recordset

    If cboParent.MatchFound Then
        LblParent = cboParent.Column(2)
        LblErrMsg = ""
    Else
        LblParent = ""
    End If
    
    If LblParent = "" Then
        rsUnit.Open "select * from Item_master where item_code='" & cboParent & "'", Db, adOpenKeyset, adLockOptimistic
        If rsUnit.EOF = False Then
            LblParent = IIf(IsNull(Trim(rsUnit("item_name"))), "", Trim(rsUnit("item_name")))
        End If
        If rsUnit.State <> adStateClosed Then rsUnit.Close
    End If
    
End Sub

Private Sub cmdBrowser_Click(Index As Integer)
 Me.MousePointer = vbHourglass
   frm_BrowseItem.getItemCode = cboParent.Text
   frm_BrowseItem.Show 1
   cboParent.Text = frm_BrowseItem.getItemCode
 Me.MousePointer = vbDefault
End Sub

Private Sub cmdClear_Click()
    cboParent.Text = ""
    txtUpload.Text = ""
    LblErrMsg.Caption = ""
    up_GridHeader
End Sub

Private Sub CmdSubMenu_Click()
    DoEvents
    frmMainMenu.Show
    DoEvents
    Unload Me
End Sub

Private Sub CmdSubmit_Click()
    LblErrMsg.Caption = ""
    Dim ErrMsg As String
    Dim X As Double
    
    Me.MousePointer = vbHourglass


    For X = 1 To grid.Rows - 1
    Dim ProductCode As String, InvalidMsg As String
        If grid.Cell(flexcpBackColor, X, bteColProdCode, X, bteColInvalidMsg) = vbRed Then
            LblErrMsg.Caption = "Cannot save the data !"
            Exit Sub
        End If
    'Me.MousePointer = vbDefault
    Next X
    
    If up_Validasi = True Then
        up_SaveData
    Else
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    
    'Me.MousePointer = vbDefault
End Sub

Private Sub up_SaveData()
Dim i As Integer
    Dim sql As String
    Dim lnAffected As Long
    Dim totSaved As Long
    Dim sBlockTy, pe As String
    Dim ls_errMsg As String
    Dim RS As New ADODB.Recordset
    Dim rsUpdate As New ADODB.Recordset
    
    Db.BeginTrans
    
    Me.MousePointer = vbHourglass
    
    For i = 1 To grid.Rows - 1
    
        If grid.TextMatrix(i, bteColProdCode) <> "" Then
        
        sql = "Select Start_Date from BOM_Master where Parent_ItemCode='" & cboParent.Text & "' and Item_Code = '" & grid.TextMatrix(i, bteColProdCode) & "'"
                            RS.Open sql, Db, adOpenDynamic, adLockReadOnly
                            If Not RS.EOF Then 'jika telah ada data
                            RS.Find "Start_Date = '" & Format(grid.TextMatrix(i, bteColDateStart), "YYYYMMDD") & "'"
                                If RS.EOF Then  'jika telah ada data dan start date sama
                                    If MsgBox("Item already exists. Do you want to continue ?", vbQuestion + vbOKCancel) <> vbOK Then
                                        LblErrMsg = DisplayMsg(8097)
                                        GoTo err
                                    End If
                                End If
                            End If
                            RS.Close
                            
                            '***** cek data max sebelonnya dan update end Date nya lebih kecil 1 hari dr baru yg diinput

                            sql = " update  BOM_Master " & _
                                    " set End_Date ='" & Format(DateAdd("d", -1, grid.TextMatrix(i, bteColDateStart)), "YYYYMMDD") & "', " & _
                                    " Last_Update = getdate(), Last_User = '" & userLogin & "' " & _
                                    " where Parent_ItemCode='" & cboParent.Text & "' and Item_Code = '" & grid.TextMatrix(i, bteColProdCode) & "' and Start_Date = (" & _
                                    " select max(Start_Date) as tglSblm from BOM_Master where Parent_ItemCode='" & cboParent.Text & "' and Item_Code = '" & grid.TextMatrix(i, bteColProdCode) & "' " & _
                                    " and Start_Date  < '" & Format(grid.TextMatrix(i, bteColDateStart), "YYYYMMDD") & "')"
                                Db.Execute sql

                            '****** cek data sesudahnya agar end date yg baru diinput menjadi start date data sesudahnya-1
                            sql = "select min(Start_Date) as tglSsdh from BOM_Master " & _
                                "where parent_ItemCode='" & cboParent.Text & "' and Item_Code = '" & grid.TextMatrix(i, bteColProdCode) & "' and " & _
                                "Start_Date  > '" & grid.TextMatrix(i, bteColDateStart) & "'"
                            Set rsUpdate = Db.Execute(sql)
                            If IsNull(rsUpdate("tglSsdh")) = False Then
                                tglSesdh = CDate(Left(rsUpdate("tglSsdh"), 4) & "-" & Mid(rsUpdate("tglSsdh"), 5, 2) & "-" & Right(rsUpdate("tglSsdh"), 2))
                                tglAkhir = Format(DateAdd("d", -1, tglSesdh), "YYYYMMDD")
                            End If
                            Set rsUpdate = Nothing
        
            With grid
                sql = "INSERT INTO BOM_Master (Parent_ItemCode, Description, Item_Code," & vbCrLf & _
                    " Item_Name,Qty, R_Qty, Unit_Cls, Start_Date,End_Date, " & vbCrLf & _
                    " Last_Update,Last_User,Revision_No) " & _
                    " VALUES (" & _
                    " '" & Trim(cboParent.Value) & "'" & vbCrLf & _
                    ", '" & Trim(LblParent.Text) & "'" & vbCrLf & _
                    ", '" & Trim(.TextMatrix(i, bteColProdCode)) & "'" & vbCrLf & _
                    ", '" & Trim(.TextMatrix(i, bteColDesc)) & "'" & vbCrLf & _
                    ", '" & Trim(.TextMatrix(i, bteColQty)) & "'" & vbCrLf & _
                    ", '0'" & vbCrLf & _
                    ", '" & Trim(.TextMatrix(i, bteColUnit)) & "'" & vbCrLf & _
                    ", " & Format(.TextMatrix(i, bteColDateStart), "yyyyMMdd") & vbCrLf & _
                    ", " & Replace(.TextMatrix(i, bteColDateEnd), "-", "") & vbCrLf & _
                    ", GETDATE(), '" & Trim(userLogin) & "'" & _
                    ", '" & Trim(.TextMatrix(i, bteColRevision)) & "')"

            End With
            
            ls_errMsg = uf_executeSQL(Db, sql)

        End If
        
    Next
    
    Db.CommitTrans
    
    LblErrMsg.Caption = DisplayMsg(1000)
    
    Me.MousePointer = vbDefault
    Exit Sub
    
err:
    Db.RollbackTrans
    Me.MousePointer = vbDefault
    
End Sub

Private Sub cmdtemplate_Click()
    
    Dim objExcel As New Excel.application
    Dim RS As New ADODB.Recordset
    Dim strSQL As String
    Dim i As Double

'On Error GoTo errHandler

    If G_CekExcelApp = False Then LblErrMsg = DisplayMsg(5000): Exit Sub

    LblErrMsg.Caption = ""
    CDExcel.filter = "Excel Files (*.xls)|*.xls"
    CDExcel.filename = "BOM Master (Upload) "
    CDExcel.CancelError = True

    On Error GoTo errCancel
    CDExcel.ShowSave

   On Error GoTo errHandler
    If Len(CDExcel.filename) = 0 Then Exit Sub
    If Dir(CDExcel.filename) <> "" Then
        If MsgBox("Overwrite existing file?", vbExclamation + vbYesNo, "Overwrite") = vbNo Then Exit Sub
    End If
    ls_PathExcel = Mid(CDExcel.filename, 1, Len(CDExcel.filename) - Len(CDExcel.FileTitle))

    MousePointer = MousePointerConstants.vbHourglass

    Set objExcel = New Excel.application
    With objExcel
        .Workbooks.Add
        .Visible = True
        .Cells.Select
        .Cells.EntireColumn.delete

        .Range("A1").EntireColumn.delete xlDown
        .Range("A1:E1").Borders.Weight = xlThin
        .Range("A2:E2").Borders.Weight = xlThin
        .Rows("1:" & grid.Rows).Select
        .Selection.Interior.Pattern = xlNone
        
        .Range("A1:E1").Select
        .Selection.Font.Bold = True
        
        .Range("A2:E2").Select
        .Selection.Font.color = &H80FF80
        
        .Range("A1").Select
        .Selection.MergeCells = True
        .Selection.horizontalAlignment = xlCenter
        .Selection.verticalAlignment = xlCenter
        .Range("A1").Value = "Product Code"


        .Range("B1").Select
        .Selection.MergeCells = True
        .Selection.horizontalAlignment = xlCenter
        .Selection.verticalAlignment = xlCenter
        .Range("B1").Value = "Qty"


        .Range("C1").Select
        .Selection.MergeCells = True
        .Selection.horizontalAlignment = xlCenter
        .Selection.verticalAlignment = xlCenter
        .Range("C1").Value = "Start Date"


        .Range("D1").Select
        .Selection.MergeCells = True
        .Selection.horizontalAlignment = xlCenter
        .Selection.verticalAlignment = xlCenter
        .Range("D1").Value = "End Date"
         
        .Range("E1").Select
        .Selection.MergeCells = True
        .Selection.horizontalAlignment = xlCenter
        .Selection.verticalAlignment = xlCenter
        .Range("E1").Value = "Revision"
        
        .Range("A2").Value = "Char (25)"
        .Range("B2").Value = "Numeric (18, 2)"
        .Range("C2").Value = "Date (MM-dd-yyyy)"
        .Range("D2").Value = "Date (MM-dd-yyyy)"
        .Range("E2").Value = "Char (50)"
        
        .Cells.Select
        .Cells.EntireColumn.AutoFit

        .ActiveWorkbook.SaveAs filename:= _
        CDExcel.filename, FileFormat:= _
        xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False _
        , CreateBackup:=False

    End With

    MousePointer = MousePointerConstants.vbDefault
    Exit Sub

errHandler:
    If err.number <> 0 Then
        MousePointer = MousePointerConstants.vbDefault
        LblErrMsg = DisplayMsg(5000)
        grid.FixedRows = 1
    End If
    If RS.State = adStateOpen Then
        RS.Close
        Set RS = Nothing
    End If
errCancel:
    
End Sub

Private Sub cmdUpload_Click()
Dim objExcel As New Excel.application
Dim objWorkSheet As New Worksheet
Dim objWorkBook As Workbook
Dim i As Long
Dim iCol As Integer
Dim colcount As Integer
Dim RS As New ADODB.Recordset
Dim strSQL As String
Dim filename As String
Dim rsTglAwal As String
Dim rs_DB As New ADODB.Recordset
Dim rs_Unit As New ADODB.Recordset
Dim ls_invalidMsg As String
Dim iGrdRow As Double
Dim rsUpdate As New ADODB.Recordset 'utk Update Data

If uf_ValidateInput = False Then Exit Sub
    
'    cboParent.Value = ""
    Label2 = "Invalid Product Code (0)"
    Label3 = "Invalid Qty Format (0)"
    Label4 = "Invalid Format Start Date (0)"
    Label5 = "Invalid Format End Date (0)"
    
    l_prod_code = 0
    l_qty_format = 0
    l_Start_Date = 0
    l_End_Date = 0
    cmdSubmit.Enabled = True
    
    CDExcel.filename = ""
    LblErrMsg.Caption = ""
    txtUpload = ""
    up_GridHeader
    
    
    
    If G_CekExcelApp = False Then LblErrMsg.Caption = "Excel Application is not found": Exit Sub
    
    Me.MousePointer = vbHourglass
    
    LblErrMsg.Caption = ""
    CDExcel.filter = "Excel Worksheets (*.xls)|*.xls|"
    
    On Error GoTo errCancel
    CDExcel.CancelError = True
    
    On Error GoTo err
    
    CDExcel.ShowOpen
    filename = CDExcel.filename

    txtUpload = filename
    txtUpload.SetFocus
    
    If CDExcel.filename <> "" Then
        'up_GridHeader
        
    'strSQL = " select MakerItem_Code, Item_Name, Unit_Cls  from Item_Master "
            
        Set objExcel = New Excel.application
        Set objWorkBook = objExcel.Workbooks.Open(CDExcel.filename)
        Set objWorkSheet = objWorkBook.Sheets("Sheet1")
        objExcel.Visible = False
        
       ' Set rs_DB = New ADODB.Recordset
        'rs_DB.Open strSQL, Db, adOpenDynamic, adLockOptimistic
        
'        load data sub division
'        strSQL = " Select * from Unit_Cls where Unit_Cls='" & Trim(rs_DB!Unit_cls) & "'"
'
'        Set rs_Unit = New ADODB.Recordset
'        rs_Unit.Open strSQL, Db, adOpenDynamic, adLockOptimistic
        
        i = 3
        iGrdRow = 1
        colcount = 22
        With objWorkSheet
            Do While .Cells(i, 1).Value <> ""
                    ls_invalidMsg = ""
                    
                                           
                        grid.AddItem ""
                        'rs_DB.filter = "MakerItem_Code='" & .Cells(i, 1) & "'"
                        rs_DB.Open "select MakerItem_Code, Item_Name, Unit_Cls from Item_Master where Item_Code='" & Trim(.Cells(i, 1)) & "'", Db, adOpenKeyset, adLockOptimistic
                        
                        rs_Unit.Open "select Unit_Cls, Description from Unit_Cls where Unit_Cls='" & Trim(rs_DB!Unit_cls) & "'", Db, adOpenKeyset, adLockOptimistic
                        
                       If rs_DB.EOF = True Then
                            grid.TextMatrix(iGrdRow, bteColInvalidMsg) = "Product Code is not Exists in Item Master"
                            grid.Cell(flexcpBackColor, iGrdRow, bteColProdCode, iGrdRow, bteColInvalidMsg) = vbRed
                            cmdSubmit.Enabled = False
                            l_prod_code = CDbl(l_prod_code) + 1
                            Label2 = "Invalid Product Code (" & l_prod_code & ")"
                        Else
                            grid.TextMatrix(iGrdRow, bteColPartNo) = Trim(rs_DB!MakerItem_Code)
                            grid.TextMatrix(iGrdRow, bteColDesc) = Trim(rs_DB!item_name)
                            grid.TextMatrix(iGrdRow, bteColUnit) = Trim(rs_Unit!Unit_cls)
                            grid.TextMatrix(iGrdRow, bteColUnitDesc) = Trim(rs_Unit!Description)
                        End If

                        grid.TextMatrix(iGrdRow, bteColProdCode) = .Cells(i, 1)
                        grid.TextMatrix(iGrdRow, bteColQty) = .Cells(i, 2)
                        grid.TextMatrix(iGrdRow, bteColDateStart) = Format(.Cells(i, 3), "dd MMM yyyy")
                        grid.TextMatrix(iGrdRow, bteColDateEnd) = .Cells(i, 4)
                        grid.TextMatrix(iGrdRow, bteColRevision) = .Cells(i, 5)
                        
                        rs_DB.Close
                        rs_Unit.Close
                        
                        If IsNumeric(.Cells(i, 2)) = False Then
                            grid.TextMatrix(iGrdRow, bteColInvalidMsg) = "Invalid Qty Format!!"
                            grid.Cell(flexcpBackColor, iGrdRow, bteColProdCode, iGrdRow, bteColInvalidMsg) = vbRed
                            l_qty_format = CDbl(l_prod_code) + 1
                            Label3 = "Invalid Qty Format (" & l_qty_format & ")"
                            cmdSubmit.Enabled = False
                        Else
                            sql = "Select Start_Date from BOM_Master where Parent_ItemCode='" & cboParent.Text & "' and Item_Code = '" & (.Cells(i, 1)) & "'"
                            RS.Open sql, Db, adOpenDynamic, adLockReadOnly
                            If Not RS.EOF Then 'jika telah ada data
                            RS.Find "Start_Date = '" & Format(.Cells(i, 3), "yyyyMMdd") & "'"
                                If Not RS.EOF Then 'jika telah ada data dan start date sama
                                    grid.TextMatrix(iGrdRow, bteColInvalidMsg) = "Product Code with this Start Date already Exist !"
                                    cmdSubmit.Enabled = False
                                    grid.Cell(flexcpBackColor, iGrdRow, bteColProdCode, iGrdRow, bteColInvalidMsg) = vbRed
                                    l_Start_Date = CDbl(l_Start_Date) + 1
                                    Label4 = "Invalid Start Date Format (" & l_Start_Date & ")"
                                End If
                            End If
                            RS.Close
                            
                        End If
                        
                        iGrdRow = iGrdRow + 1
                        
                    DoEvents
                    i = i + 1
                    
            Loop
            
        End With
        
        objWorkBook.Close
        Set objWorkSheet = Nothing
        Set objWorkBook = Nothing
        Set objExcel = Nothing
        
    LblErrMsg.Caption = "Reading Excel finish"
        
    Me.MousePointer = vbDefault
        
    End If
    Exit Sub
    
errCancel:
err:
    LblErrMsg.Caption = err.Description
    objExcel.Workbooks.Close
    Set objWorkSheet = Nothing
    Set objWorkBook = Nothing
    Set objExcel = Nothing
    Me.MousePointer = vbDefault
    
End Sub

Private Sub Form_Load()
    If gb_Simulation = True Then Call up_InitSimulation(Me)
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
    
    Label2 = "Invalid Product Code (0)"
    Label3 = "Invalid Qty Format (0)"
    Label4 = "Invalid Format Start Date (0)"
    Label5 = "Invalid Format End Date (0)"
    CDExcel.filename = ""
    LblErrMsg.Caption = ""
    txtUpload = ""
    
    up_FillCombo
    up_GridHeader
End Sub

Private Function up_Validasi() As Boolean
    
    If hakUpdate(Me.Name) = 0 Then _
        LblErrMsg = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Function
    
    up_Validasi = True
    If cboParent.MatchFound = False Then
        up_Validasi = False
        LblErrMsg = DisplayMsg(1007)
        cboParent.SetFocus
        Exit Function
    ElseIf txtUpload.Text = "" Then
        up_Validasi = False
        LblErrMsg.Caption = "Please Choose file !"
        txtUpload.SetFocus
        Exit Function
        
    End If
    Me.MousePointer = vbDefault
    
End Function


Private Function uf_executeSQL(ByVal Db As ADODB.Connection, ByVal sql As String) As String
    Dim lnAffected As Long
    On Error GoTo err
    
    Db.Execute sql, lnAffected
    
    uf_executeSQL = ""
    Exit Function
err:
    uf_executeSQL = err.Description
End Function

