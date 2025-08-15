VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm_BrowseItem_WIP 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Browse Item Code"
   ClientHeight    =   9570
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8775
   FillColor       =   &H00FF00FF&
   Icon            =   "frm_BrowseItem_WIP.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   93715.18
   ScaleMode       =   0  'User
   ScaleWidth      =   32527.4
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H0080FFFF&
      Caption         =   "&OK"
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
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9000
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Cancel"
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
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9000
      Width           =   1155
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   120
      TabIndex        =   5
      Top             =   8280
      Width           =   8535
      Begin VB.Label Label4 
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
         Left            =   0
         TabIndex        =   7
         Top             =   600
         Width           =   8430
      End
      Begin MSForms.Label LblErrMsg 
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   230
         Width           =   8175
         ForeColor       =   16711935
         BackColor       =   16637923
         Size            =   "14420;450"
         FontName        =   "Verdana"
         FontEffects     =   1073741827
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
   End
   Begin VB.TextBox txtBrowse 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#,###"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   315
      Left            =   3075
      MaxLength       =   30
      TabIndex        =   1
      Top             =   600
      Width           =   5580
   End
   Begin VSFlex8Ctl.VSFlexGrid gridSearch 
      Height          =   7050
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   8535
      _cx             =   15055
      _cy             =   12435
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
      GridColor       =   8421504
      GridColorFixed  =   8421504
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   275
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frm_BrowseItem_WIP.frx":0E42
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
      Editable        =   0
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
   Begin MSForms.ComboBox cbxChoose 
      Height          =   315
      Left            =   1200
      TabIndex        =   3
      Top             =   600
      Width           =   1740
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "3069;556"
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
      Caption         =   "Search By"
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
      Left            =   120
      TabIndex        =   2
      Top             =   630
      Width           =   885
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Browse Item"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8535
   End
End
Attribute VB_Name = "frm_BrowseItem_WIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public getItemName As String
Public getItemCode As String
Public getParent As String

Public Sub isiCbx()
    With cbxChoose
    For i = 0 To gridSearch.ColS - 1
        .AddItem Trim(gridSearch.TextMatrix(0, i))
    Next
    End With
    cbxChoose.ListIndex = 1
End Sub

Public Sub IsiGrid()
    Dim RS As New ADODB.Recordset
    Dim query As String
        
    query = " select item_code, makeritem_code, item_name, unit_cls, (select description from unit_cls uc where uc.unit_cls=item_master.unit_cls)  Unit_Desc, WH_Code " & _
            " from item_master where FinishGoodPart_Cls = '02' and Item_Code not in (select Item_Code from BOM_Master where Parent_ItemCode='" & getParent & "')" & _
            " and use_endday > convert(char(8), getdate(), 112) order by item_name"
    
    If RS.State = adStateOpen Then RS.Close
    RS.Open query, Db, adOpenForwardOnly, adLockReadOnly

    gridSearch.clear
    Header
    
    gridSearch.ColS = 2
    gridSearch.Editable = flexEDNone

    i = 0
    While Not RS.EOF
        With gridSearch
        i = i + 1
        .AddItem ""
        .TextMatrix(i, 0) = Trim(RS!Item_Code & "")
        .TextMatrix(i, 1) = Trim(RS!item_name & "")
        RS.MoveNext
        End With
    Wend

'    gridSearch.Select 1, 0
'    gridSearch.Sort = flexSortGenericAscending
End Sub

Public Sub search()

On Error GoTo ErrMsg
    
    R = 1
    With gridSearch
        If txtBrowse.Text = "" Then
            .Row = R
            .TopRow = R
            .SetFocus
        Else
            For R = 1 To .Rows - 1
                If cbxChoose.Text = "" Then
                    s = Trim(.TextMatrix(R, 0))
                    cbxChoose.ListIndex = 0
                Else
                    s = Trim(.TextMatrix(R, cbxChoose.ListIndex))
                End If
                
                If UCase(Left(s, Len(Trim(txtBrowse.Text)))) = UCase(txtBrowse.Text) Then
                    .Row = R
                    Exit For
                End If
            Next
            .SetFocus
        End If
        .TopRow = .Row
    End With
    txtBrowse.SetFocus
    
    Exit Sub
ErrMsg:
LblErrMsg = err.Description

End Sub

Public Sub Header()
    With gridSearch
        .TextMatrix(0, 0) = "Item Code"
        .TextMatrix(0, 1) = "Description"
        .ColWidth(0) = 2050
        .ColWidth(1) = 5000
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter

    End With
End Sub

Private Sub cbxChoose_Change()
    If cbxChoose.ListIndex = 0 Then
        gridSearch.Select 1, 0
        gridSearch.Sort = flexSortGenericAscending
    ElseIf cbxChoose.ListIndex = 1 Then
        gridSearch.Select 1, 1
        gridSearch.Sort = flexSortGenericAscending
    ElseIf cbxChoose.ListIndex = 2 Then
        gridSearch.Select 1, 2
        gridSearch.Sort = flexSortGenericAscending
    End If
End Sub

Private Sub cmdBack_Click()
    getItemCode = CStr(gridSearch.TextMatrix(gridSearch.RowSel, 0))
    getItemName = CStr(gridSearch.TextMatrix(gridSearch.RowSel, 1))
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Public Sub awal()
    txtBrowse.Text = ""
    txtBrowse.TabIndex = 0
    gridSearch.TabIndex = 1
    cmdBack.TabIndex = 2
    cmdCancel.TabIndex = 3
    'LblErrMsg.Alignment = 2
End Sub

Private Sub Form_Load()
    IsiGrid
    isiCbx
    awal
    Me.MousePointer = vbDefault
End Sub

Private Sub gridSearch_Click()
    message
End Sub

Private Sub gridSearch_DblClick()
    cmdBack_Click
End Sub

Private Sub gridSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdBack_Click
End Sub

Private Sub gridSearch_RowColChange()
    LblErrMsg.Caption = "Item Name = " & gridSearch.TextMatrix(gridSearch.RowSel, 1)
End Sub

Private Sub txtBrowse_Change()
    search
    message
End Sub

Public Sub message()
    LblErrMsg.Caption = "Item Name = " & gridSearch.TextMatrix(gridSearch.RowSel, 1)
End Sub

Public Function validasi() As Boolean
    If txtBrowse.Text = "" Then
        validasi = False
    ElseIf txtBrowse.Text <> 0 Then
        validasi = True
    End If
End Function

Private Sub txtBrowse_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then gridSearch.SetFocus
End Sub


