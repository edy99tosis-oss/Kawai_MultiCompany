VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmPriceMol 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Price Mold"
   ClientHeight    =   10980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15120
   Icon            =   "FrmPriceMol.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10980
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtDesc 
      BackColor       =   &H00FDDFE3&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.#0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9360
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   8400
      Width           =   705
   End
   Begin VB.TextBox txtmaxqty 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   13410
      TabIndex        =   3
      Text            =   "XX"
      Top             =   8400
      Width           =   1425
   End
   Begin VB.TextBox txtprice 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   11910
      TabIndex        =   2
      Text            =   "XX"
      Top             =   8400
      Width           =   1425
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   435
      Left            =   13043
      TabIndex        =   15
      Top             =   270
      Width           =   1845
      _extentx        =   3254
      _extenty        =   767
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
      Left            =   12563
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9750
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
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9720
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   600
      Left            =   293
      TabIndex        =   10
      Top             =   9000
      Width           =   14625
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
         TabIndex        =   11
         Top             =   240
         Width           =   14265
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
      Left            =   293
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9750
      Width           =   1140
   End
   Begin VB.CommandButton CmdCancel 
      BackColor       =   &H0080FFFF&
      Caption         =   "Ca&ncel"
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
      Left            =   11363
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9750
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   885
      Left            =   300
      TabIndex        =   9
      Top             =   900
      Width           =   14625
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FDDFE3&
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
         Index           =   2
         Left            =   3510
         TabIndex        =   20
         Top             =   375
         Width           =   960
      End
      Begin VB.Line Line2 
         X1              =   3480
         X2              =   7935
         Y1              =   615
         Y2              =   615
      End
      Begin MSForms.ComboBox cboSupplier 
         Height          =   315
         Left            =   1530
         TabIndex        =   0
         Top             =   315
         Width           =   1815
         VariousPropertyBits=   612386843
         MaxLength       =   15
         DisplayStyle    =   3
         Size            =   "3201;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FDDFE3&
         Caption         =   "Supplier"
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
         TabIndex        =   12
         Top             =   375
         Width           =   705
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid grid 
      Height          =   5610
      Left            =   300
      TabIndex        =   13
      Top             =   1920
      Width           =   14625
      _cx             =   25797
      _cy             =   9895
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
      FocusRect       =   5
      HighLight       =   2
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   3
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   0
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
      Editable        =   1
      ShowComboButton =   1
      WordWrap        =   -1  'True
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
   Begin MSComCtl2.DTPicker dtsdate 
      Height          =   315
      Left            =   10170
      TabIndex        =   21
      Top             =   8400
      Width           =   1620
      _ExtentX        =   2858
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
      CurrentDate     =   37781
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Currency"
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
      Left            =   8400
      TabIndex        =   27
      Top             =   8040
      Width           =   795
   End
   Begin MSForms.ComboBox cbocurr 
      Height          =   315
      Left            =   8400
      TabIndex        =   26
      Top             =   8400
      Width           =   870
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "1535;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblitem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   1
      Left            =   4920
      TabIndex        =   24
      Top             =   8400
      Width           =   3435
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00A6D2FF&
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
      Left            =   4920
      TabIndex        =   23
      Top             =   8010
      Width           =   960
   End
   Begin VB.Label lblitem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   0
      Left            =   2760
      TabIndex        =   22
      Top             =   8400
      Width           =   2130
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00A6D2FF&
      Caption         =   "Max Qty"
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
      Left            =   14160
      TabIndex        =   19
      Top             =   8010
      Width           =   705
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00A6D2FF&
      Caption         =   "Price"
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
      Left            =   12855
      TabIndex        =   18
      Top             =   8010
      Width           =   420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00A6D2FF&
      Caption         =   "Start Date"
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
      Left            =   10170
      TabIndex        =   17
      Top             =   8010
      Width           =   885
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00A6D2FF&
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
      Left            =   2760
      TabIndex        =   16
      Top             =   8010
      Width           =   1080
   End
   Begin MSForms.ComboBox CboItem 
      Height          =   315
      Left            =   420
      TabIndex        =   1
      Top             =   8400
      Width           =   2235
      VariousPropertyBits=   612386843
      DisplayStyle    =   3
      Size            =   "3942;556"
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00A6D2FF&
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
      Left            =   420
      TabIndex        =   14
      Top             =   8010
      Width           =   1155
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00A6D2FF&
      BackStyle       =   1  'Opaque
      Height          =   360
      Index           =   1
      Left            =   300
      Top             =   7950
      Width           =   14655
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080C0FF&
      Height          =   855
      Index           =   0
      Left            =   293
      Top             =   7980
      Width           =   14655
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Price Mold"
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
      Left            =   293
      TabIndex        =   8
      Top             =   270
      Width           =   14580
   End
End
Attribute VB_Name = "FrmPriceMol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Const isiPart = "Finish Good,Parts/wip/material"
Dim ubah As Boolean
Dim TextGrid As String

Dim bteColSelect As Byte
Dim bteColItemCode As Byte
Dim bteColPartNumber As Byte
Dim bteColDesc As Byte
Dim bteColStartDate As Byte
Dim bteColCurr As Byte
Dim bteColPrice As Byte
Dim bteColQty As Byte
Dim bteColLastUpdate As Byte
Dim bteColLastuser As Byte
Dim bteColhcurr As Byte



Private Sub cbocurr_Change()
If cbocurr.ListIndex >= 0 Then
txtDesc = cbocurr.List(cbocurr.ListIndex, 1)
Else
txtDesc = ""
End If
End Sub

Private Sub CboItem_Change()
Call cboitem_Click
End Sub

Private Sub cboitem_Click()
    LblErrMsg = ""

    If cboitem.ListIndex <> -1 Then
        lblitem(0).Caption = "  " + cboitem.Column(1)
        lblitem(1).Caption = "  " + cboitem.Column(2)
                
    Else
        lblitem(0) = ""
        lblitem(1) = ""
                
        cboitem.SetFocus
        LblErrMsg.Caption = DisplayMsg(4003)
        Exit Sub
    End If
End Sub

Private Sub CboSupplier_Change()
LblErrMsg.Caption = ""
If cboSupplier.MatchFound = True Then
    Label1(2).Caption = cboSupplier.Column(1)
    Browse
Else
    Label1(2).Caption = ""
End If
End Sub

Private Sub cmdClear_Click()
Call Header
End Sub

Private Sub CmdSubMenu_Click()
  Unload Me
    frmMainMenu.Show
End Sub

Private Sub CmdSubmit_Click()
   Me.MousePointer = vbHourglass
  If hakUpdate(Me.Name) = 0 Then _
            LblErrMsg = DisplayMsg(3008): Me.MousePointer = vbDefault: Exit Sub
  
  
 If cboitem.Enabled = True And TextGrid = "" And validasi Then
        If save = True Then Call Browse: LblErrMsg.Caption = "Data Save Succes": Me.MousePointer = vbDefault
 End If
 
 If TextGrid = "D" And validasi Then
 
        With grid
           Dim LblInput As String
           For i = 1 To .Rows - 1
             If .TextMatrix(i, 0) = "D" Then
                   LblInput = MsgBox("Do you really want to delete ?", _
                    vbYesNo + vbQuestion, "Confirmation")
                    If LblInput = vbYes Then
                        'Call delete(Trim(cboSupplier.Text), Trim(.TextMatrix(i, bteColItemCode)))
                        LblErrMsg = DisplayMsg(1201)
                        TextGrid = ""
                     Else
                        Me.MousePointer = vbDefault
                        TextGrid = ""
                    End If
             
           Browse
           Exit Sub
             End If
           Next
           End With
           
 End If
  Me.MousePointer = vbDefault
 TextGrid = ""
End Sub

Private Function validasi() As Boolean

If TextGrid = "" Or TextGrid = "S" Then ' Validasi Update dan insert
        If cboSupplier.MatchFound = False Then cboSupplier.SetFocus: LblErrMsg.Caption = "Please Select valid Supplier..!": Me.MousePointer = vbDefault: Exit Function
        If cboitem.MatchFound = False Then cboitem.SetFocus: LblErrMsg.Caption = "Please Select valid Product Code..!": Me.MousePointer = vbDefault: Exit Function
        If cbocurr.MatchFound = False Then cbocurr.SetFocus: LblErrMsg.Caption = "Please Select valid Currency..!": Me.MousePointer = vbDefault: Exit Function
        If IsNumeric(txtprice.Text) Then
             If CDbl(txtprice.Text) <= 0 Then txtprice.SetFocus: LblErrMsg.Caption = "Please Input Price": Me.MousePointer = vbDefault: Exit Function
        Else
             txtprice.SetFocus: LblErrMsg.Caption = "Please Input Price": Me.MousePointer = vbDefault: Exit Function
        End If
        
        If IsNumeric(txtmaxqty.Text) Then
        
            If CDbl(txtmaxqty.Text) <= 0 Then txtmaxqty.SetFocus: LblErrMsg.Caption = "Please Input Max Qty": Me.MousePointer = vbDefault: Exit Function
        Else
            txtmaxqty.SetFocus: LblErrMsg.Caption = "Please Input Max Qty": Me.MousePointer = vbDefault: Exit Function
        End If
 End If
 
If TextGrid = "D" Then ' Validasi Delete
    If cboSupplier.MatchFound = False Then cboSupplier.SetFocus: LblErrMsg.Caption = "Please Select valid Supplier..!": Me.MousePointer = vbDefault: Exit Function
End If

    validasi = True
End Function


Private Function delete(TradeCode As String, ItemCode As String, StartDate As String) As Boolean
Dim sql As String

On Error GoTo abc

        sql = "Delete Price_Mold where Trade_Code='" & TradeCode & "' and Item_Code='" & ItemCode & "' and StartDate='" & StartDate & "'  "
        Db.Execute sql
        delete = True
 Exit Function
abc:
      delete = False

End Function
Private Function save() As Boolean
Dim sql As String
Dim EfekRow As Integer
On Error GoTo abc


sql = " UPDATE [KAWAI].[dbo].[Price_Mold] " & vbCrLf & _
            "    SET [StartDate] = '" & Format(dtsdate.Value, "yyyy-mm-dd") & "' " & vbCrLf & _
            "       ,[Curr] = '" & cbocurr.Text & "'" & vbCrLf & _
            "       ,[Price] = " & CDbl(txtprice.Text) & " " & vbCrLf & _
            "       ,[Qty] = " & CDbl(txtmaxqty.Text) & " " & vbCrLf & _
            "       ,[Last_update] = getdate() " & vbCrLf & _
            "       ,[Last_user] = '" & userLogin & "' " & vbCrLf & _
            "  WHERE Trade_Code='" & cboSupplier.Text & "'  and Item_Code='" & cboitem.Text & "' and StartDate='" & Format(dtsdate.Value, "yyyy-mm-dd") & "' " & vbCrLf & _
            "  "
Db.Execute sql, EfekRow
        
If EfekRow <= 0 Then

        sql = " INSERT INTO [dbo].[Price_Mold] " & vbCrLf & _
                    "            ([Trade_Code] " & vbCrLf & _
                    "            ,[Item_Code] " & vbCrLf & _
                    "            ,[StartDate] " & vbCrLf & _
                    "            ,[Curr] " & vbCrLf & _
                    "            ,[Price] " & vbCrLf & _
                    "            ,[Qty] " & vbCrLf & _
                    "            ,[Last_update] " & vbCrLf & _
                    "            ,[Last_user]) " & vbCrLf & _
                    "      VALUES " & vbCrLf & _
                    "            ('" & cboSupplier.Text & "' " & vbCrLf & _
                    "            ,'" & cboitem.Text & "' " & vbCrLf
        
        sql = sql + "        ,'" & Format(dtsdate.Value, "yyyy-mm-dd") & "' " & vbCrLf & _
                    "            ,'" & cbocurr.Text & "' " & vbCrLf & _
                    "            ," & CDbl(txtprice.Text) & " " & vbCrLf & _
                    "            ," & CDbl(txtmaxqty.Text) & " " & vbCrLf & _
                    "            ,getdate() " & vbCrLf & _
                    "            ,'" & userLogin & "' )" & vbCrLf & _
                    "  "
        Db.Execute sql
        
End If
save = True
Exit Function
abc:
      save = False

End Function
Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
  AddToComboSupplier
  Header
  adtocboitem
  Call up_FillCombo(cbocurr, "curr_cls")
End Sub
Sub AddToComboSupplier()
    
    Dim sqlcust As String
    Dim RsCust As New Recordset

    sqlcust = "select trade_code, trade_name, address1, country_cls, po_cls, Epte_Cls " & _
        "from trade_master where trade_cls='2' or Trade_Cls='3'"
        
    Set RsCust = Db.Execute(sqlcust)
    With cboSupplier
        .clear
        .columnCount = 2
        .ColumnWidths = "50pt;300pt;0pt;50pt;0pt;0pt"
        .ListWidth = 350
        .ListRows = 15
        i = 0
        Do While Not RsCust.EOF
            .AddItem
            .List(i, 0) = Trim(RsCust("Trade_code"))
            .List(i, 1) = IIf(IsNull(RsCust("trade_name")), " ", Trim(RsCust("Trade_Name")))
            RsCust.MoveNext
            i = i + 1
        Loop
        RsCust.Close
    End With
    
End Sub
Sub Header()
    With grid
        bteColSelect = 0
        bteColItemCode = 1
        bteColPartNumber = 2
        bteColDesc = 3
        bteColStartDate = 4
        bteColCurr = 5
        bteColPrice = 6
        bteColQty = 7
        bteColLastUpdate = 8
        bteColLastuser = 9
        bteColhcurr = 10
        
                        
        .clear
        .Rows = 1
        .ColS = 11
        
        .TextMatrix(0, bteColSelect) = " "
        .TextMatrix(0, bteColItemCode) = "Item Code"
        .TextMatrix(0, bteColPartNumber) = "Part Number"
        .TextMatrix(0, bteColDesc) = "Description"
        .TextMatrix(0, bteColStartDate) = "Start Date"
        .TextMatrix(0, bteColCurr) = "Curr"
        .TextMatrix(0, bteColPrice) = "Price"
        .TextMatrix(0, bteColQty) = "Max Qty"
        .TextMatrix(0, bteColLastUpdate) = "Last Update"
        .TextMatrix(0, bteColLastuser) = "Last User"
        .TextMatrix(0, bteColhcurr) = "H Curr"
        
        
        
        .ColWidth(bteColSelect) = 500
        .ColWidth(bteColItemCode) = 1800
        .ColWidth(bteColPartNumber) = 2000
        .ColWidth(bteColDesc) = 2000
        .ColWidth(bteColStartDate) = 1300
        .ColWidth(bteColPrice) = 1700
        .ColWidth(bteColQty) = 1700
        .ColWidth(bteColLastUpdate) = 2000
        .ColWidth(bteColLastuser) = 2000
                
        .ColDataType(bteColStartDate) = flexDTDate
        .ColHidden(bteColhcurr) = True

        
        .Cell(flexcpAlignment, 0, 0, 0, .ColS - 1) = flexAlignCenterCenter
        .EditMaxLength = 1
    End With
End Sub

Sub adtocboitem()
Dim sqlitem As String
Dim RsItem As New Recordset
Dim i As Long

   sqlitem = "select item_code, makeritem_code, item_name , finishgoodpart_cls from item_master " & _
          "where use_endday >= convert(char(8), getdate(), 112) "
    Set RsItem = Db.Execute(sqlitem)
    
    With cboitem
        .clear
        .columnCount = 3
        .ColumnWidths = "120pt;120pt;240pt;0pt"
        .ListWidth = 500
        .ListRows = 15
        
        i = 0
        Do While Not RsItem.EOF
            .AddItem
            .List(i, 0) = Trim(RsItem("item_code"))
            .List(i, 1) = Trim(RsItem("makeritem_code"))
            .List(i, 2) = Trim(RsItem("item_Name"))
            .List(i, 3) = Split(isiPart, ",")(Val(Trim(RsItem("finishgoodpart_cls"))) - 1)
            RsItem.MoveNext
            i = i + 1
        
        Loop
    End With
RsItem.Close
Set RsItem = Nothing
End Sub

Sub Browse()
    
    Dim RS As New Recordset
    Dim rsnama As New Recordset
    Dim rsreason As New Recordset
    Dim i As Long
    Dim nama As String, reason As String, sqlnama As String, sqlreason As String, p As Double
    Dim tglAwal As String, tglAkhir As String
    Dim strSQL As String
        
    strSQL = " SELECT A.Item_Code,B.MakerItem_Code,B.Item_Name,A.StartDate,A.CURR,C.Description " & vbCrLf & _
            "      ,A.Price,A.Qty,A.Last_Update,A.Last_User  " & vbCrLf & _
            "  FROM " & vbCrLf & _
            " (SELECT *  " & vbCrLf & _
            "   FROM [Price_Mold]  WHERE Trade_CODE='" & cboSupplier.Text & "')A " & vbCrLf & _
            "       LEFT JOIN dbo.Item_Master B ON A.Item_Code=B.Item_Code " & vbCrLf & _
            "       LEFT JOIN dbo.Curr_Cls C ON A.Curr=c.Curr_Cls " & vbCrLf

If RS.State <> adStateClosed Then RS.Close
    RS.Open strSQL, Db, adOpenForwardOnly, adLockReadOnly, adCmdText
Call Header
    i = 1
    If Not (RS.BOF And RS.EOF) Then
        
        With grid
            Do While Not RS.EOF
                .Rows = .Rows + 1

                .TextMatrix(i, bteColItemCode) = Trim(RS("Item_Code"))
                .TextMatrix(i, bteColPartNumber) = Trim(RS("MakerItem_Code"))
                .TextMatrix(i, bteColDesc) = Trim(RS("Item_Name"))
                .TextMatrix(i, bteColStartDate) = Format(RS("StartDate"), "dd MMM yyyy")
                .TextMatrix(i, bteColCurr) = Trim(RS("Description"))
                .TextMatrix(i, bteColPrice) = Format(RS("Price"), gs_formatAmountIDR)
                .TextMatrix(i, bteColQty) = Format(RS("Qty"), gs_formatAmountIDR)
                .TextMatrix(i, bteColLastUpdate) = Trim(RS("Last_Update"))
                .TextMatrix(i, bteColLastuser) = Trim(RS("Last_User"))
                
                
                .TextMatrix(i, bteColhcurr) = Trim(RS("CURR"))
                               
                RS.MoveNext
                i = i + 1
            Loop
            
        End With
    Else
        
        Header
        
    End If
    RS.Close
    
    Set rsnama = Nothing
    Set rsreason = Nothing
    Set RS = Nothing
    
End Sub
Private Sub kosongColGrid(Optional Kolom As String)
    Dim i As Long
    
    With grid
        .Col = bteColSelect
    
        If Kolom <> "" Then
           For i = 1 To .Rows - 1

              If .Text = Kolom Then .Text = ""
              If .TextMatrix(i, bteColSelect) <> "D" Then .TextMatrix(i, bteColSelect) = ""
           Next i
           'kosonggrid
        Else
           For i = 1 To .Rows - 1

              If .TextMatrix(i, bteColSelect) <> "" Then .TextMatrix(i, bteColSelect) = ""

           Next i
        End If
    End With
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)

Dim k As Boolean
Dim j As Integer

k = False
With grid
    TextGrid = UCase(grid.Text)

    If TextGrid = "S" Then
        
        
        cboitem.Text = Trim(.TextMatrix(Row, bteColItemCode))
        cboitem.Enabled = False
        dtsdate.Value = Trim(.TextMatrix(Row, bteColStartDate))
        cbocurr.Text = Trim(.TextMatrix(Row, bteColhcurr))
        txtprice.Text = Trim(.TextMatrix(Row, bteColPrice))
        txtmaxqty.Text = Trim(.TextMatrix(Row, bteColQty))
        
        ubah = True
       Call kosongColGrid
    ElseIf TextGrid = "D" Then
       Call kosongColGrid("S")
    End If
    
    .TextMatrix(Row, Col) = TextGrid
        
    For j = 1 To .Rows - 1
        If .TextMatrix(j, bteColSelect) <> "" Then
            k = True
        End If
    Next j

End With
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
  If grid.Col <> bteColSelect Then Cancel = True
End Sub

Private Sub txtmaxqty_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then
      KeyAscii = 0
End If
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub txtPrice_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then
      KeyAscii = 0
End If
If KeyAscii = 13 Then SendKeys vbTab
End Sub
