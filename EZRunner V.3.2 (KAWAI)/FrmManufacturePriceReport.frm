VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmManufacturePriceReport 
   BackColor       =   &H00FDDFE3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Finish Good / WIP Manufacture Price Report"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10110
   Icon            =   "FrmManufacturePriceReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   10110
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FDDFE3&
      Caption         =   "Finance Report"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   300
      TabIndex        =   14
      Top             =   1410
      Width           =   2055
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4710
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Height          =   1950
      Left            =   285
      TabIndex        =   7
      Top             =   1830
      Width           =   9390
      Begin VB.CommandButton cmdBrowser 
         Caption         =   "..."
         Height          =   300
         Left            =   4665
         TabIndex        =   16
         Top             =   825
         Width           =   300
      End
      Begin MSComCtl2.DTPicker DtPeriod 
         Height          =   330
         Left            =   2070
         TabIndex        =   2
         Top             =   1260
         Width           =   1335
         _ExtentX        =   2355
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
      Begin VB.Line Line8 
         Index           =   1
         X1              =   5085
         X2              =   9105
         Y1              =   675
         Y2              =   675
      End
      Begin VB.Label lblFinishGood_Cls 
         AutoSize        =   -1  'True
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
         Left            =   5055
         TabIndex        =   15
         Top             =   405
         Width           =   4050
      End
      Begin MSForms.ComboBox CboCls 
         Height          =   330
         Left            =   2070
         TabIndex        =   0
         Top             =   360
         Width           =   1335
         VariousPropertyBits=   746604571
         BackColor       =   16777215
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "2355;582"
         ListRows        =   15
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox CboProduk 
         Height          =   330
         Left            =   2070
         TabIndex        =   1
         Top             =   810
         Width           =   2565
         VariousPropertyBits=   746604571
         MaxLength       =   6
         DisplayStyle    =   3
         Size            =   "4524;582"
         ListRows        =   15
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblNm 
         AutoSize        =   -1  'True
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
         Left            =   5055
         TabIndex        =   11
         Top             =   840
         Width           =   4050
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
         Left            =   180
         TabIndex        =   10
         Top             =   870
         Width           =   1155
      End
      Begin VB.Line Line8 
         Index           =   0
         X1              =   5055
         X2              =   9105
         Y1              =   1110
         Y2              =   1110
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
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   9
         Top             =   1350
         Width           =   1110
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
         Left            =   180
         TabIndex        =   8
         Top             =   390
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   600
      Left            =   285
      TabIndex        =   5
      Top             =   3900
      Width           =   9390
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
         TabIndex        =   6
         Top             =   210
         Width           =   8985
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
      Left            =   8460
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4710
      Width           =   1185
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   405
      Left            =   7830
      TabIndex        =   13
      Top             =   540
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   714
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Finish Good / WIP Manufacture Price"
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
      Left            =   2655
      TabIndex        =   12
      Top             =   540
      Width           =   4185
   End
End
Attribute VB_Name = "FrmManufacturePriceReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS As New ADODB.Recordset
Dim term As Integer
Dim SqlData, sQlcls, basecurr, sQlproduk, sQlcost, TampungDt As String
Dim dateUp As Date

Dim booClsCange As Boolean

Private Sub CboCls_LostFocus()
    If booClsCange Then CariProduk
    booClsCange = False
End Sub

Private Sub cmdBrowser_Click()
 Me.MousePointer = vbHourglass
 frm_BrowseItem.getItemCode = CboProduk.Text
 frm_BrowseItem.Show 1
 CboProduk.Text = frm_BrowseItem.getItemCode
 Me.MousePointer = vbDefault
End Sub

Private Sub DtPeriod_change()
If Format(DtPeriod.Value, "MM") < Format(dateUp, "MM") And Val(Format(DtPeriod.Value, "MM")) = 1 And Val(Format(dateUp, "MM")) = 12 Then _
            DtPeriod.Year = DtPeriod.Year + 1: GoTo pass
    If Format(DtPeriod.Value, "MM") > Format(dateUp, "MM") And Val(Format(DtPeriod.Value, "MM")) = 12 And Val(Format(dateUp, "MM")) = 1 Then _
            DtPeriod.Year = DtPeriod.Year - 1
pass:
    dateUp = Format(DtPeriod.Value, "dd MMM yyyy")
End Sub

Private Sub cbocls_Change()
  If cboCls.MatchFound Then
     lblFinishGood_Cls = cboCls.Column(1)
  Else
     lblFinishGood_Cls = ""
  End If
  cboCls_Click
End Sub

Private Sub cmdReport_Click()
Me.MousePointer = vbHourglass
  If cboCls.ListIndex = -1 Then
        LblErrMsg = DisplayMsg(8057) '"Please Input Finish Good Part Cls"
        Exit Sub
 End If
  
  If CboProduk.ListIndex = -1 Then
        LblErrMsg = DisplayMsg(4061) ' "Please Input Produk"
        Exit Sub
  End If
  
  Dim ls_sql As String
  Dim RS As New ADODB.Recordset
  Call up_DropSQLFunctionALL
  Call up_CreateSQLFunctionALL
  ls_sql = "  " & vbCrLf & _
                  " Declare @Year as char(4) " & vbCrLf & _
                  " Declare @Month as char(4) " & vbCrLf & _
                  "  " & vbCrLf & _
                  " Set @Year = '" & Format(DtPeriod, "yyyy") & "' " & vbCrLf & _
                  " Set @Month = '" & Format(DtPeriod, "MM") & "'" & vbCrLf & _
                  "  " & vbCrLf & _
                  " select   " & vbCrLf & _
                  "     TB.item_code,    " & vbCrLf & _
                  "     IM.Item_name, " & vbCrLf & _
                  "     TB.receipt_date,  "

ls_sql = ls_sql + "     Lot_no=isnull((Select lot_no from daily_Production dp where dp.Seq_No=TB.DailySeq_No),''),  " & vbCrLf & _
                  "     isnull(WM.Working_Time,0) Working_Time, " & vbCrLf & _
                  "     isnull(WM.TotalLoss_Time,0) TotalLoss_time, " & vbCrLf & _
                  "     isnull(WM.TotalWorking_Time,0) TotalWorking_Time, " & vbCrLf & _
                  "     TB.Qty, " & vbCrLf & _
                  "     UC.Description Unit_Description, " & vbCrLf & _
                  "     Material_Cost= dbo.UF_GetAverageProductionCost( year(TB.receipt_date), month(TB.receipt_date),rtrim(TB.Item_Code) ,TB.Seq_No,'Material'), " & vbCrLf & _
                  "     Process_Cost= dbo.UF_GetAverageProductionCost( year(TB.receipt_date), month(TB.receipt_date),rtrim(TB.Item_Code) ,TB.Seq_No,'ProcessCost'), " & vbCrLf & _
                  "     Additional_Cost= dbo.UF_GetAverageProductionCost( year(TB.receipt_date), month(TB.receipt_date),rtrim(TB.Item_Code) ,TB.Seq_No,'AdditionalCost'), " & vbCrLf & _
                  "     TotalCost= dbo.UF_GetAverageProductionCost( year(TB.receipt_date), month(TB.receipt_date),rtrim(TB.Item_Code) ,TB.Seq_No,'ALL') " & vbCrLf & _
                  " from part_receipt  TB "

ls_sql = ls_sql + " Left Join Item_Master IM on TB.Item_Code=IM.Item_Code " & vbCrLf & _
                  " Left Join Unit_Cls UC on IM.Unit_Cls= UC.Unit_Cls " & vbCrLf & _
                  " Left Join WorkingTime_Master WM on WM.ProductionSeq_No= TB.Seq_No " & vbCrLf & _
                  " where TB.receipt_cls in ('P1')   " & vbCrLf & _
                  " and month(TB.receipt_date) = @Month " & vbCrLf & _
                  " and year(TB.receipt_date) = @Year "
                  
ls_sql = ls_sql + sQlproduk + sQlcls
                  
ls_sql = ls_sql + " order by TB.Item_Code, TB.Receipt_Date "

If RS.State <> adStateClosed Then RS.Close
RS.Open ls_sql, Db, adOpenKeyset, adLockOptimistic

Call up_DropSQLFunctionALL

If RS.EOF = True Then
    LblErrMsg = DisplayMsg(4006)
    Me.MousePointer = vbDefault
    If RS.State <> adStateClosed Then RS.Close
    Exit Sub
End If
LblErrMsg = ""
Dim xlapp As New Excel.application

Dim Idx As Integer

Dim xlColProductCode As String
Dim xlColProductName As String
Dim xlColDate As String
Dim xlColLotNo As String
Dim xlColQty As String
Dim xlColUnit As String
Dim xlColMaterialCost As String
Dim xlColMaterialCostPerUnit As String
Dim xlColWorkingTime As String
Dim xlColLossTime As String
Dim xlColTotalWorkingTime As String
Dim xlColProcessCost As String
Dim xlColProcessCostPerUnit As String
Dim xlColAdditionalCost As String
Dim xlColAdditionalCostPerUnit As String
Dim xlColTotalCost As String
Dim xlColTotalCostPerUnit As String

 xlColProductCode = "a"
 xlColProductName = "b"
 xlColDate = "c"
 xlColLotNo = "d"
 xlColQty = "e"
 xlColUnit = "f"
 xlColMaterialCost = "g"
 xlColMaterialCostPerUnit = "h"
  xlColWorkingTime = "i"
 xlColLossTime = "j"
 xlColTotalWorkingTime = "k"
 xlColProcessCost = "l"
 xlColProcessCostPerUnit = "m"
 xlColAdditionalCost = "n"
 xlColAdditionalCostPerUnit = "o"
 xlColTotalCost = "p"
 xlColTotalCostPerUnit = "q"
 
With xlapp
        .Workbooks.Add

        .Range(xlColProductCode & "2", xlColTotalCostPerUnit & "2").Merge
        .Range(xlColProductCode & "2") = "FINISH GOOD / WIP MANUFACTURE PRICE"
        .Range(xlColProductCode & "2").horizontalAlignment = xlCenter
        .Range(xlColProductCode & "2").Font.Bold = True
             
        .Range(xlColProductCode & "3:" & xlColTotalCostPerUnit & "3").Merge
        .Range(xlColProductCode & "3") = "Base Currency : " + uf_GetCurrencyDescription(gs_DefaultCurrencyCode)
        .Range(xlColProductCode & "3").horizontalAlignment = xlCenter
        .Range(xlColProductCode & "3").Font.Bold = False
 
        .Range(xlColProductCode & "4") = "Finish Good Part Cls " + " : " + cboCls.Text
        .Range(xlColProductCode & "5").NumberFormat = "@"
        .Range(xlColProductCode & "5") = "Product Code " + " : " + CboProduk.List(CboProduk.ListIndex, 0)
        .Range(xlColProductCode & "6") = "Period " + " : " + Format(DtPeriod.Value, "MMM yyyy")
        .Range(xlColAdditionalCostPerUnit & "6", xlColTotalCostPerUnit & "6").Merge
        .Range(xlColAdditionalCostPerUnit & "6") = "Issued Date : " + Format(Now, "dd MMM yyyy  hh:MM:ss")
        .Range(xlColAdditionalCostPerUnit & "6").horizontalAlignment = xlRight
        .Range(xlColProductCode & "4", xlColProductName & "4").Merge
        .Range(xlColProductCode & "5", xlColProductName & "5").Merge
        .Range(xlColProductCode & "6", xlColProductName & "6").Merge
        .Range(xlColProductCode & "6", xlColProductName & "6").Merge
        
        .Range(xlColProductCode & "8") = "Product Code"
        .Range(xlColProductName & "8") = "Description"
        .Range(xlColDate & "8") = "Production Date"
        .Range(xlColLotNo & "8") = "Lot No."
        .Range(xlColQty & "8") = "Qty"
        .Range(xlColUnit & "8") = "Unit"
        .Range(xlColMaterialCost & "8") = "Material Cost"
        .Range(xlColMaterialCostPerUnit & "8") = "Material Cost" & vbCrLf & "[Per unit]"
        .Range(xlColProcessCost & "8") = "Process Cost"
        .Range(xlColProcessCostPerUnit & "8") = "Process Cost" & vbCrLf & "[Per unit]"
        .Range(xlColAdditionalCost & "8") = "Other Cost"
        .Range(xlColAdditionalCostPerUnit & "8") = "Other Cost" & vbCrLf & "[Per unit]"
        .Range(xlColTotalCost & "8") = "Total Cost"
        .Range(xlColTotalCostPerUnit & "8") = "Total Cost" & vbCrLf & "[Per unit]"
        
        .Range(xlColWorkingTime & "8") = "Working Time" & vbCrLf & "[Minute]"
        .Range(xlColLossTime & "8") = "Loss Time" & vbCrLf & "[Minute]"
        .Range(xlColTotalWorkingTime & "8") = "Total Working Time" & vbCrLf & "[Minute]"

        .Range(xlColProductCode & "8", xlColTotalCostPerUnit & "8").horizontalAlignment = xlCenter
        .Range(xlColProductCode & "8", xlColTotalCostPerUnit & "8").verticalAlignment = xlCenter
        
          Idx = 8
        Dim ld_TotalQty As Double
        Dim ld_TotalCost As Double
        ld_TotalQty = 0
        ld_TotalCost = 0
        Dim ls_PreviousItemCode As String
            '#Fill Data
            While RS.EOF = False
                Idx = Idx + 1
               
                If Idx <> 9 And ls_PreviousItemCode <> Trim(RS!Item_Code) Then
                
                    .Range(xlColProductCode & Idx & ":" & xlColLotNo & Idx).Merge
                    .Range(xlColUnit & Idx & ":" & xlColMaterialCostPerUnit & Idx).Merge
                    .Range(xlColWorkingTime & Idx & ":" & xlColTotalWorkingTime & Idx).Merge
                    .Range(xlColProcessCost & Idx & ":" & xlColTotalCostPerUnit & Idx).Merge
                    .Range(xlColProductCode & Idx) = "Sub Total"
                    
                    .Range(xlColProductCode & Idx & ":" & xlColTotalCostPerUnit & Idx).Font.Bold = True
                    .Range(xlColProductCode & Idx & ":" & xlColAdditionalCostPerUnit & Idx).horizontalAlignment = xlRight
                    .Range(xlColQty & Idx) = ld_TotalQty
                    .Range(xlColTotalCost & Idx) = ld_TotalCost
                    ld_TotalQty = 0
                    ld_TotalCost = 0
                    Idx = Idx + 1
                End If
                .Range(xlColProductCode & Idx) = RS!Item_Code
                .Range(xlColProductName & Idx) = RS!item_name
                .Range(xlColDate & Idx) = Format(RS!Receipt_Date)
                .Range(xlColLotNo & Idx) = RS!Lot_no
                .Range(xlColQty & Idx) = RS!Qty
                .Range(xlColUnit & Idx) = RS!unit_description
                .Range(xlColWorkingTime & Idx) = RS!Working_Time
                .Range(xlColLossTime & Idx) = RS!TotalLoss_Time
                .Range(xlColTotalWorkingTime & Idx) = RS!TotalWorking_Time
                .Range(xlColMaterialCost & Idx) = RS!material_Cost
                .Range(xlColMaterialCostPerUnit & Idx) = RS!material_Cost / RS!Qty
                .Range(xlColProcessCost & Idx) = RS!process_cost
                .Range(xlColProcessCostPerUnit & Idx) = RS!process_cost / RS!Qty
                .Range(xlColAdditionalCost & Idx) = RS!Additional_Cost
                .Range(xlColAdditionalCostPerUnit & Idx) = RS!Additional_Cost / RS!Qty
                .Range(xlColTotalCost & Idx) = RS!TotalCost
                .Range(xlColTotalCostPerUnit & Idx) = RS!TotalCost / RS!Qty
                
                ld_TotalQty = ld_TotalQty + RS!Qty
                ld_TotalCost = ld_TotalCost + RS!TotalCost
                 ls_PreviousItemCode = Trim(RS!Item_Code)
                RS.MoveNext
            Wend
            
             Idx = Idx + 1
            .Range(xlColProductCode & Idx & ":" & xlColLotNo & Idx).Merge
            .Range(xlColUnit & Idx & ":" & xlColMaterialCostPerUnit & Idx).Merge
            .Range(xlColWorkingTime & Idx & ":" & xlColTotalWorkingTime & Idx).Merge
            .Range(xlColProcessCost & Idx & ":" & xlColTotalCostPerUnit & Idx).Merge
            .Range(xlColProductCode & Idx) = "Sub Total"
            .Range(xlColQty & Idx) = ld_TotalQty
            .Range(xlColTotalCost & Idx) = ld_TotalCost
            
           .Range(xlColProductCode & Idx & ":" & xlColTotalCostPerUnit & Idx).Font.Bold = True
           .Range(xlColProductCode & Idx & ":" & xlColAdditionalCostPerUnit & Idx).horizontalAlignment = xlRight
            
            '#Run Macro
            .Range(xlColProductCode & "2:" & xlColTotalCostPerUnit & "2").Select
            With .Selection.Font
                .Size = 18
            End With
            
            .Range(xlColProductCode & "8:" & xlColTotalCostPerUnit & Idx).Select
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
        
            .Columns(xlColProductCode & ":" & xlColProductCode).columnWidth = 14.29
            .Columns(xlColProductName & ":" & xlColProductName).columnWidth = 23.57
            .Columns(xlColDate & ":" & xlColDate).columnWidth = 14.29
            .Columns(xlColLotNo & ":" & xlColLotNo).columnWidth = 16.43
            .Columns(xlColQty & ":" & xlColQty).columnWidth = 10
            .Columns(xlColUnit & ":" & xlColUnit).columnWidth = 5.71
            .Columns(xlColMaterialCost & ":" & xlColMaterialCost).columnWidth = 14.57
            .Columns(xlColMaterialCostPerUnit & ":" & xlColMaterialCostPerUnit).columnWidth = 14.57
            .Columns(xlColWorkingTime & ":" & xlColWorkingTime).columnWidth = 14.57
            .Columns(xlColLossTime & ":" & xlColLossTime).columnWidth = 14.57
            .Columns(xlColTotalWorkingTime & ":" & xlColTotalWorkingTime).columnWidth = 14.57
            .Columns(xlColProcessCost & ":" & xlColProcessCost).columnWidth = 14.57
            .Columns(xlColProcessCostPerUnit & ":" & xlColProcessCostPerUnit).columnWidth = 14.57
            .Columns(xlColAdditionalCost & ":" & xlColAdditionalCost).columnWidth = 14.57
            .Columns(xlColTotalCost & ":" & xlColTotalCost).columnWidth = 14.57
            .Columns(xlColTotalCostPerUnit & ":" & xlColTotalCostPerUnit).columnWidth = 14.57
              
            .Range(xlColWorkingTime & "9:" & xlColTotalWorkingTime & Idx).Select
              With .Selection
                  .NumberFormat = gs_formatWorkingTime
              End With

        
              .Range(xlColQty & "9:" & xlColQty & Idx).Select
              With .Selection
                  .NumberFormat = gs_formatQty
              End With
              
              .Range(xlColMaterialCost & "9:" & xlColMaterialCostPerUnit & Idx).Select
              With .Selection
                  .NumberFormat = gs_formatPrice
              End With
        
              .Range(xlColProcessCost & "9:" & xlColTotalCostPerUnit & Idx).Select
              With .Selection
                  .NumberFormat = gs_formatPrice
              End With
        .Range("A1:A1").Select
        If Check1.Value = 1 Then
            .Range(xlColWorkingTime & "8:" & xlColTotalWorkingTime & Idx).Select
            .Selection.delete Shift:=xlToLeft
        End If

        
            .Visible = True
            .ActiveSheet.PageSetup.PaperSize = xlPaperA4
            .ActiveSheet.PageSetup.Orientation = 2
            .WindowState = xlMaximized
            .ActiveWindow.Zoom = 80
End With
        
Me.MousePointer = vbDefault
If RS.State <> adStateClosed Then RS.Close

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

Private Sub cboCls_Click()
    With cboCls
        If .ListIndex < 0 Then
            sQlcls = ""
            Exit Sub
        Else
            Select Case .ListIndex
                Case 0: sQlcls = ""
                Case 1: sQlcls = " and finishgoodpart_cls = '01' "
                Case 2: sQlcls = " and finishgoodpart_cls = '02' "
            End Select
        End If
    End With
    booClsCange = True
End Sub

Private Sub CboProduk_Change()
    If CboProduk.MatchFound Then
      lblNm.Caption = CboProduk.List(CboProduk.ListIndex, 1)
      If Trim(CboProduk.List(CboProduk.ListIndex, 0)) = strAll Then
        sQlproduk = ""
      Else
        sQlproduk = " and im.item_code = '" & CboProduk.List(CboProduk.ListIndex, 0) & "' "
      End If
    Else
      lblNm = ""
    End If
End Sub

Private Sub CboProduk_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = vbKeyReturn Then CboProduk_Change
End Sub

Private Sub CmdSubMenu_Click()
    DoEvents
    frmMainMenu.Show
    DoEvents
    Unload Me
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
    CtrlMenu1.FormName = Me.Name
    Me.Caption = Me.Caption & " (Menu ID : " & CtrlMenu1.MenuText & ")"
    SqlData = ""
    sQlcls = ""
    lblNm = ""
    DtPeriod.Value = Now
    dateUp = DtPeriod.Value
    
    cboCls.clear
    cboCls.columnCount = 2
    
    cboCls.AddItem
    cboCls.List(0, 0) = strAll
    cboCls.List(0, 1) = strAll
    cboCls.AddItem
    cboCls.List(1, 0) = "01"
    cboCls.List(1, 1) = "Finish Goods"
    cboCls.AddItem
    cboCls.List(2, 0) = "02"
    cboCls.List(2, 1) = "Parts/WIP/Material"
    
    cboCls.ListWidth = 120
    cboCls.ColumnWidths = "30 pt ; 90 pt "
    cboCls.ListIndex = 0
    cboCls.Text = cboCls.List(0, 0)
    
    CariProduk
    booClsCange = False
    
End Sub

Sub CariProduk()
Dim i As Long
lblNm = ""
    With CboProduk
        .clear
        .columnCount = 2
        .ColumnWidths = "130pt;200pt"
        .ListWidth = 330
        .ListRows = 15
    
    SqlData = " select item_code, item_name from item_master where use_endday >= convert(char(8), getdate(), 112) "
    If sQlcls <> "" Then SqlData = SqlData + sQlcls
    
    RS.Open SqlData, Db, 1, 3
    i = 0
    .AddItem ""
    .List(i, 0) = strAll
    .List(i, 1) = strAll
    i = 1
    While Not RS.EOF
        .AddItem ""
        .List(i, 0) = Trim$(RS!Item_Code)
        .List(i, 1) = Trim$(RS!item_name)
        RS.MoveNext
        i = i + 1
    Wend
    RS.Close
    .ListIndex = 0
    End With
End Sub
