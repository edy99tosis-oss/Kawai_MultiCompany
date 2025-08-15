VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmMainMenu 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EZ Runner ver.3 - Main Menu"
   ClientHeight    =   9105
   ClientLeft      =   1440
   ClientTop       =   1125
   ClientWidth     =   12240
   Icon            =   "frmMainMenu.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmMainMenu.frx":0E42
   MousePointer    =   99  'Custom
   ScaleHeight     =   9105
   ScaleWidth      =   12240
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   4995
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMenu.frx":114C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMenu.frx":1F9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMenu.frx":2DF0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView Tree 
      Height          =   9105
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   16060
      _Version        =   393217
      Indentation     =   441
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList2"
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox vbalImageList1 
      Height          =   480
      Left            =   3690
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   7
      Top             =   4560
      Width           =   1200
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4920
      Top             =   2370
   End
   Begin VB.CommandButton cmdLogout 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Log&out"
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
      Left            =   10890
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8415
      UseMaskColor    =   -1  'True
      Width           =   1000
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Height          =   600
      Left            =   5670
      TabIndex        =   0
      Top             =   7740
      Width           =   6210
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
         Height          =   240
         Left            =   60
         TabIndex        =   1
         Top             =   240
         Width           =   6090
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   510
      Top             =   8130
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMenu.frx":3C42
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMenu.frx":3F5C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblSim 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Simulation"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   525
      Left            =   6420
      TabIndex        =   8
      Top             =   4965
      Width           =   2010
   End
   Begin VB.Label Copyright 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ver.3"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   225
      Left            =   11430
      TabIndex        =   3
      Top             =   7560
      Width           =   435
   End
   Begin VB.Image Image4 
      Height          =   270
      Left            =   11040
      Picture         =   "frmMainMenu.frx":4DAE
      Top             =   120
      Width           =   990
   End
   Begin VB.Image Image2 
      Height          =   1110
      Index           =   0
      Left            =   5400
      Picture         =   "frmMainMenu.frx":52C4
      Top             =   6600
      Width           =   4200
   End
   Begin VB.Image Image2 
      Height          =   3135
      Index           =   1
      Left            =   5640
      Picture         =   "frmMainMenu.frx":145D6
      Top             =   1830
      Width           =   6270
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   1
      Left            =   7830
      TabIndex        =   5
      Top             =   1050
      Width           =   1830
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EZ Runner"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   6510
      TabIndex        =   4
      Top             =   690
      Visible         =   0   'False
      Width           =   4485
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      Height          =   795
      Left            =   90
      Top             =   2760
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000013&
      BackStyle       =   1  'Opaque
      Height          =   855
      Left            =   120
      Top             =   1560
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   4350
      Left            =   6240
      Picture         =   "frmMainMenu.frx":54780
      Stretch         =   -1  'True
      Top             =   2010
      Visible         =   0   'False
      Width           =   4425
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ParentCode As Node
Dim bolshow As Boolean

Sub NoAkses()
    lblErrMsg = DisplayMsg(3007)
End Sub

Private Sub Form_Load()
    lblErrMsg = ""
    bolshow = True
    lblSim.Visible = gb_Simulation
End Sub

Private Sub cmdLogout_Click()
    DoEvents
    frmLogin.Show
    frmLogin.txtPass = ""
    frmLogin.txtPass = "nec"
    DoEvents
    Me.Hide
    lblErrMsg = ""
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub Form_Resize()
    Tree.Height = Me.Height - 350
    If Me.WindowState = vbMaximized Then
        Label1(0).Left = Me.Width - 7400
        Label1(1).Left = Me.Width - 6100
        Image1.Left = Me.Width - 7300
        Frame1.Left = Me.Width - 8100
        Image4.Left = Me.Width - 2700
        cmdLogout.Left = Me.Width - 2900
        Image2(0).Left = Me.Width - 8400
        Image2(1).Left = Me.Width - 8400
        Label1(0).top = Me.Height - 10500
        Label1(1).top = Me.Height - 10000
        Image1.top = Me.Height - 8300
        Frame1.top = Me.Height - 2500
        Image2(0).top = Me.Height - 3400
        Image2(1).top = Me.Height - 9000
        cmdLogout.top = Me.Height - 1500
        Copyright.Left = Me.Width - 2300
        Copyright.top = Me.Height - 2700
    Else
        Label1(0).Left = Me.Width - 5800
        Label1(1).Left = Me.Width - 4200
        Image1.Left = Me.Width - 5800
        Image2(0).Left = Me.Width - 6900
        Image2(1).Left = Me.Width - 6800
        Image4.Left = Me.Width - 1300
        Frame1.Left = Me.Width - 6600
        cmdLogout.Left = Me.Width - 1400
        Label1(0).top = Me.Height - 8700
        Label1(1).top = Me.Height - 8200
        Image1.top = Me.Height - 7100
        Frame1.top = Me.Height - 2000
        Image2(0).top = Me.Height - 2900
        Image2(1).top = Me.Height - 7600
        cmdLogout.top = Me.Height - 1300
        Copyright.Left = Me.Width - 820
        Copyright.top = Me.Height - 2180
    End If
End Sub

Private Sub Tree_Click()
    bolshow = True
End Sub

Private Sub Tree_DblClick()
    Dim X As Integer
    
    lblErrMsg = ""
    If bolshow = False Then bolshow = True: Exit Sub
    
    DoEvents
    X = Len(Tree.selectedItem.Key) - 1
    Select Case Tree.selectedItem.Key
        
        '**************** Master ****************
        Case "nItem Master":
            If hakAkses("frm_item_master2") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            
            frm_item_master2.Show
            Me.Hide
        Case "nItem Inquiry":
            If hakAkses("frm_item_inquiry") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql


            frm_item_inquiry.Show
            Me.Hide
        Case "nManufacture Master":
            If hakAkses("F_ManufactureLine") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql


            F_ManufactureLine.Show
            Me.Hide
        Case "nWarehouse Master":
            If hakAkses("FrmWarehouse") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmWarehouse.Show
            Me.Hide
        Case "nTrade Master":
            If hakAkses("FrmTradeMaster") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmTradeMaster.Show
            Me.Hide
        Case "nTrade Master Inquiry":
            If hakAkses("FrmTradeMasterInquiry") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmTradeMasterInquiry.Show
            Me.Hide
        Case "nBOM Master":
            If hakAkses("FrmBOMMaster") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmBOMMaster.Show
            Me.Hide
        Case "nBOM Inquiry":
            If hakAkses("frmBOMInquiry") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmBOMInquiry.Show
            Me.Hide
        Case "nClassification Master":
            If hakAkses("FRM_CLS") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FRM_CLS.Show
            Me.Hide
        Case "nTax Classification":
            If hakAkses("F_TAX") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            F_TAX.Show
            Me.Hide
        Case "nPrice Master":
            If hakAkses("frmPriceMaster") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            If hakPrice("frmPriceMaster") = 0 Then lblErrMsg = DisplayMsg("0006"): frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmPriceMaster.Show
            Me.Hide
        Case "nCalendar Master":
            If hakAkses("frmCalendar") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmCalendar.Show
            Me.Hide
        Case "nTax Exchange Rate":
            If hakAkses("frmtax_rate") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmTax_Rate.Show
            Me.Hide
        Case "nBank Master":
            If hakAkses("FrmBankMaster") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmBankMaster.Show
            Me.Hide
        Case "nDaily Exchange Rate":
            If hakAkses("FrmDailyExRate") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmDailyExRate.Show
            Me.Hide
        Case "nBook Keeping Exchange Rate":
            If hakAkses("F_EXCHANGE") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            F_EXCHANGE.Show
            Me.Hide
        Case "nReport Book Keeping Exchange Rate":
            If hakAkses("Frm_ExchReport") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            Frm_ExchReport.Show
            Me.Hide
         Case "nHS Master":
            If hakAkses("FrmHSMaster") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmHSmaster.Show
            Me.Hide
        Case "nPacking Master":
            If hakAkses("FrmPackingMaster") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmPackingMaster.Show
            Me.Hide
        Case "nMachine Master":
            If hakAkses("F_MachineMaster") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            F_MachineMaster.Show
            Me.Hide

        Case "nPurchase Set Master":
            If hakAkses("FrmPOSetMaster") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmPOSetMaster.Show
            Me.Hide
            
        Case "nBC Type Master":
            If hakAkses("FrmBCMaster") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmBCMaster.Show
            Me.Hide
            
        Case "nPrice Mold":
            If hakAkses("FrmPriceMol") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmPriceMol.Show
            Me.Hide
            
        Case "nEmail Configuration":
                   If hakAkses("FrmEmailConfig") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
                   
                   sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
                   Db.Execute sql
                   
                   FrmEmailConfig.Show
                   Me.Hide
                   
        Case "nBOM Masterlist":
                   If hakAkses("FrmBOMMasterlist") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
                   
                   sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
                   Db.Execute sql
                   
                   FrmBOMMasterlist.Show
                   Me.Hide
        
         Case "nPrice Master Contract":
            If hakAkses("frmPriceMasterContract") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmPriceMasterContract.Show
            Me.Hide
            
            
        '**************** Production Plan ****************
        Case "nProduction Planning (Forecast)":
            If hakAkses("Frm_Production_planning") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            Frm_Production_Planning.Show
            Me.Hide
        Case "nProduction Planning / Result Inquiry":
            If hakAkses("Frm_Production_Result") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            Frm_Production_Result.Show
            Me.Hide
        
        '**************** Order Control ****************
        Case "nOrder Entry (Update)":
            If hakAkses("frmOrderEntry") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmOrderEntry.Show
            Me.Hide
        Case "nOrder Inquiry":
            If hakAkses("frm_order_inquiry") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frm_order_inquiry.Show
            Me.Hide
        Case "nOrder Inquiry (Delivery Date)":
            If hakAkses("frm_order_inquiry_date") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frm_order_inquiry_date.Show
            Me.Hide
        Case "nOrder Entry Status":
            If hakAkses("frm_order_entry_status") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frm_order_entry_status.Show
            Me.Hide
        Case "nPurchase Order Adjustment Price":
            If hakAkses("FrmPOAdjustment_price") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmPOAdjustment_Price.Show
            Me.Hide
        
        Case "nSerial No Information":
            If hakAkses("frm_SerialNo_Information") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            Frm_SerialNo_Information.Show
            Me.Hide
            
        '**************** Delivery Note ****************

        Case "nDelivery Note Create":
            If hakAkses("frmDOCreate") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmDOCreate.Show
            Me.Hide
        Case "nDelivery Note Status":
            If hakAkses("frmDOStatus") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmDOStatus.Show
            Me.Hide
        Case "nDelivery Note Print Out (Customer & DN Date)":
            If hakAkses("Frm_SheetDODD") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FRM_SheetDODD.Show
            Me.Hide
        Case "nDelivery Note Print Out (DN Date)":
            If hakAkses("Frm_SheetDOD") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            Frm_SheetDOD.Show
            Me.Hide
        Case "nDelivery Note Detail List":
            If hakAkses("FrmSalesReport") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmSalesReport.Show
            Me.Hide
        Case "nDelivery Note Return":
            If hakAkses("frmDNReturn") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmDNReturn.Show
            Me.Hide
        Case "nDelivery Note Return [Unscheduled]":
            If hakAkses("frmDNReturnUnscheduled ") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmDNReturnUnscheduled.Show
            Me.Hide
            
        '**************** Packing List ****************
        Case "nPacking List Create":
            If hakAkses("FrmPackingCreate") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmPackingCreate.Show
            Me.Hide
        Case "nPacking List Status":
            If hakAkses("FrmPackingStatus") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmPackingStatus.Show
            Me.Hide
            
        '**************** Invoice ****************
        Case "nInvoice Create":
            If hakAkses("Frm_invoice_Create") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            Frm_invoice_Create.Show
            Me.Hide
        Case "nInvoice Create (Export)":
            If hakAkses("FrmInvoiceExport") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmInvoiceExport.Show
            Me.Hide
        Case "nInvoice Inquiry":
            If hakAkses("frm_invoice_inquiry") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frm_invoice_inquiry.Show
            Me.Hide
        Case "nInvoice Status":
            If hakAkses("Frm_Invoice_status") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            Frm_Invoice_Status.Show
            Me.Hide
        Case "nInvoice Detail List":
            If hakAkses("F_DetailSalesReport") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            F_DetailSalesReport.Show
            Me.Hide
        Case "nInvoice Summary List":
            If hakAkses("F_SummarySalesReport") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            F_SummarySalesReport.Show
            Me.Hide
        Case "nAR List":
            If hakAkses("frmARList") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            If hakPrice("frmARList") = 0 Then lblErrMsg = DisplayMsg("0006"): frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmARList.Show
            Me.Hide
        Case "nAR Progress Control":
            If hakAkses("frmAR_progress") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            If hakPrice("frmAR_progress") = 0 Then lblErrMsg = DisplayMsg("0006"): frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmAR_Progress.Show
            Me.Hide
        Case "nGeneral / Sub Ledger - AR":
            If hakAkses("frmRptLedgerAR") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmRptLedgerAR.Show
            Me.Hide
            
        '**************** Faktur Pajak ****************
        Case "nFaktur Pajak Create":
            If hakAkses("frmPajak_Create_New") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmPajak_Create_New.Show
            Me.Hide
        Case "nFaktur Pajak Status":
            If hakAkses("frmpajak_status") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmpajak_status.Show
            Me.Hide
        Case "nNota Retur":
            If hakAkses("frmNotaRetur") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmNotaRetur.Show
            Me.Hide
                    
        '**************** Purchase Request ****************
        Case "nPurchase Request (Part/Material)":
            If hakAkses("frmPRParts") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmPRParts.Show
            Me.Hide
        Case "nPurchase Request (Steel/Coil)":
            If hakAkses("frmPRSteelCoil") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmPRSteelCoil.Show
            Me.Hide
        Case "nPurchase Request (Other Item)":
            If hakAkses("frmPROther") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmPROther.Show
            Me.Hide
        Case "nPurchase Request Status":
            If hakAkses("frmPRStatus") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmPRStatus.Show
            Me.Hide
        Case "nPurchase Request Inquiry":
            If hakAkses("frmPRInquiry") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmPRInquiry.Show
            Me.Hide
        Case "nPurchase Request Progress Report":
            If hakAkses("frmPRProgress") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmPRProgress.Show
            Me.Hide
        
        '**************** Purchase Control ****************
'        Case "nPurchase Order Scheduled (Part/Material)":
'            If hakAkses("frmPOPartsScheduled ") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
'
'            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.SelectedItem.Key, x) & "',Getdate()) "
'            Db.Execute sql
'
'            frmPOPartsScheduled.Show
'            Me.Hide
        Case "nPurchase Order Scheduled (Steel/Coil)":
            If hakAkses("frmPOSteelCoilScheduled") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmPOSteelCoilScheduled.Show
            Me.Hide
        Case "nPurchase Order Scheduled (Subcon)":
            If hakAkses("frmPOSubconScheduled") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmPOSubconScheduled.Show
            Me.Hide
        Case "nPurchase Order Scheduled (Other Item)":
            If hakAkses("frmPOOtherScheduled") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmPOOtherScheduled.Show
            Me.Hide
        Case "nPurchase Order Part/Material":
            If hakAkses("frmPOParts") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmPOParts.Show
            Me.Hide
        Case "nPurchase Order Part/Material Upload":
            If hakAkses("frmPOPartsUpload") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
          frmPOPartsUpload.Show
            Me.Hide
        Case "nPurchase Order Unscheduled (Steel/Coil)":
            If hakAkses("frmPOSteelCoil") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmPOSteelCoil.Show
            Me.Hide
        Case "nPurchase Order Unscheduled (Other Item)":
            If hakAkses("frmPOOther") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmPOOther.Show
            Me.Hide
        Case "nPurchase Order Subcon"
            If hakAkses("frmPOSubcon") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmPOSubcon.Show
            Me.Hide
        Case "nPurchase Order Status":
            If hakAkses("frm_po_status") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frm_po_status.Show
            Me.Hide
         Case "nPurchase Order/Result Inquiry (Material)":
            If hakAkses("FrmPORecInquiryMaterial") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmPORecInquiryMaterial.Show
            Me.Hide
         Case "nPurchase Order/Result Inquiry (Supplier)":
            If hakAkses("FrmPOResultInquiry") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmPOResultInquiry.Show
            Me.Hide
         Case "nForecast for Part/Material":
            If hakAkses("frmforecast") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmForecast.Show
            Me.Hide
        Case "nPart/Material Receipt List":
            If hakAkses("frm_AccPay") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frm_accPay.Show
            Me.Hide
        Case "nUpdate Price Process":
            If hakAkses("FrmUpdatePriceProcess") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmUpdatePriceProcess.Show
            Me.Hide
        Case "nPurchase Order Part/Material (Set)":
            If hakAkses("frmPOPartsSET") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmPOPartsSet.Show
            Me.Hide
            
        Case "nPO Contract Inquiry":
            If hakAkses("FrmPOContractInquiry") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmPOContractInquiry.Show
            Me.Hide
        ' Tambahan 20120223
        ' -------------------------
        Case "nIncoming Material Report":
            If hakAkses("frm_IncomingMaterialReport") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frm_IncomingMaterialReport.Show
            Me.Hide
            
        Case "nPO Correction":
            If hakAkses("frmPOCorrection") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmPOCorrection.Show
            Me.Hide
            
        Case "nPO Correction Approval":
            If hakAkses("frmPOCorrectionApproval") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmPOCorrectionApproval.Show
            Me.Hide
            
        Case "nPurchase Order Contract":
            If hakAkses("FrmPOContract_Mst") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmPOContract_Mst.Show
            Me.Hide
            
         Case "nPurchase Order Contract Inquiry":
            If hakAkses("FrmPOContract_Inq") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmPOContract_Inq.Show
            Me.Hide
          
        '**************** Invoice Supplier / AP Control ****************
        Case "nAP Invoice Information Entry":
            If hakAkses("FrmAPInvoiceInformationEntry") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmAPInvoiceInformationEntry.Show
            Me.Hide
        Case "nPayment Amount Entry":
            If hakAkses("FrmApList") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            If hakPrice("FrmApList") = 0 Then lblErrMsg = DisplayMsg("0006"): frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmApList.Show
            Me.Hide
        Case "nAccount Payable Progress":
            If hakAkses("FrmAPProgress") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            If hakPrice("FrmAPProgress") = 0 Then lblErrMsg = DisplayMsg("0006"): frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmAPProgress.Show
            Me.Hide
        Case "nAP Status":
            If hakAkses("frmAPStatus") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmAPStatus.Show
            Me.Hide
        Case "nGeneral / Sub Ledger - AP":
            If hakAkses("frmRptLedgerAP") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmRptLedgerAP.Show
            Me.Hide
        Case "nAP Detail List Report":
            If hakAkses("FrmAPListReport") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmAPListReport.Show
            Me.Hide
            
        '**************** Stock Control ****************
        Case "nParts (Material) Receipt [Schedule]":
            If hakAkses("Frmpart_rec") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmPart_Rec.Show
            Me.Hide
        Case "nParts (Material) Receipt [Unscheduled]":
            If hakAkses("Frmpart_recUn") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmPart_RecUn.Show
            Me.Hide
        Case "nParts (Material) Supply Request [Automatic]":
            If hakAkses("frm_part_SupplyAuto") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frm_ProdResultAutoRequest.Show
            Me.Hide
        Case "nParts (Material) Supply [Automatic]":
            If hakAkses("frm_part_supplyIn") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frm_part_supplyIn.Show
            Me.Hide
        Case "nParts (Material) Supply [Unscheduled]":
            If hakAkses("frm_part_supply") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frm_part_supply.Show
            Me.Hide
        Case "nParts (Material) Supply [By BOM]":
            If hakAkses("FrmPartMaterialSupply_BOM") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            'FrmPartMaterialSupply_BOM.Show
            frmPartMaterialSupplyByBom.Show
            Me.Hide
        Case "nParts (Material) Receipt / Supply Inquiry":
            If hakAkses("frmPartsRecSupInquiry") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmPartsRecSupInquiry.Show
            Me.Hide
        Case "nReceipt Supply Schedule Inquiry":
            If hakAkses("frm_ReceiptSupplyScheculeInquiry") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frm_ReceiptSupplyScheculeInquiry.Show
            Me.Hide
        Case "nStock Inquiry (Item Code)":
            If hakAkses("F_StockInquiry") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            F_StockInquiry.Show
            Me.Hide
        Case "nStock Inquiry (Location)":
            If hakAkses("F_StockLocation") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            F_StockLocation.Show
            Me.Hide
        Case "nPhysical Inventory Update":
            If hakAkses("frm_pi_update") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frm_pi_update.Show
            Me.Hide
        Case "nPhysical Inventory List":
            If hakAkses("frm_pi_list") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frm_pi_list.Show
            Me.Hide
        Case "nPhysical Inventory List (Detail)":
            If hakAkses("frm_pi_list_Detail") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frm_pi_list_Detail.Show
            Me.Hide
        Case "nInventory Report":
            If hakAkses("frm_pi_report") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frm_pi_report.Show
            Me.Hide
        Case "nInventory Stock Closing":
            If hakAkses("frminventoryclosing") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmInventoryClosing.Show
            Me.Hide
        Case "nParts (Material) Alarm List":
            If hakAkses("FrmAlarmList") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmAlarmList.Show
            Me.Hide
        Case "nParts (Material) Supply List Report":
            If hakAkses("FrmRptSupplyList") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmRptSupplyList.Show
            Me.Hide
        Case "nStock Transfer":
            If hakAkses("FrmStockTransfer") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmStockTransfer.Show
            Me.Hide
        
        Case "nFinished Goods Stock Report":
            If hakAkses("FrmSerialDetailList") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmSerialDetailList.Show
            Me.Hide
            
        Case "nPart (Material) Receipt Scheduled Upload":
            If hakAkses("FrmPart_RecUpload") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmPart_RecUpload.Show
            Me.Hide
            
         Case "nParts (Material) Receipt Status":
            If hakAkses("FrmPart_RecStatus") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmPart_RecStatus.Show
            Me.Hide
                    
        '**************** Production ****************
        Case "nDaily Production Schedule Entry (Update)":
            If hakAkses("frmDailyProdEntry") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmDailyProdEntry.Show
            Me.Hide
        Case "nProduction Schedule/Result Inquiry":
            If hakAkses("frmProdResultInquiry") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmProdResultInquiry.Show
            Me.Hide
'        Case "nProduction Result":
'            If hakAkses("frmProdResult") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
'            frmProdResult.Show
'            frmProdScanBarcode.Show vbModal
'            Me.Hide
        Case "nDaily Production Schedule Status":
            If hakAkses("frmDailyProdStatus") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmDailyProdStatus.Show
            Me.Hide
        Case "nProduction Result Report":
            If hakAkses("frmRptProdResult") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmRptProdResult.Show
            Me.Hide
        Case "nDaily Production Schedule Report":
            If hakAkses("frmRptDailySchedule") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmRptDailySchedule.Show
            Me.Hide
        Case "nProduction Schedule Calculation":
            If hakAkses("frmProductionScheduleCalculation") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmProductionScheduleCalculation.Show
            Me.Hide
        Case "nProduction Schedule Calculation Detail":
            If hakAkses("frmProductionScheduleCalculationDetail") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmProductionScheduleCalculationDetail.Show
            Me.Hide
        Case "nLot Traceability Inquiry":
            If hakAkses("frmLotTraceInquiry") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmLotTraceInquiry.Show
            Me.Hide
        
          Case "nReceipt By Serial No":
            If hakAkses("FrmReceiptBySerialNo") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmReceiptBySerialNo.Show
            Me.Hide
            
            Case "nSupply By Serial No":
            If hakAkses("FrmSupplyBySerialNo") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmSupplyBySerialNo.Show
            Me.Hide
            
        '**************** Efficiency Control ****************
        Case "nDefective Type Pareto Diagram":
            If hakAkses("frmEffBadTypeControl") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmEffBadTypeControl.Show
            Me.Hide
        Case "nDefective Type By Material Pareto Diagram":
            If hakAkses("frmEffBadByMaterialControl") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmEffBadByMaterialControl.Show
            Me.Hide
        Case "nDefective Material Pareto Diagram":
            If hakAkses("frmEffBadMaterialControl") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmEffBadMaterialControl.Show
            Me.Hide
        Case "nProduction Schedule/Result Difference Control":
            If hakAkses("frmEffProdResultDiffControl") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmEffProdResultDiffControl.Show
            Me.Hide
        Case "nProduction Schedule/Result Control":
            If hakAkses("frmEffProdResultControl") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmEffProdResultControl.Show
            Me.Hide
        Case "nWorking Loss Time Control":
            If hakAkses("frmEffWorkingLossControl") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmEffWorkingLossControl.Show
            Me.Hide
        Case "nEfficiency Control":
            If hakAkses("frmEffProdResultEfficiencyControl") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmEffProdResultEfficiencyControl.Show
            Me.Hide
        Case "nMaterial Consumption Report":
            If hakAkses("frmEffMaterialConsumptionRpt") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmEffMaterialConsumptionRpt.Show
            Me.Hide
            
        '**************** Structure System ****************
        Case "nUser Setup":
            If StatusAdmin = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmUserSetup.Show
            Me.Hide
        Case "nApproval Sign Code":
            If StatusAdmin = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmApprove.Show
            Me.Hide
        Case "nCompany Profile":
            If StatusAdmin = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmCompanyProfile.Show
            Me.Hide
            
        Case "nLog User History":
            If StatusAdmin = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frm_log.Show
            Me.Hide
        
        '**************** MRP ****************
        Case "nMRP Calculation":
            If hakAkses("frmMRPCalculation") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmMRPCalculation.Show
            Me.Hide
        Case "nMRP Inquiry":
            If hakAkses("frmMRPInquiry") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmMRPInquiry.Show
            Me.Hide
        Case "nMRP Setting":
            If hakAkses("frmMRPSetting") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmMRPSetting.Show
            Me.Hide
            
            
        
        '**************** Database ****************
        Case "nDatabase Backup":
            If hakAkses("FrmDatabaseBackup") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmDatabaseBackup.Show
            Me.Hide
            
        '**************** Inventory Valuation ****************
'       Case "nCost Classification Master":
'            If hakAkses("FrmCostMaster") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
'            FrmCostMaster.Show
'            Me.Hide
        Case "nInterface Inventory Valuation"
            If hakAkses("FrmInterfaceAmount") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values ('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmInterfaceAmount.Show
            Me.Hide
        Case "nProcess Master":
            If hakAkses("FrmProcessMaster") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmProcessMaster.Show
            Me.Hide
        Case "nCost/Minute Master":
                    If hakAkses("frmcostminutemaster") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
                    
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
                    
                    frmcostminutemaster.Show
                    Me.Hide


        Case "nMaterial Consumption Report Detail":
            If hakAkses("FrmMaterialConsumptionReportDetail") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmMaterialConsumptionReportDetail.Show
            Me.Hide
'----
        Case "nValuation Price Calculation":
            If hakAkses("FrmValuationPriceCalculation") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmValuationPriceCalculation.Show
            Me.Hide
        Case "nValuation Price Setup":
            If hakAkses("FrmValuationPriceSetup") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmValuationPriceSetup.Show
            Me.Hide
'        Case "nValuation Price Report Detail":
'            If hakAkses("FrmValuationPriceReportDetail") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
'            FrmValuationPriceReportDetail.Show
'            Me.Hide
'        Case "nValuation Price Report":
'            If hakAkses("FrmValuationPriceReport") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
'            FrmValuationPriceReport.Show
'            Me.Hide
        Case "nValuation Price Report Per Warehouse":
            If hakAkses("FrmValuationPriceReportWH") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmValuationPriceReportWH.Show
            Me.Hide
            
        Case "nValuation Price Report Detail":
            If hakAkses("FrmValuationPriceReportDetail") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmValuationPriceReportDetail.Show
            Me.Hide
            
        Case "nValuation Price Adjust":
            If hakAkses("FrmValuationPriceAdj") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmValuationPriceAdj.Show
            Me.Hide


        Case "nInventory Interface":
            If hakAkses("FrmInterfaceInv") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmInterfaceInv.Show
            Me.Hide
        Case "nGoods Issue Interface":
            If hakAkses("FrmInterfaceGI") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmInterfaceGI.Show
            Me.Hide
        Case "nI/F SAP - Invoice Serial No":
            If hakAkses("FrmInterface_InvoiceSerial") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmInterface_InvoiceSerial.Show
            Me.Hide
'interface AR / AP

        Case "nInterface Link Setup":
            If hakAkses("FrmInterfaceLink") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmInterfaceLink.Show
            Me.Hide
            
        Case "nAP Interface":
            If hakAkses("FrmInterfaceAP") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmInterfaceAP.Show
            Me.Hide
            
        Case "nAR Interface":
            If hakAkses("FrmInterfaceAR") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmInterfaceAR.Show
            Me.Hide
            
        Case "nInterface Inventory Valuation":
            If hakAkses("FrmInterfaceAmount") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmInterfaceAmount.Show
            Me.Hide
            
        '**************** Bom Cost ****************
        Case "nBOM Cost Calculation By Period":
            If hakAkses("FrmBomCostCalculation") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmBomCostCalculation.Show
            Me.Hide
            
        Case "nBOM Cost Report By Period":
            If hakAkses("FrmBomCostReport") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmBomCostReport.Show
            Me.Hide
            
         Case "nBOM Receipt Inquiry":
            If hakAkses("FrmBOMReceiptInquiry") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmBOMReceiptInquiry.Show
            Me.Hide
            
        '**************** Ceisa ****************
        Case "nBC 23 List":
            If hakAkses("FrmBC23List") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmBC23List.Show
            Me.Hide
            
        Case "nBC 25 List":
            If hakAkses("FrmBC25List") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmBC25List.Show
            Me.Hide
            
        Case "nBC 27 List":
           If hakAkses("FrmBC27List") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            frmBC27List.Show
            Me.Hide
            
        Case "nBC 40 List":
           If hakAkses("FrmBC40List") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmBC40List.Show
            Me.Hide
           
        Case "nBC 41 List":
           If hakAkses("FrmBC41List") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
        
            FrmBC41List.Show
            Me.Hide
            
         Case "nConnection To Mysql":
           If hakAkses("FrmBCConnection") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
        
            FrmBCConnection.Show
            Me.Hide
            
        '**************** BOM Master Upload ****************
        Case "nBOM Master Upload":
            If hakAkses("FrmBomMasterUpload") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmBomMasterUpload.Show
            Me.Hide
            
        Case "nPart Supply Upload":
            If hakAkses("FrmPartSupplyUpload") = 0 Then Call NoAkses: frmMainMenu.Show: Exit Sub
            
            sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(userLogin) & "','" & Right(Tree.selectedItem.Key, X) & "',Getdate()) "
            Db.Execute sql
            
            FrmPartSupplyUpload.Show
            Me.Hide
              
        Case Else
            lblErrMsg.Caption = "[0000] Function is not available !"
    End Select
    DoEvents
End Sub

Private Sub tree_BeforeLabelEdit(Cancel As Integer)
Cancel = 1
End Sub


Private Sub tree_Collapse(ByVal Node As MSComctlLib.Node)
bolshow = False
Tree.Nodes(1).Image = 1
Dim mynode As Node
If UCase(Node.Key) = "EZ" Then
   For Each mynode In Tree.Nodes
      mynode.Expanded = False
   Next
End If
End Sub

Private Sub Tree_Expand(ByVal Node As MSComctlLib.Node)
Dim mynode As Node
bolshow = False
If UCase(Node.Key) = "EZ" Then
   For Each mynode In Tree.Nodes
      If mynode.Key <> Node.Key Then mynode.Expanded = True
   Next
   Tree.Nodes(1).EnsureVisible
End If
End Sub

Sub loadtree()
Dim rsParent As Recordset, rsChild As Recordset
Dim cdparent As String, nmparent As String
Dim cdchild As String, nmchild As String
Dim cnode As Node
Dim ShowMenu As Boolean

sql = "select distinct group_indeks, group_id from user_menu where app_id ='P01' order by group_indeks"
Set rsParent = New Recordset
rsParent.Open sql, Db, adOpenDynamic, adLockOptimistic

Tree.Nodes.clear
Set ParentCode = Tree.Nodes.Add(, , "EZ", "EZ Runner - Main Menu", 1)
ParentCode.ExpandedImage = 2
ParentCode.Expanded = True

Do While Not rsParent.EOF

    ShowMenu = True
    cdparent = CStr(rsParent!group_indeks)
    nmparent = Trim(rsParent!group_Id)
    
'    If cdparent = 11 Then ShowMenu = (StatusAdmin = 1)
    If ShowMenu Then
    
        Set ParentCode = Tree.Nodes.Add("EZ", tvwChild, nmparent, nmparent, 1)
        ParentCode.ExpandedImage = 2
        
        sql = "Select a.* " & _
            "from user_menu a " & _
            "where a.group_indeks = '" & cdparent & "' and a.app_id ='P01' " & _
            "and IsNull((select status from User_Privilege where Menu_ID = a.Menu_ID And UserName = '" & userLogin & "'), 1) = 1 " & _
            "order by menu_indeks"
        
        Set rsChild = New Recordset
        rsChild.Open sql, Db, adOpenDynamic, adLockOptimistic
        Do While Not rsChild.EOF
            cdchild = rsChild!menu_indeks
            nmchild = Trim(rsChild!menu_desc)
            Set cnode = Tree.Nodes.Add(nmparent, tvwChild, "n" & nmchild, nmchild, 3)
            cnode.Expanded = True
            rsChild.MoveNext
        Loop
        If Tree.Nodes(nmparent).Children = 0 Then Tree.Nodes.Remove (nmparent)
    
    End If
    rsParent.MoveNext
Loop

End Sub

Sub treedoubleclick()
Call Tree_DblClick
End Sub



Private Sub tree_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then bolshow = True: treedoubleclick
End Sub

