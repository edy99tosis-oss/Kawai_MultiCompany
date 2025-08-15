VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form FrmRpt3 
   Caption         =   "Preview"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   4680
   Icon            =   "FrmRpt3.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3150
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer1 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      lastProp        =   500
      _cx             =   7435
      _cy             =   5530
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
End
Attribute VB_Name = "FrmRpt3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CRViewer1_PrintButtonClicked(UseDefault As Boolean)
If reportcode <> "SupplyList" Then
    UseDefault = False
    If TutupPtr = False Then
        'FrmPrint.Show 1, Me
        Frm_Print.Show 1, Me

    Else
        MsgBox "Data can't be printed "
    End If
End If
End Sub

Private Sub Form_Load()
TutupPtr = False
End Sub

Private Sub Form_Resize()
On Error Resume Next
CRViewer1.Left = 10
CRViewer1.Height = ScaleHeight
CRViewer1.Width = Me.Width - 100
End Sub



