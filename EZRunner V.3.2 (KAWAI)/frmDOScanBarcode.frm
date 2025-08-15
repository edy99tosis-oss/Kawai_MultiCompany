VERSION 5.00
Begin VB.Form frmDOScanBarcode 
   BackColor       =   &H00FDDFE3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scan Barcode"
   ClientHeight    =   2040
   ClientLeft      =   3645
   ClientTop       =   6705
   ClientWidth     =   6885
   Icon            =   "frmDOScanBarcode.frx":0000
   KeyPreview      =   -1  'True
   LinkMode        =   1  'Source
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   6885
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameMsg 
      BackColor       =   &H00FDDFE3&
      Height          =   600
      Left            =   225
      TabIndex        =   2
      Top             =   1215
      Width           =   6435
      Begin VB.Label lblErrMsg 
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
         Left            =   105
         TabIndex        =   3
         Top             =   195
         Width           =   6210
      End
   End
   Begin VB.TextBox txtBarcode 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   225
      MaxLength       =   25
      TabIndex        =   1
      Top             =   720
      Width           =   6435
   End
   Begin VB.Label lblJudul 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Scan Barcode"
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
      Left            =   2685
      TabIndex        =   0
      Top             =   135
      Width           =   1515
   End
End
Attribute VB_Name = "frmDOScanBarcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ShowData()
Dim strSerial As String
Dim StrItem As String
Dim strSQL As String

strSerial = Right(Trim(txtBarcode), 7)
StrItem = Left(Trim(txtBarcode), Len(Trim(txtBarcode)) - 8)

Call frmDODetail.CheckSerial(StrItem, strSerial)

End Sub


Private Sub Form_Unload(Cancel As Integer)
FrmProdResultDetail.Enabled = True
End Sub

Private Sub txtBarcode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyReturn: ShowData
    Case vbKeyEscape: Unload Me
    End Select
End Sub

