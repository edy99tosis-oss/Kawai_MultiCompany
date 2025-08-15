VERSION 5.00
Begin VB.Form frm_pi_barcode 
   BackColor       =   &H00FDDFE3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scan Barcode"
   ClientHeight    =   2040
   ClientLeft      =   3645
   ClientTop       =   6705
   ClientWidth     =   6885
   Icon            =   "frm_pi_barcode.frx":0000
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
Attribute VB_Name = "frm_pi_barcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ShowData()
    Dim adoRs As New ADODB.Recordset
    
    LblErrMsg.Caption = ""
        
    sql = "Select * From Stock_Master Where WareHouse_Code='" & Left(txtBarcode, 6) & "' And " & _
            " Item_Code='" & Mid(txtBarcode, 7, 15) & "'"
        
    adoRs.Open sql, Db, adOpenStatic, adLockReadOnly, adCmdText
    
    If Not adoRs.EOF Then
        frm_pi_barcode.Hide
        If Trim(frm_pi_update.CboLocationCD) = Trim(adoRs(0)) Then
            If frm_pi_update.grid.Rows <= 1 Then
                frm_pi_update.CboLocationCD = Trim(adoRs(0))
                Call frm_pi_update.qBrowse
            End If
         Else
                frm_pi_update.CboLocationCD = Trim(adoRs(0))
                Call frm_pi_update.qBrowse
        End If
        frm_pi_update.txtItemCode = ""
        frm_pi_update.txtItemCode = Trim(adoRs(1))
        Call frm_pi_update.SetPosisi
        Unload frm_pi_barcode
    Else
        LblErrMsg.Caption = DisplayMsg("0013")
        txtBarcode.Text = ""
    End If
End Sub

Private Sub txtBarcode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyReturn: ShowData
    Case vbKeyEscape: Unload Me
    End Select
End Sub
