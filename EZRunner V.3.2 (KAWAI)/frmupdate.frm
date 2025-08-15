VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmupdate 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Update Surat Jalan"
   ClientHeight    =   2850
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5550
   Icon            =   "frmupdate.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2850
   ScaleWidth      =   5550
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdMenu 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Back"
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
      Left            =   300
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2400
      Width           =   1185
   End
   Begin VB.CommandButton CmdSubmit 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Update"
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
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2400
      Width           =   1185
   End
   Begin VB.TextBox lblsjdate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
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
      Left            =   3720
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1620
      Width           =   1695
   End
   Begin VB.TextBox lblsjno 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
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
      Left            =   3720
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox txtsjno 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
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
      Left            =   1980
      MaxLength       =   25
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox txtpo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
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
      Left            =   1980
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   780
      Width           =   1695
   End
   Begin VB.TextBox txtsupplier 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
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
      Left            =   1980
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   360
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker sjdate 
      Height          =   345
      Left            =   1980
      TabIndex        =   9
      Top             =   1620
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   609
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
      Format          =   174718979
      CurrentDate     =   37868
   End
   Begin VB.Label LblPart 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Receipt Date"
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
      Index           =   4
      Left            =   300
      TabIndex        =   8
      Top             =   1695
      Width           =   1095
   End
   Begin VB.Label LblPart 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Surat Jalan No"
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
      Left            =   300
      TabIndex        =   7
      Top             =   1260
      Width           =   1245
   End
   Begin VB.Label LblPart 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PO No"
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
      TabIndex        =   3
      Top             =   840
      Width           =   525
   End
   Begin VB.Label LblPart 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier Code "
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
      TabIndex        =   0
      Top             =   420
      Width           =   1275
   End
End
Attribute VB_Name = "frmupdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdMenu_Click()
Unload Me
End Sub

Private Sub CmdSubmit_Click()
Dim strSQL As String
strSQL = "update Part_Receipt set SuratJalan_No='" & Trim(txtsjno.Text) & "',Receipt_Date='" & Format(sjdate.Value, "yyyy-mm-dd") & "' where supplier_Code='" & Trim(txtsupplier.Text) & "'" & vbCrLf & _
          "and PO_NO='" & Trim(txtpo.Text) & "' and suratjalan_no='" & Trim(lblSJNo.Text) & "'"
          
Db.Execute strSQL

Call FrmPart_Rec.display
Unload Me


End Sub
