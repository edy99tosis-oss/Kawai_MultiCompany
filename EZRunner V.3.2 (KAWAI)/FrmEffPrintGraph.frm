VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmEffPrintGraph 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Printer Setup"
   ClientHeight    =   3330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5805
   Icon            =   "FrmEffPrintGraph.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3330
   ScaleWidth      =   5805
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H0080FFFF&
      Caption         =   "OK"
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
      Left            =   4590
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2670
      Width           =   915
   End
   Begin VB.OptionButton optPort 
      BackColor       =   &H00FDDFE3&
      Caption         =   "Portrait"
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
      Left            =   330
      TabIndex        =   11
      Top             =   1920
      Width           =   945
   End
   Begin VB.OptionButton optLand 
      BackColor       =   &H00FDDFE3&
      Caption         =   "Landscape"
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
      Left            =   1290
      TabIndex        =   10
      Top             =   1950
      Width           =   1215
   End
   Begin VB.TextBox txtCopies 
      Height          =   315
      Left            =   4470
      MaxLength       =   2
      TabIndex        =   2
      Text            =   "1"
      Top             =   1950
      Width           =   720
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FDDFE3&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   330
      Picture         =   "FrmEffPrintGraph.frx":0E42
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   2550
      Width           =   480
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H0080FFFF&
      Caption         =   "Cancel"
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
      Left            =   3570
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2670
      Width           =   945
   End
   Begin MSForms.ComboBox cboPrinters 
      Height          =   375
      Left            =   1620
      TabIndex        =   12
      Top             =   630
      Width           =   3615
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "6376;661"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FDDFE3&
      Caption         =   "Orientation"
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
      Index           =   7
      Left            =   390
      TabIndex        =   9
      Top             =   1620
      Width           =   945
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FDDFE3&
      Caption         =   "Copies"
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
      Left            =   2700
      TabIndex        =   7
      Top             =   1620
      Width           =   630
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FDDFE3&
      Caption         =   "Printer"
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
      Index           =   6
      Left            =   420
      TabIndex        =   3
      Top             =   180
      Width           =   570
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FDDFE3&
      Caption         =   "Number of copies:"
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
      Index           =   0
      Left            =   2790
      TabIndex        =   8
      Top             =   2010
      Width           =   1605
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000010&
      Height          =   765
      Index           =   0
      Left            =   270
      Top             =   1710
      Width           =   2265
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FDDFE3&
      Caption         =   "Status:"
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
      Left            =   660
      TabIndex        =   6
      Top             =   1170
      Width           =   615
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H00FDDFE3&
      Caption         =   "Default printer; Ready"
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
      Left            =   1620
      TabIndex        =   5
      Top             =   1170
      Width           =   3720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FDDFE3&
      Caption         =   "Name:"
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
      Left            =   660
      TabIndex        =   4
      Top             =   690
      Width           =   570
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000000&
      Height          =   1305
      Index           =   1
      Left            =   270
      Top             =   270
      Width           =   5235
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000000&
      Height          =   735
      Index           =   1
      Left            =   2625
      Top             =   1725
      Width           =   2880
   End
End
Attribute VB_Name = "frmEffPrintGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DefPrinter As String   ' Default printer, SQTSQL AS STRING
Public Orient As Integer

Private Sub ListPrinters()
    ' Show printers list
    Dim i As Integer
    cboPrinters.ColumnCount = 3
    cboPrinters.ColumnWidths = "150 pt;0 pt; 0 pt"
    For i = 0 To Printers.Count - 1
        cboPrinters.AddItem ""
        cboPrinters.List(i, 0) = Printers(i).DeviceName
        cboPrinters.List(i, 1) = Printer.Port
        cboPrinters.List(i, 2) = Printer.DriverName
        If Printers(i).DeviceName = Printer.DeviceName Then
            cboPrinters.Text = Printer.DeviceName
        End If
    Next i
    DefPrinter = Printer.DeviceName
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
    ListPrinters
    If Orient = 2 Then
        Me.optLand.Value = True
    Else
        Me.optPort.Value = True
    End If
End Sub

Private Sub cboPrinters_Click()
    ' Selects a printer
    Dim Prt As Printer
    For Each Prt In Printers
        If Prt.DeviceName = cboPrinters.Text Then
            Set Printer = Prt
            Exit For
        End If
    Next
    
    If cboPrinters.Text = DefPrinter Then
        LblStatus.Caption = "Default printer; Ready"
    Else
        LblStatus.Caption = "Ready"
    End If
End Sub

Private Sub cmdOK_Click()
On Error GoTo HandleErr

    Printer.PaperSize = vbPRPSA4
    Printer.Orientation = Orient
    
    If InStr(1, Printer.DeviceName, "Microsoft Office Document Image Writer") <= 0 Then
        Printer.Copies = CInt(txtCopies)
    End If
    StPrint = True
    
HandleErr:
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cboPrinters_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = 0
End Sub

Private Sub optLand_Click()
    Orient = 2
End Sub

Private Sub optPort_Click()
    Orient = 1
End Sub

Private Sub txtCopies_GotFocus()
    txtCopies.SelStart = 0
    txtCopies.SelLength = Len(txtCopies)
End Sub

Private Sub txtCopies_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub


