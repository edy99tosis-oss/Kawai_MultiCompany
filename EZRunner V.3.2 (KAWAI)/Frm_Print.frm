VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frm_Print 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Printer Setup"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5805
   Icon            =   "Frm_Print.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   5805
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDDFE3&
      ForeColor       =   &H80000008&
      Height          =   1065
      Left            =   270
      TabIndex        =   21
      Top             =   3210
      Width           =   3195
      Begin VB.TextBox txtRange 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Index           =   1
         Left            =   2460
         MaxLength       =   3
         TabIndex        =   25
         Top             =   570
         Width           =   540
      End
      Begin VB.OptionButton optRange 
         BackColor       =   &H00FDDFE3&
         Caption         =   "&Page(s)"
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
         Index           =   1
         Left            =   240
         TabIndex        =   24
         Top             =   600
         Width           =   1185
      End
      Begin VB.OptionButton optRange 
         BackColor       =   &H00FDDFE3&
         Caption         =   "&All"
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
         Left            =   240
         TabIndex        =   23
         Top             =   300
         Value           =   -1  'True
         Width           =   1665
      End
      Begin VB.TextBox txtRange 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Index           =   0
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   22
         Top             =   570
         Width           =   540
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FDDFE3&
         Caption         =   "To"
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
         Index           =   9
         Left            =   2130
         TabIndex        =   26
         Top             =   600
         Width           =   285
      End
   End
   Begin VB.ComboBox Combo1 
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
      Left            =   1590
      TabIndex        =   15
      Text            =   "Combo1"
      Top             =   1560
      Width           =   1935
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
      TabIndex        =   14
      Top             =   2610
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
      TabIndex        =   13
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox txtCopies 
      Height          =   315
      Left            =   4470
      TabIndex        =   4
      Text            =   "1"
      Top             =   2640
      Width           =   720
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   270
      Left            =   4965
      Max             =   -1
      Min             =   -9
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2670
      Value           =   -1
      Width           =   195
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FDDFE3&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   5010
      Picture         =   "Frm_Print.frx":0E42
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   3750
      Width           =   480
   End
   Begin VB.CommandButton Command1 
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
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4380
      Width           =   945
   End
   Begin VB.CommandButton cmok 
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
      TabIndex        =   0
      Top             =   4380
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FDDFE3&
      Caption         =   "Source:"
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
      Index           =   5
      Left            =   660
      TabIndex        =   20
      Top             =   1620
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FDDFE3&
      Caption         =   "Where:"
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
      Index           =   3
      Left            =   3750
      TabIndex        =   19
      Top             =   1860
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FDDFE3&
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4440
      TabIndex        =   18
      Top             =   1110
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSForms.ComboBox cbxPrinters 
      Height          =   375
      Left            =   1620
      TabIndex        =   17
      Top             =   630
      Width           =   3615
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "6376;661"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Verdana"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FDDFE3&
      Caption         =   "Label2"
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
      Left            =   4440
      TabIndex        =   16
      Top             =   1350
      Visible         =   0   'False
      Width           =   795
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
      TabIndex        =   12
      Top             =   2310
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
      Left            =   2760
      TabIndex        =   10
      Top             =   2310
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
      TabIndex        =   5
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
      TabIndex        =   11
      Top             =   2700
      Width           =   1605
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000000&
      Height          =   735
      Index           =   1
      Left            =   2625
      Top             =   2415
      Width           =   2880
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000010&
      Height          =   765
      Index           =   0
      Left            =   270
      Top             =   2400
      Width           =   2265
   End
   Begin VB.Label lblPort 
      AutoSize        =   -1  'True
      BackColor       =   &H00FDDFE3&
      Caption         =   "Label4"
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
      Left            =   4710
      TabIndex        =   9
      Top             =   1830
      Visible         =   0   'False
      Width           =   555
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
      TabIndex        =   8
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
      TabIndex        =   7
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
      TabIndex        =   6
      Top             =   690
      Width           =   570
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000000&
      Height          =   1965
      Index           =   1
      Left            =   285
      Top             =   270
      Width           =   5235
   End
End
Attribute VB_Name = "Frm_Print"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DC_BINS = 6
Private Const DC_BINNAMES = 12

Private Declare Function DeviceCapabilities Lib "winspool.drv" _
Alias "DeviceCapabilitiesA" (ByVal lpDeviceName As String, _
ByVal lpPort As String, ByVal iIndex As Long, lpOutput As Any, _
ByVal dev As Long) As Long

Dim DefPrinter As String   ' Default printer, SQTSQL AS STRING
Dim Orient As Integer, strSQL As String
Dim BolClosed As Boolean

Sub Printout(rptX As CRAXDDRT.report, xcopy As Integer, xcollated As Boolean, startPage As Long, EndPage As Long)
Dim xnumber As Long
If optRange(0).Value = True Then
    rptX.Printout False, xcopy, True
    BolClosed = True
Else
    xnumber = rptX.PrintingStatus.NumberOfPages
    If Val(txtRange(0)) > xnumber Then
        MsgBox "Invalid Page!, Number of Page is " & xnumber & " ", vbCritical + vbSystemModal, "Error!"
        txtRange(0).SetFocus
        BolClosed = False
    ElseIf Val(txtRange(1)) > xnumber Then
        MsgBox "Invalid Page!, Number of Page is " & xnumber & " ", vbCritical + vbSystemModal, "Error!"
        txtRange(1).SetFocus
        BolClosed = False
    ElseIf Val(txtRange(0)) = 0 Or Val(txtRange(1)) = 0 Then
        MsgBox "Invalid Page!", vbCritical + vbSystemModal, "Error!"
        txtRange(1).SetFocus
        BolClosed = False
    Else
        rptX.Printout (False), xcopy, xcollated, startPage, EndPage
        BolClosed = True
    End If
End If
End Sub

Public Function GetBinNumbers() As Variant
'Code adapted from Microsoft KB article Q194789
'HOWTO: Determine Available PaperBins with DeviceCapabilities API
Dim iBins As Long
Dim iBinArray() As Integer
Dim sPort As String
Dim sCurrentPrinter As String
Dim activePrinter As String
'Get the printer & port name of the current printer
sPort = Label2 'lblPort
sCurrentPrinter = cbxPrinters.Text
'Find out how many printer bins there are
iBins = DeviceCapabilities(sCurrentPrinter, sPort, _
DC_BINS, ByVal vbNullString, 0)
'Set the array of bin numbers to the right size
ReDim iBinArray(0 To iBins - 1)
'Load the array with the bin numbers
iBins = DeviceCapabilities(sCurrentPrinter, sPort, _
DC_BINS, iBinArray(0), 0)
'Return the array to the calling routine
GetBinNumbers = iBinArray
End Function
Public Function GetBinNames() As Variant

'Code adapted from Microsoft KB article Q194789
'HOWTO: Determine Available PaperBins with DeviceCapabilities API
Dim iBins As Long
Dim ct As Long
Dim sNamesList As String
Dim sNextString As String
Dim sPort As String
Dim sCurrentPrinter As String
Dim activePrinter As String
Dim ActivePort As String
Dim vBins As Variant
activePrinter = cbxPrinters.Text
'Get the printer & port name of the current printer
sPort = Label2 'lblPort ''Trim$(Mid$(ActivePrinter, InStrRev(ActivePrinter, " ") + 1))
'sCurrentPrinter = Trim$(Left$(ActivePrinter, _
'InStr(ActivePrinter, " on ")))
sCurrentPrinter = cbxPrinters.Text
'Find out how many printer bins there are
iBins = DeviceCapabilities(sCurrentPrinter, sPort, _
DC_BINS, ByVal vbNullString, 0)
'Set the string to the right size to hold all the bin names
'24 chars per name
sNamesList = String(24 * iBins, 0)
'Load the string with the bin names
iBins = DeviceCapabilities(sCurrentPrinter, sPort, _
DC_BINNAMES, ByVal sNamesList, 0)
'Set the array of bin names to the right size
If iBins > 0 Then
ReDim vBins(0 To iBins - 1)

For ct = 0 To iBins - 1
'Get each bin name in turn and assign to the next item in the array
sNextString = Mid(sNamesList, 24 * ct + 1, 24)
vBins(ct) = Left(sNextString, InStr(1, sNextString, Chr(0)) - 1)
combo1.AddItem vBins(ct)
Next ct
End If
'Return the array to the calling routine
GetBinNames = vBins
End Function

Private Sub ListPrinters()
    ' Show printers list
    Dim i As Integer
    cbxPrinters.columnCount = 3
    cbxPrinters.ColumnWidths = "150 pt;0 pt; 0 pt"
    For i = 0 To Printers.Count - 1
        cbxPrinters.AddItem ""
        cbxPrinters.List(i, 0) = Printers(i).DeviceName
        cbxPrinters.List(i, 1) = Printer.Port
        cbxPrinters.List(i, 2) = Printer.DriverName
        If Printers(i).DeviceName = Printer.DeviceName Then
            cbxPrinters.Text = Printer.DeviceName
            Label2 = Printer.Port
            Label3 = Printer.DriverName
        End If
    Next i
    DefPrinter = Printer.DeviceName
End Sub

Private Sub cbxPrinters_Click()
    ' Selects a printer
    Dim Prt As Printer
    For Each Prt In Printers
        If Prt.DeviceName = cbxPrinters.Text Then
            Set Printer = Prt
            Exit For
        End If
    Next
    If cbxPrinters.Text = DefPrinter Then
        LblStatus.Caption = "Default printer; Ready"
    Else
        LblStatus.Caption = "Ready"
    End If
    combo1.clear
    GetBinNames
    'combo1.ListIndex = 0
   ' Label2 = cbxPrinters.Column(1)
    'Label3 = cbxPrinters.Column(2)
    
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cbxPrinters_KeyPress(KeyAscii As MSForms.ReturnInteger)
KeyAscii = 0
End Sub
Private Sub cmok_Click()
    If Val(txtCopies) <= 0 Or Val(txtCopies) > -VScroll1.Min Then
        MsgBox "Invalid number of copies", 16, " ¡Error!"
        txtCopies.SetFocus
        Exit Sub
    End If
    MousePointer = vbHourglass
    If reportcode = "DO" Then
        If pesaninvalid <> "" Then
            MsgBox "Cannot Print DO! " & pesaninvalid, vbCritical & vbOKOnly, "Information"
        Else
            DOPrint
            '## Update data ##
            strSQL = "Update do_master set reissue_cls= '1', last_update = getdate(), last_user = '" & userLogin & "' where do_no in (" & do_no & " )"
            Db.Execute strSQL
            TutupPtr = True
        End If
    ElseIf reportcode = "packinglist_Kwi" Then
        PackingListKwiPrint
            strSQL = "Update Packing_Master set reissue_cls= '1', last_update = getdate(), last_user = '" & userLogin & "' where packing_no in (" & packing_no & " )"
            Db.Execute strSQL
            TutupPtr = True
    ElseIf reportcode = "packing_list_ex" Then
        If pesaninvalid <> "" Then
            MsgBox "Cannot Print Packing List! " & pesaninvalid, vbCritical & vbOKOnly, "Information"
        Else
            PackingListPrint_Ex
            '## Update data ##
            strSQL = "Update Packing_Master set reissue_cls= '1', last_update = getdate(), last_user = '" & userLogin & "' where packing_no in (" & packing_no & " )"
            Db.Execute strSQL
            TutupPtr = True
        End If
    ElseIf reportcode = "invoice" Then
        If cantprint(inv_no) Then
            MsgBox "Cannot Print Invoice! Invoice No " & Trim(inv_no) & " is not fix!", vbCritical & vbOKOnly, "Information"
             MousePointer = vbDefault
            Exit Sub
        Else
            Invoice
            strSQL = "update invoice_master set reissue_cls ='1', last_update = getdate(), last_user = '" & userLogin & "' where invoice_no in (" & inv_no & ")"
            Db.Execute strSQL
            TutupPtr = True
        End If
    ElseIf reportcode = "invoice_kwi" Then
        InvoiceKWI
        strSQL = "update invoice_master set reissue_cls ='1', last_update = getdate(), last_user = '" & userLogin & "' where invoice_no in (" & inv_no & ")"
        Db.Execute strSQL
        TutupPtr = True
    ElseIf reportcode = "invoice_ex" Then
'        If cantprint(inv_no) Then
'            MsgBox "Cannot Print Invoice! Invoice No " & Trim(inv_no) & " is not fix!", vbCritical & vbOKOnly, "Information"
'             MousePointer = vbDefault
'            Exit Sub
'        Else
            InvoiceEx
            strSQL = "update invoice_master set reissue_cls ='1', last_update = getdate(), last_user = '" & userLogin & "' where invoice_no in (" & inv_no & ")"
            Db.Execute strSQL
            TutupPtr = True
'        End If
'    ElseIf reportcode = "Pajak" Then
'        Dim RsFP As Recordset
'        If FrmPajak_Create.NomPajak <> "" Then
'            Set RsFP = Db.Execute("Select * from fakturpajak_master where fakturpajak_no='" & FrmPajak_Create.NomPajak & "' and fix_cls='1'")
'
'            If RsFP.EOF Then
'                MsgBox "Cannot Print Faktur Pajak No. " & Trim(FrmPajak_Create.NomPajak) & " is not fix!", vbCritical & vbOKOnly, "Information"
'            Else
'                pajak
'            End If
'        Else
'            MsgBox "Cannot Print Faktur Pajak! ", vbCritical & vbOKOnly, "Information"
'        TutupPtr = False
'        End If
    ElseIf reportcode = "ForecastPart" Then
        ForecastPart
    ElseIf reportcode = "itemmaster" Then
        itemMaster
    ElseIf reportcode = "ForecastMaterial" Then
        ForecastMaterial
'    ElseIf reportcode = "6" Then
'        salesReportCust
    ElseIf reportcode = "ExchangeList" Then
        ExchangeList
    ElseIf reportcode = "pireport" Then
        piReport
    ElseIf reportcode = "pireportwh" Then
        piReportwh
    ElseIf reportcode = "8" Then
        WIP
    ElseIf reportcode = "9" Then
        RecSupInquiry
    ElseIf reportcode = "pilist" Then
        PiList
    ElseIf reportcode = "pilistdet" Then
        PiListDet
    ElseIf reportcode = "12" Then
        rawMaterial
    ElseIf reportcode = "Receiptsupplyschedule" Then
        RecSupScheduleInquiry
    ElseIf reportcode = "DI" Then
        If pesaninvalid <> "" Then
            MsgBox "Cannot Print Delivery Instruction! " & pesaninvalid, vbCritical & vbOKOnly, "Information"
        Else
            DI
            TutupPtr = True
        End If
    ElseIf reportcode = "materialConsumption" Then
            MaterialConsumption
    ElseIf reportcode = "forecast report" Then
            forecastreport
            TutupPtr = False
    ElseIf reportcode = "prodplanning" Then
        ProductionPlanning
        TutupPtr = False
    ElseIf reportcode = "InvoiceDetailList" Then
        InvoiceDetailList
        TutupPtr = False
    ElseIf reportcode = "InvoiceSummaryList" Then
        InvoiceSummaryList
        TutupPtr = False
    ElseIf reportcode = "Alarm" Then
        Alarmlist
    ElseIf reportcode = "AlarmOP" Then
        AlarmlistOP
    ElseIf reportcode = "trademaster" Then
        TradeMaster
    ElseIf reportcode = "Faktur Pajak" Then
        Dim RsFP As Recordset
        Set RsFP = Db.Execute("Select * from fakturpajak_master where fakturpajak_no='" & pajak_No & "' and fix_cls='1'")
        If RsFP.EOF Then
            MsgBox "Cannot Print Faktur Pajak No. " & pajak_No & " is not fix!", vbCritical & vbOKOnly, "Information"
        Else
            Pajak
        End If
    ElseIf reportcode = "AccPay" Then
        AccPay
    ElseIf reportcode = "BOM" Then
        BOM
    ElseIf reportcode = "rptFormula" Then
        rptFormula
    ElseIf reportcode = "rptProdResultInquiry" Then
        ProdResultInquiry
    ElseIf reportcode = "rptProdResultInquiry2" Then
        rptProdResultInquiry2
    ElseIf reportcode = "ProdResultByFactory" Then
        ProdResultByFactory
    ElseIf reportcode = "rptProdResultByItem" Then
        ProdResultByItem
    ElseIf reportcode = "pricemaster" Then
        pricemaster
    ElseIf reportcode = "custexchrate" Then
        custexchrate
    ElseIf reportcode = "dailyproduction" Then
        dailyproduction
    ElseIf reportcode = "salesreport" Then
        salesreport
    ElseIf reportcode = "pmsBOM" Then
        pmsBOM
    ElseIf reportcode = "bookkeepingexchrate" Then
        BookKeepingExchRate
    ElseIf reportcode = "PORequestPrint" Then
        RptPORequestPrint
    ElseIf reportcode = "polocal" Then
'        If cantprintpo(frmPOParts.txtpono.Text) Then
'            MsgBox "Cannot Print PO Parts! " & Trim(frmPOParts.txtpono.Text) & " is not fix!", vbCritical & vbOKOnly, "Information"
'        Else
'            POLocalPrint (frmPOParts.txtpono.Text)
'            TutupPtr = True
'        End If
    ElseIf reportcode = "poimport" Then
        If cantprintpo(frmPOParts.txtPoNo.Text) Then
            MsgBox "Cannot Print PO Parts! " & Trim(frmPOParts.txtPoNo.Text) & " is not fix!", vbCritical & vbOKOnly, "Information"
        Else
            POImportPrint (frmPOParts.txtPoNo.Text)
            TutupPtr = True
        End If
    ElseIf reportcode = "pocoil" Then
        If cantprintpo(frmPOSteelCoil.txtPoNo.Text) Then
            MsgBox "Cannot Print PO Steel Coil! " & Trim(frmPOSteelCoil.txtPoNo.Text) & " is not fix!", vbCritical & vbOKOnly, "Information"
        Else
            POCoilPrint (frmPOSteelCoil.txtPoNo.Text)
            TutupPtr = True
        End If
    ElseIf reportcode = "posubcon" Then
        If cantprintpo(frmPOParts.txtPoNo.Text) Then
            MsgBox "Cannot Print PO Parts! " & Trim(frmPOParts.txtPoNo.Text) & " is not fix!", vbCritical & vbOKOnly, "Information"
        Else
            POImportPrint (frmPOParts.txtPoNo.Text)
            TutupPtr = True
        End If
    ElseIf reportcode = "outstandingmat" Then
        OutstandingMaterial
    ElseIf reportcode = "resinrequest" Then
        ResinRequest
    ElseIf reportcode = "materialrequest" Then
        MaterialRequest
    ElseIf reportcode = "purchasingmat" Then
        PurchasingMaterial
    ElseIf reportcode = "Formula" Then
        Formula
    ElseIf reportcode = "pocust" Then
        pocust
    ElseIf reportcode = "rptrequestauto" Then
        RptRequestAuto
    ElseIf reportcode = "MatrixDaily" Then
        MatrixDaily
    ElseIf reportcode = "rptWorksheet" Then
        rptWorksheet
    ElseIf reportcode = "RptSupplyByBom" Then
        RptSupplyByBom
    ElseIf reportcode = "MatrixDaily22" Then
        MatrixDaily22
    ElseIf reportcode = "SupplyList" Then
        SupplyList
    ElseIf reportcode = "SupplyListValue" Then
        SupplyListValue
    ElseIf reportcode = "RptMaterialConsumptionReport" Then
        MaterialConsumptionReport
    ElseIf reportcode = "PurchasingPrice" Then
        PurchasingPrice
    ElseIf reportcode = "bc40" Then
        PrintBC40
    End If
    
    TutupPtr = False
    MousePointer = vbDefault
    Unload Me
End Sub
Private Sub RptSupplyByBom()
Dim application As New CRAXDDRT.application
Dim report As New CRAXDDRT.report
Dim rs1 As Recordset


    Set rs1 = New Recordset
    rs1.Open sqlprint, Db, adOpenKeyset, adLockOptimistic
    If Not rs1.EOF Then

        Set report = application.OpenReport(App.path & "\Reports\Rpt_supply_By_Bom.rpt")
        report.Database.Tables(1).SetDataSource rs1
        'report.FormulaFields(2).Text = "'" & ginvno & "'"
        'report.FormulaFields(3).Text = "'" & Fbulan & "'"
        'report.FormulaFields(4).Text = "'" & Ftahun & "'"
        'report.FormulaFields(6).Text = "'" & xbln & "'"
        '#####################################################################
        '# Qty Digit and decimal
        'report.FormulaFields(7).Text = gi_decimalDigitQty
        
        'report.FormulaFields(8).Text = gi_decimalDigitPrice
        
        'report.FormulaFields(9).Text = gi_decimalDigitAmount
        '#####################################################################
        Dim vBinNumbers As Variant
        If combo1.ListIndex >= 0 Then
        vBinNumbers = GetBinNumbers

        report.SelectPrinter Label3, cbxPrinters.Text, Label2
        report.PaperOrientation = Orient
        report.PaperSource = vBinNumbers(combo1.ListIndex)
        'report.PaperSize = crPaperA4
        Else
        MsgBox "No paper tray has been selected."
        End If

        Dim nCopy As Integer
        nCopy = Me.txtCopies
        Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
        Set report = Nothing
        rs1.Close
        Set rs1 = Nothing
    End If
End Sub
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
  If gb_Simulation = True Then Call up_InitSimulation(Me)
    ListPrinters
    If printorient = 2 Then
        Me.optLand.Value = True
    Else
        Me.optPort.Value = True
    End If
End Sub
Private Sub optLand_Click()
Orient = 2
End Sub

Private Sub optPort_Click()
Orient = 1
End Sub

Private Sub optRange_Click(Index As Integer)
    txtRange(0).Enabled = (optRange(1).Value)
    txtRange(1).Enabled = (optRange(1).Value)
End Sub

Private Sub txtCopies_GotFocus()
    txtCopies.SelStart = 0
    txtCopies.SelLength = Len(txtCopies)
End Sub

Private Sub txtCopies_KeyPress(KeyAscii As Integer)
If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub VScroll1_Change()
    ' Set max number of copies by changing Vscroll1.min
    txtCopies = -VScroll1.Value
End Sub

Sub MaterialConsumptionReport()
Dim application As New CRAXDDRT.application
Dim report As New CRAXDDRT.report
Dim rs1 As Recordset


    Set rs1 = New Recordset
    rs1.Open sqlprint, Db, adOpenKeyset, adLockOptimistic
    If Not rs1.EOF Then

        Set report = application.OpenReport(App.path & "\Reports\RptMaterialConsumptionReport.rpt")
        report.Database.Tables(1).SetDataSource rs1
        report.FormulaFields(2).Text = "'" & ginvno & "'"
        report.FormulaFields(3).Text = "'" & Fbulan & "'"
        report.FormulaFields(4).Text = "'" & Ftahun & "'"
        report.FormulaFields(6).Text = "'" & xbln & "'"
        '#####################################################################
        '# Qty Digit and decimal
        report.FormulaFields(7).Text = gi_decimalDigitQty
        report.FormulaFields(8).Text = gi_decimalDigitPrice
        report.FormulaFields(9).Text = gi_decimalDigitAmount
        '#####################################################################
        Dim vBinNumbers As Variant
        If combo1.ListIndex >= 0 Then
        vBinNumbers = GetBinNumbers

        report.SelectPrinter Label3, cbxPrinters.Text, Label2
        report.PaperOrientation = Orient
        report.PaperSource = vBinNumbers(combo1.ListIndex)
        'report.PaperSize = crPaperA4
        Else
        MsgBox "No paper tray has been selected."
        End If

        Dim nCopy As Integer
        nCopy = Me.txtCopies
        Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
        Set report = Nothing
        rs1.Close
        Set rs1 = Nothing
    End If
End Sub

Sub PurchasingPrice()
Dim application As New CRAXDDRT.application
Dim report As New CRAXDDRT.report
Dim rs1 As Recordset


    Set rs1 = New Recordset
    rs1.Open sqlprint, Db, adOpenKeyset, adLockOptimistic
    If Not rs1.EOF Then

        Set report = application.OpenReport(App.path & "\Reports\RptMaterialConsumptionReport.rpt")
        report.Database.Tables(1).SetDataSource rs1
        
        report.FormulaFields(1).Text = ginvno
        report.FormulaFields(2).Text = Fbulan
        report.FormulaFields(3).Text = Ftahun
        
        '#####################################################################
        '# Qty Digit and decimal
        report.FormulaFields(4).Text = gi_decimalDigitQty
        report.FormulaFields(5).Text = gi_decimalDigitPrice
        report.FormulaFields(6).Text = gi_decimalDigitAmount
        '#####################################################################
        Dim vBinNumbers As Variant
        If combo1.ListIndex >= 0 Then
        vBinNumbers = GetBinNumbers

        report.SelectPrinter Label3, cbxPrinters.Text, Label2
        report.PaperOrientation = Orient
        report.PaperSource = vBinNumbers(combo1.ListIndex)
        'report.PaperSize = crPaperA4
        Else
        MsgBox "No paper tray has been selected."
        End If

        Dim nCopy As Integer
        nCopy = Me.txtCopies
        Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
        Set report = Nothing
        rs1.Close
        Set rs1 = Nothing
    End If
End Sub
Sub DI()
Dim application As New CRAXDDRT.application
Dim report As New CRAXDDRT.report
Dim rs1 As Recordset


    Set rs1 = New Recordset
    rs1.Open sqlprint, Db, adOpenKeyset, adLockOptimistic
    If Not rs1.EOF Then

        Set report = application.OpenReport(App.path & "\Reports\Delivery_Instruction.rpt")
        report.Database.Tables(1).SetDataSource rs1

        '#####################################################################
        '# Qty Digit and decimal
         report.FormulaFields.GetItemByName("DecimalQty").Text = "" & gi_decimalDigitQty & ""
         report.FormulaFields.GetItemByName("DecimalBox").Text = "" & gi_decimalDigitBox & ""
        '#####################################################################
        Dim vBinNumbers As Variant
        If combo1.ListIndex >= 0 Then
        vBinNumbers = GetBinNumbers

        report.SelectPrinter Label3, cbxPrinters.Text, Label2
        report.PaperOrientation = Orient
        report.PaperSource = vBinNumbers(combo1.ListIndex)
        'report.PaperSize = crPaperA4
        Else
        MsgBox "No paper tray has been selected."
        End If

        Dim nCopy As Integer
        nCopy = Me.txtCopies
        Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
        Set report = Nothing
        rs1.Close
        Set rs1 = Nothing
    End If
End Sub

Sub DOPrint()

    Dim application As New CRAXDDRT.application
    Dim report As New CRAXDDRT.report
    Dim rs1 As Recordset
    
    Set rs1 = New Recordset
    rs1.Open sqlprint, Db, adOpenKeyset, adLockOptimistic
    If Not rs1.EOF Then

        Set report = application.OpenReport(App.path & "\Reports\delivery_order.rpt")
        report.Database.Tables(1).SetDataSource rs1
        
        report.FormulaFields.GetItemByName("DecimalBox").Text = "" & gi_decimalDigitBox & ""
        report.FormulaFields.GetItemByName("DecimalQty").Text = "" & gi_decimalDigitQty & ""
        
        Dim vBinNumbers As Variant
        If combo1.ListIndex >= 0 Then
            vBinNumbers = GetBinNumbers
            report.SelectPrinter Label3, cbxPrinters.Text, Label2
            report.PaperOrientation = Orient
            report.PaperSource = vBinNumbers(combo1.ListIndex)
        Else
            MsgBox "No paper tray has been selected."
        End If

        Dim nCopy As Integer
        nCopy = Me.txtCopies
        Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
        Set report = Nothing
        rs1.Close
        Set rs1 = Nothing
    End If

End Sub

Private Sub PackingListPrint()
    
    Dim application As New CRAXDDRT.application
    Dim report As New CRAXDDRT.report
    Dim rs1 As Recordset
    
    Set rs1 = New Recordset
    rs1.Open sqlprint, Db, adOpenKeyset, adLockOptimistic
    If Not rs1.EOF Then

        Set report = application.OpenReport(App.path & "\Reports\packing_list.rpt")
        report.Database.Tables(1).SetDataSource rs1
        
        report.FormulaFields(1).Text = "" & gi_decimalDigitQty & ""
        report.FormulaFields(2).Text = "" & gi_decimalDigitBox & ""
        report.FormulaFields(3).Text = "" & gi_decimalDigitWeight & ""
        
        Dim vBinNumbers As Variant
        If combo1.ListIndex >= 0 Then
            vBinNumbers = GetBinNumbers
            report.SelectPrinter Label3, cbxPrinters.Text, Label2
            report.PaperOrientation = Orient
            report.PaperSource = vBinNumbers(combo1.ListIndex)
        Else
            MsgBox "No paper tray has been selected."
        End If

        Dim nCopy As Integer
        nCopy = Me.txtCopies
        Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
        Set report = Nothing
        rs1.Close
        Set rs1 = Nothing
    End If
    
End Sub

Private Sub PackingListKwiPrint()
    Dim application As New CRAXDDRT.application
    Dim report As New CRAXDDRT.report
    Dim rsMain As New Recordset
    Dim rsSub As New Recordset
    Dim RsSubDetail As New Recordset
    
    Call Shipping(packing_no)
    
    If rsMain.State <> adStateClosed Then rsMain.Close
    rsMain.Open sqlprint, Db, adOpenKeyset, adLockOptimistic
    
    If Not rsMain.EOF Then

        Set report = application.OpenReport(App.path & "\Reports\packinglist_kwi.rpt")
        report.Database.Tables(1).SetDataSource rsMain
        report.DiscardSavedData
        
        Dim vBinNumbers As Variant
        If combo1.ListIndex >= 0 Then
            vBinNumbers = GetBinNumbers
            report.SelectPrinter Label3, cbxPrinters.Text, Label2
            report.PaperOrientation = Orient
            report.PaperSource = vBinNumbers(combo1.ListIndex)
        Else
            MsgBox "No paper tray has been selected."
        End If

        Dim nCopy As Integer
        nCopy = Me.txtCopies
        Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
        Set report = Nothing
        Set rsMain = Nothing
        Set rsSub = Nothing
      End If
      
End Sub


Private Sub Shipping(pc As String)
Dim application As New CRAXDDRT.application
Dim report As New CRAXDDRT.report
Dim rsMain As New Recordset
Dim rsSub As New Recordset
Dim RsSubDetail As New Recordset
Dim SQLr As String

SQLr = " Select PM.Packing_No,PM.Packing_Date,PM.Stuffing_Date,PM.ETD,PM.ETA,PM.Payment_Days,PM.Payment,PM.Country_Origin, " & vbCrLf & _
            "   PM.Transportation_Cls,PM.Vessel,PM.Mother_Vessel,PM.From_Port,PM.To_Port,PM.Final_Destination, " & vbCrLf & _
            "   PM.Remarks,PM.Last_User, " & vbCrLf & _
            "   PM.Cust_Code,TM.Trade_Name Customer_Name,TM.Address1 CsAddress1,TM.Address2 CsAddress2,TM.City CsCity, TM.Country CsCountry,TM.Contact_Person, " & vbCrLf & _
            "   PM.Consignee,PM.ConsigneeTitle,TM1.Trade_Name Consignee_Name,TM1.Address1 CgAddress1,TM1.Address2 CgAddress2,TM1.City CgCity, TM1.Country CgCountry, " & vbCrLf & _
            "   PM.Payment_Terms,Isnull(PT.Description,'') Payment_Desription, " & vbCrLf & _
            "   PD.Order_No,PD.Container_No,PD.Container_Size,PD.SerialNoFrom,PD.SerialNoTo,PD.Qty,PD.Unit_Cls, " & vbCrLf & _
            "   PD.QtyWeight_Netto,Pd.QtyWeight_Gross,PD.Ctn_No,PD.Qty_Ctn,PD.Length,PD.Width,PD.Thickness, " & vbCrLf & _
            "   PD.Item_Code,IM.Item_Name  " & vbCrLf & _
            "   From Packing_Master PM " & vbCrLf & _
            "   Inner Join Packing_Detail PD On PM.Packing_No=PD.Packing_No "

SQLr = SQLr + "   Inner Join Trade_Master TM on PM.Cust_Code=TM.Trade_Code " & vbCrLf & _
            "   Inner Join Trade_Master TM1 on PM.Consignee=TM1.Trade_Code " & vbCrLf & _
            "   Inner Join Item_Master IM on PD.Item_Code=IM.Item_Code " & vbCrLf & _
            "   Left Join PaymentTerm_Cls PT on PM.Payment_Terms=PT.PaymentTerm_Cls " & vbCrLf & _
            "   where PM.Packing_No=" & pc & vbCrLf & _
            "   Order By PM.Packing_No, PD.PackingSeq_No "
            
    If rsMain.State <> adStateClosed Then rsMain.Close
    rsMain.Open SQLr, Db, adOpenKeyset, adLockOptimistic
    
    If Not rsMain.EOF Then

        Set report = application.OpenReport(App.path & "\Reports\ShippingAdv_Kwi.rpt")
        report.Database.Tables(1).SetDataSource rsMain
        report.DiscardSavedData
        
        Dim vBinNumbers As Variant
        If combo1.ListIndex >= 0 Then
            vBinNumbers = GetBinNumbers
            report.SelectPrinter Label3, cbxPrinters.Text, Label2
            report.PaperOrientation = Orient
            report.PaperSource = vBinNumbers(combo1.ListIndex)
        Else
            MsgBox "No paper tray has been selected."
        End If

        Dim nCopy As Integer
        nCopy = Me.txtCopies
        Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
        Set report = Nothing
        Set rsMain = Nothing
        Set rsSub = Nothing
      End If

End Sub
Private Sub PackingListPrint_Ex()
    
    Dim application As New CRAXDDRT.application
    Dim report As New CRAXDDRT.report
    Dim rsMain As New Recordset
    Dim rsSub As New Recordset
    Dim RsSubDetail As New Recordset
    
    If rsMain.State <> adStateClosed Then rsMain.Close
    rsMain.Open sqlprint, Db, adOpenKeyset, adLockOptimistic
    
    If Not rsMain.EOF Then

        Set report = application.OpenReport(App.path & "\Reports\packing_list_ex.rpt")
        report.Database.Tables(1).SetDataSource rsMain
        report.FormulaFields.GetItemByName("DecimalQty").Text = gi_decimalDigitQty
        report.FormulaFields.GetItemByName("DecimalWeight").Text = gi_decimalDigitWeight
        
        If rsSub.State <> adStateClosed Then rsSub.Close
        rsSub.Open sqlprint2, Db, adOpenKeyset, adLockOptimistic
        With report.OpenSubreport("sub_packing_list_ex")
         .Database.Tables(1).SetDataSource rsSub
         .FormulaFields.GetItemByName("SubDecimalQty").Text = gi_decimalDigitQty
         .FormulaFields.GetItemByName("SubDecimalWeight").Text = gi_decimalDigitWeight
        End With
        
        If RsSubDetail.State <> adStateClosed Then RsSubDetail.Close
        RsSubDetail.Open sqlprint3, Db, adOpenKeyset, adLockOptimistic
    
        If Not RsSubDetail.EOF Then
         With report.OpenSubreport("subDetailOf_packing_list_ex")
          .Database.Tables(1).SetDataSource RsSubDetail
          .FormulaFields.GetItemByName("SubDetailDecimalQty").Text = gi_decimalDigitQty
          .FormulaFields.GetItemByName("SubDetailDecimalWeight").Text = gi_decimalDigitWeight
         End With
        End If

                
        Dim vBinNumbers As Variant
        If combo1.ListIndex >= 0 Then
            vBinNumbers = GetBinNumbers
            report.SelectPrinter Label3, cbxPrinters.Text, Label2
            report.PaperOrientation = Orient
            report.PaperSource = vBinNumbers(combo1.ListIndex)
        Else
            MsgBox "No paper tray has been selected."
        End If

        Dim nCopy As Integer
        nCopy = Me.txtCopies
        Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
        Set report = Nothing
        Set rsMain = Nothing
        Set rsSub = Nothing
      End If
    
End Sub
Sub ProductionPlanning()
Dim xrpt As Recordset
Dim application As New CRAXDDRT.application
Dim report As New CRAXDDRT.report
Dim i As Integer

Set xrpt = New Recordset
xrpt.Open sqlprint, Db, adOpenKeyset
Set report = application.OpenReport(App.path & "\REPORTs\Production_planning.rpt")
report.Database.Tables(1).SetDataSource xrpt
For i = 0 To 5
    If xthn(i) = Ftahun Then
        report.FormulaFields(i + 1).Text = "'" & MonthName(zbln(i), True) & "'"
    Else
        report.FormulaFields(i + 1).Text = "'" & MonthName(zbln(i), True) & "  " & xthn(i) & "'"
    End If
Next
report.FormulaFields(13).Text = "'Factory Code  :  " & F_Factory & "'"
report.FormulaFields(14).Text = "'Period            :  " & MonthName(Fbulan) & " " & Ftahun & "'"
'#####################################################################
'# Qty Digit and decimal
  report.FormulaFields(17).Text = "" & gi_decimalDigitQty & ""
  report.FormulaFields(18).Text = "" & gi_decimalDigitQty & ""
  report.FormulaFields(27).Text = "" & gi_decimalDigitPrice & "" 'Else
  report.FormulaFields(28).Text = "" & gi_decimalDigitPriceIDR & ""
  report.FormulaFields(29).Text = "" & gi_decimalDigitAmount & "" 'Else
  report.FormulaFields(30).Text = "" & gi_decimalDigitAmountIDR & ""
'#####################################################################

Dim vBinNumbers As Variant
If combo1.ListIndex >= 0 Then
vBinNumbers = GetBinNumbers
report.SelectPrinter Label3, cbxPrinters.Text, Label2

report.PaperSize = crPaperA4
report.PaperOrientation = Orient
report.PaperSource = vBinNumbers(combo1.ListIndex)
Else
MsgBox "No paper tray has been selected."
End If

Dim nCopy As Integer
nCopy = Me.txtCopies
Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
Set report = Nothing
xrpt.Close
Set xrpt = Nothing

End Sub

Sub forecastreport()
Dim xrpt As Recordset
Dim application As New CRAXDDRT.application
Dim report As New CRAXDDRT.report

Set xrpt = New Recordset
xrpt.Open sqlprint, Db, adOpenKeyset
Set report = application.OpenReport(App.path & "\REPORTs\Forecast.rpt")
report.Database.Tables(1).SetDataSource xrpt
    Dim xy As Byte
    For xy = 1 To 7
        report.FormulaFields(xy).Text = "'" & MonthName(Month((DateAdd("m", xy - 4, Format(Fbulan, "yyyy-mm-dd")))), True) & " " & Year((DateAdd("m", xy - 4, Format(Fbulan, "yyyy-mm-dd")))) & "'"
    Next
        report.FormulaFields(8).Text = "'Periode    : " & MonthName(Month(Fbulan), False) & " " & Year(Fbulan) & "'"
        report.FormulaFields(9).Text = "'Customer   : " & F_Factory & " / " & F_Cust_Name & "'"
        report.FormulaFields(10).Text = "'" & MonthName(Month(Fbulan), True) & "'"
        report.FormulaFields(11).Text = "'" & MonthName(Month(Fbulan), True) & "'"
        
    Dim rsCompany As New Recordset
    sql = "select Company_Name,Address1,Address2,City,Postal_Code,Phone1,Fax FROM company_Profile"
    Set rsCompany = Db.Execute(sql)
    If Not rsCompany.EOF Then
        report.FormulaFields(12).Text = "'" & Trim(rsCompany!company_name) & "'"
        report.FormulaFields(13).Text = "'" & Trim(rsCompany!address1) & "'"
        report.FormulaFields(14).Text = "'" & Trim(rsCompany!address2) & "'"
        report.FormulaFields(15).Text = "'" & Trim(rsCompany!City) & " " & Trim(rsCompany!postal_code) & "'"
        report.FormulaFields(16).Text = "'" & "Telp : " & Trim(rsCompany!phone1) & ", Fax : " & Trim(rsCompany!fax) & "'"
    End If
        
    Dim vBinNumbers As Variant
    If combo1.ListIndex >= 0 Then
    vBinNumbers = GetBinNumbers
    report.SelectPrinter Label3, cbxPrinters.Text, Label2
    report.PaperSize = crPaperA4
    report.PaperOrientation = Orient
    report.PaperSource = vBinNumbers(combo1.ListIndex)
    Else
    MsgBox "No paper tray has been selected."
    End If

    Dim nCopy As Integer
    nCopy = Me.txtCopies
    Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
    Set report = Nothing
    xrpt.Close
    Set xrpt = Nothing

End Sub


Function cantprint(noinvoice$) As Boolean
Dim rstinv As Recordset

sql = "select * from invoice_master where invoice_no in (" & noinvoice & ")"
Set rstinv = New Recordset
rstinv.Open sql, Db, adOpenDynamic, adLockOptimistic
If Not rstinv.EOF Then
    If IsNull(rstinv!fix_cls) Then
        cantprint = True
    Else
        cantprint = False
    End If
End If
End Function

Sub Invoice()

    Dim application As New CRAXDDRT.application
    Dim report As New CRAXDDRT.report
    Dim rs1 As New Recordset
    Dim Y As Long, xy As Long, cbol As Boolean, tbank As String

    If rs1.State <> adStateClosed Then rs1.Close
    rs1.CursorLocation = adUseClient
    rs1.Open sqlprint, Db, adOpenKeyset
    If Not rs1.EOF Then
    
        Set report = application.OpenReport(App.path & "\Reports\invoice.rpt")
        report.Database.Tables(1).SetDataSource rs1
        
        If Trim(ginvno) <> "" Then inv_no = ginvno
        
        Dim rs2 As New Recordset
        sql = "select cb.*, cc.Description curr " & _
              vbLf & "from Company_bank cb " & _
              vbLf & "left join Curr_Cls cc " & _
              vbLf & "on cb.Currency_Code = cc.Curr_Cls " & _
              vbLf & "order by cb.bank_name, cb.address1, cb.address2, cb.currency_code "
              
        If rs2.State <> adStateClosed Then rs2.Close
        rs2.Open sql, Db, adOpenKeyset, adLockOptimistic
            
        If Not rs2.EOF Then
            For Y = 0 To rs2.RecordCount - 1
                If Y = 0 Then
                    cbol = True
                    tbank = Trim(rs2!bank_name) & Trim(rs2!address1) & Trim(rs2!address2)
                    report.FormulaFields.GetItemByName("Bank1").Text = "'" & Trim(rs2!bank_name) & "'"
                    report.FormulaFields.GetItemByName("Bank1Add1").Text = "'" & Trim(rs2!address1) & "'"
                    report.FormulaFields.GetItemByName("Bank1Add2").Text = "'" & Trim(rs2!address2) & "'"
                    report.FormulaFields.GetItemByName("Bank1City").Text = "'" & Trim(rs2!City) & "'"
                    report.FormulaFields.GetItemByName("Bank1KodePos").Text = "'" & Trim(rs2!postal_code) & "'"
                    report.FormulaFields.GetItemByName("Bank1Acc1").Text = "'" & Trim(rs2!Curr) & " " & Trim(rs2!Account_No) & "'"
                    xy = 1
                Else
                    If tbank = Trim(rs2!bank_name) & Trim(rs2!address1) & Trim(rs2!address2) Then
                        If xy < 7 Then xy = xy + 1 'Account s/d 7
                        If cbol Then
                            report.FormulaFields.GetItemByName("Bank1Acc" & xy).Text = "'" & Trim(rs2!Curr) & " " & Trim(rs2!Account_No) & "'"
                        Else
                            report.FormulaFields.GetItemByName("Bank2Acc" & xy).Text = "'" & Trim(rs2!Curr) & " " & Trim(rs2!Account_No) & "'"
                        End If
                      
                    Else
                        If cbol = False Then Exit For
                        cbol = False 'Utk pembatas jumlah (nama_bank + addr1 + addr2) yg beda sebyk 2
                        xy = 1
                        tbank = Trim(rs2!bank_name) & Trim(rs2!address1) & Trim(rs2!address2)
                        report.FormulaFields.GetItemByName("Bank2").Text = "'" & Trim(rs2!bank_name) & "'"
                        report.FormulaFields.GetItemByName("Bank2Add1").Text = "'" & Trim(rs2!address1) & "'"
                        report.FormulaFields.GetItemByName("Bank2Add2").Text = "'" & Trim(rs2!address2) & "'"
                        report.FormulaFields.GetItemByName("Bank2City").Text = "'" & Trim(rs2!City) & "'"
                        report.FormulaFields.GetItemByName("Bank2KodePos").Text = "'" & Trim(rs2!postal_code) & "'"
                        report.FormulaFields.GetItemByName("Bank2Acc1").Text = "'" & Trim(rs2!Curr) & " " & Trim(rs2!Account_No) & "'"
                    End If
                End If
                rs2.MoveNext
            Next
        End If
        
        report.FormulaFields.GetItemByName("DecimalQty").Text = gi_decimalDigitQty
        report.FormulaFields.GetItemByName("DecimalPrice").Text = gi_decimalDigitPrice
        report.FormulaFields.GetItemByName("DecimalPriceIDR").Text = gi_decimalDigitPriceIDR
        report.FormulaFields.GetItemByName("DecimalAmount").Text = gi_decimalDigitAmount
        report.FormulaFields.GetItemByName("DecimalAmountIDR").Text = gi_decimalDigitAmountIDR
        
        Dim vBinNumbers As Variant
        If combo1.ListIndex >= 0 Then
            vBinNumbers = GetBinNumbers
            report.SelectPrinter Label3, cbxPrinters.Text, Label2
            report.PaperOrientation = Orient
            report.PaperSource = vBinNumbers(combo1.ListIndex)
        Else
            MsgBox "No paper tray has been selected."
        End If
        
        Dim nCopy As Integer
        nCopy = Me.txtCopies
        Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
        Set report = Nothing
        Set rs1 = Nothing
        Set rs2 = Nothing

    End If

End Sub

Sub InvoiceEx()

    Dim application As New CRAXDDRT.application
    Dim report As New CRAXDDRT.report
    Dim rsMain As New Recordset
    Dim rsSub As New Recordset
    Dim RsSubDetail As New Recordset

    rsMain.CursorLocation = adUseClient
    If rsMain.State <> adStateClosed Then rsMain.Close
    rsMain.Open sqlprint, Db, adOpenKeyset
    If Not rsMain.EOF Then
    
        Set report = application.OpenReport(App.path & "\Reports\invoice_ex.rpt")
        report.Database.Tables(1).SetDataSource rsMain
       
        report.FormulaFields.GetItemByName("DecimalQty").Text = gi_decimalDigitQty
        report.FormulaFields.GetItemByName("DecimalPrice").Text = gi_decimalDigitPrice
        report.FormulaFields.GetItemByName("DecimalPriceIDR").Text = gi_decimalDigitPriceIDR
        report.FormulaFields.GetItemByName("DecimalAmount").Text = gi_decimalDigitAmountIDR
        
        If rsSub.State <> adStateClosed Then rsSub.Close
        rsSub.Open sqlprint2, Db, adOpenKeyset, adLockOptimistic
        With report.OpenSubreport("sub_packing_list_ex")
         .Database.Tables(1).SetDataSource rsSub
         .FormulaFields.GetItemByName("SubDecimalQty").Text = gi_decimalDigitQty
         .FormulaFields.GetItemByName("SubDecimalPrice").Text = gi_decimalDigitPrice
         .FormulaFields.GetItemByName("subDecimalPriceIDR").Text = gi_decimalDigitPriceIDR
         .FormulaFields.GetItemByName("subDecimalAmount").Text = gi_decimalDigitAmountIDR
        End With
        
        If RsSubDetail.State <> adStateClosed Then RsSubDetail.Close
        RsSubDetail.Open sqlprint3, Db, adOpenKeyset, adLockOptimistic
    
        If Not RsSubDetail.EOF Then
         With report.OpenSubreport("subDetailOf_packing_list_ex")
          .Database.Tables(1).SetDataSource RsSubDetail
          .FormulaFields.GetItemByName("SubDetailDecimalQty").Text = gi_decimalDigitQty
          .FormulaFields.GetItemByName("SubDetailDecimalPrice").Text = gi_decimalDigitPrice
          .FormulaFields.GetItemByName("SubDetailDecimalPriceIDR").Text = gi_decimalDigitPriceIDR
          .FormulaFields.GetItemByName("SubDetailDecimalAmount").Text = gi_decimalDigitAmountIDR
         End With
        End If
        
        Dim vBinNumbers As Variant
        If combo1.ListIndex >= 0 Then
            vBinNumbers = GetBinNumbers
            report.SelectPrinter Label3, cbxPrinters.Text, Label2
            report.PaperOrientation = Orient
            report.PaperSource = vBinNumbers(combo1.ListIndex)
        Else
            MsgBox "No paper tray has been selected."
        End If
        
        Dim nCopy As Integer
        nCopy = Me.txtCopies
        Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
        Set report = Nothing
        
        Set RsSubDetail = Nothing
        Set rsSub = Nothing
        Set rsMain = Nothing
    
    End If

End Sub

Sub InvoiceKWI()

    Dim application As New CRAXDDRT.application
    Dim report As New CRAXDDRT.report
    Dim rsMain As New Recordset
    Dim rsSub As New Recordset
    Dim RsSubDetail As New Recordset

    rsMain.CursorLocation = adUseClient
    If rsMain.State <> adStateClosed Then rsMain.Close
    rsMain.Open sqlprint, Db, adOpenKeyset
    If Not rsMain.EOF Then
    
        Set report = application.OpenReport(App.path & "\Reports\invoice_kwi.rpt")
        report.Database.Tables(1).SetDataSource rsMain
        report.DiscardSavedData
               
        Dim vBinNumbers As Variant
        If combo1.ListIndex >= 0 Then
            vBinNumbers = GetBinNumbers
            report.SelectPrinter Label3, cbxPrinters.Text, Label2
            report.PaperOrientation = Orient
            report.PaperSource = vBinNumbers(combo1.ListIndex)
        Else
            MsgBox "No paper tray has been selected."
        End If
        
        Dim nCopy As Integer
        nCopy = Me.txtCopies
        Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
        Set report = Nothing
        
        Set RsSubDetail = Nothing
        Set rsSub = Nothing
        Set rsMain = Nothing
    
    End If

End Sub

Sub Pajak()

    Dim application As New CRAXDDRT.application
    Dim report As New CRAXDDRT.report
    Dim rs1 As Recordset
    
    Set rs1 = New Recordset
    rs1.CursorLocation = adUseClient
    rs1.Open sqlprint, Db, adOpenKeyset
    If Not rs1.EOF Then
        Set report = application.OpenReport(App.path & "\Reports\faktur_pajak_std.rpt")
        report.Database.Tables(1).SetDataSource rs1
        
    '#####################################################################
    '# Qty Digit and decimal
    report.FormulaFields(39).Text = "" & gi_decimalDigitAmount & ""
    report.FormulaFields(40).Text = "" & gi_decimalDigitAmount & ""
    report.FormulaFields(42).Text = "" & gi_decimalDigitExchangeRate & ""
    report.FormulaFields(43).Text = "" & gi_decimalDigitExchangeRate & ""
    '#####################################################################
    
        Dim vBinNumbers As Variant
        If combo1.ListIndex >= 0 Then
            vBinNumbers = GetBinNumbers
            report.SelectPrinter Label3, cbxPrinters.Text, Label2
            report.PaperSize = crPaperA4
            report.PaperOrientation = Orient
            report.PaperSource = vBinNumbers(combo1.ListIndex)
        Else
            MsgBox "No paper tray has been selected."
        End If
        
        Dim nCopy As Integer
        nCopy = Me.txtCopies
        Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
        Set report = Nothing
        rs1.Close
        Set rs1 = Nothing
    End If


End Sub

Sub TradeMaster()
  Dim application As New CRAXDDRT.application
  Dim report As New CRAXDDRT.report
  Dim rsRpt As New ADODB.Recordset

  If rsRpt.State <> adStateClosed Then rsRpt.Close
  rsRpt.Open sqlprint, Db, adOpenDynamic, adLockOptimistic

  If rsRpt.EOF Then Exit Sub

  Set report = application.OpenReport(App.path & "\Reports\rptTradeMaster.rpt")
  report.Database.Tables(1).SetDataSource rsRpt
  report.ReportTitle = "Trade Master"

    Dim vBinNumbers As Variant
    If combo1.ListIndex >= 0 Then
        vBinNumbers = GetBinNumbers
        report.SelectPrinter Label3, cbxPrinters.Text, Label2
        report.PaperSize = crPaperA4
        report.PaperOrientation = Orient
        report.PaperSource = vBinNumbers(combo1.ListIndex)
    Else
        MsgBox "No paper tray has been selected."
    End If

    Dim nCopy As Integer
    nCopy = Me.txtCopies
    Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
    Set report = Nothing
    rsRpt.Close
    Set rsRpt = Nothing
End Sub

Sub BOM()
Dim application As New CRAXDDRT.application
Dim report As New CRAXDDRT.report
Dim rsRpt As New ADODB.Recordset

    Set rsRpt = Db.Execute(sqlprint)

    If frmBOMInquiry.cboExplosion.Text = "1" Then
        Set report = application.OpenReport(App.path & "\Reports\BOM_implosion.rpt")
    Else
        Set report = application.OpenReport(App.path & "\Reports\BOM.rpt")
    End If
    report.Database.Tables(1).SetDataSource rsRpt

    '#####################################################################
    '# Qty Digit and decimal
    report.FormulaFields(6).Text = "" & gi_decimalDigitQtyBOM & ""
    report.FormulaFields(7).Text = "" & gi_decimalDigitQtyBOM & ""
    Select Case frmBOMInquiry.cboExplosion.Text
    Case 0: report.FormulaFields(8).Text = "'Explosion Of BOM'"
    Case 1: report.FormulaFields(8).Text = "'Implosion Of BOM'"
    Case 2: report.FormulaFields(8).Text = "'1 Level Explosion Of BOM'"
    End Select
    '#####################################################################
    
    Dim vBinNumbers As Variant
    If combo1.ListIndex >= 0 Then
        vBinNumbers = GetBinNumbers
        report.SelectPrinter Label3, cbxPrinters.Text, Label2
        report.PaperSize = crPaperA4
        report.PaperOrientation = Orient
        report.PaperSource = vBinNumbers(combo1.ListIndex)
    Else
        MsgBox "No paper tray has been selected."
    End If

    Dim nCopy As Integer
    nCopy = Me.txtCopies
    Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))

    Set rsRpt = Nothing
End Sub

Sub rptFormula()
Dim application As New CRAXDDRT.application
Dim report As New CRAXDDRT.report
Dim rsRpt As New ADODB.Recordset

    Set rsRpt = Db.Execute(sqlprint)

    Set report = application.OpenReport(App.path & "\Reports\rptFormula.rpt")
    report.Database.Tables(1).SetDataSource rsRpt

    Dim vBinNumbers As Variant
    If combo1.ListIndex >= 0 Then
        vBinNumbers = GetBinNumbers
        report.SelectPrinter Label3, cbxPrinters.Text, Label2
        'report.PaperSize = crPaperA3
        report.PaperOrientation = Orient
        report.PaperSource = vBinNumbers(combo1.ListIndex)
    Else
        MsgBox "No paper tray has been selected."
    End If

    Dim nCopy As Integer
    nCopy = Me.txtCopies
    Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))

    Set rsRpt = Nothing
End Sub

Sub ProdResultInquiry()
Dim application As New CRAXDDRT.application
Dim report As New CRAXDDRT.report
Dim rsRpt As New ADODB.Recordset
Dim sqlResult As String

    Set rsRpt = Db.Execute(sqlprint)
    Set report = application.OpenReport(App.path & "\Reports\rptProdResultInquiry.rpt")
    report.Database.Tables(1).SetDataSource rsRpt

    report.FormulaFields(1).Text = "'" & tglAwalRptPrint & "'"
    report.FormulaFields(2).Text = "'" & tglAkhirRptPrint & "'"

    Dim vBinNumbers As Variant
    If combo1.ListIndex >= 0 Then
        vBinNumbers = GetBinNumbers
        report.SelectPrinter Label3, cbxPrinters.Text, Label2
        report.PaperSize = crPaperA3
        report.PaperOrientation = Orient
        report.PaperSource = vBinNumbers(combo1.ListIndex)
    Else
        MsgBox "No paper tray has been selected."
    End If

    Dim nCopy As Integer
    nCopy = Me.txtCopies
   Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))

    Set rsRpt = Nothing
End Sub

Sub ProdResultByFactory()
Dim application As New CRAXDDRT.application
Dim report As New CRAXDDRT.report
Dim rsRpt As New ADODB.Recordset
Dim sqlResult As String

    Set rsRpt = Db.Execute(sqlprint)
    Set report = application.OpenReport(App.path & "\Reports\rptProdResult.rpt")
    report.Database.Tables(1).SetDataSource rsRpt

    report.FormulaFields(1).Text = "'" & tglAwalRptPrint & "'"
    report.FormulaFields(2).Text = "'" & tglAkhirRptPrint & "'"
    
    '#####################################################################
    '# Qty Digit and decimal
    report.FormulaFields(7).Text = "" & gi_decimalDigitQty & ""
    report.FormulaFields(8).Text = "" & gi_decimalDigitQty & ""
    '#####################################################################
    
    Dim vBinNumbers As Variant
    If combo1.ListIndex >= 0 Then
        vBinNumbers = GetBinNumbers
        report.SelectPrinter Label3, cbxPrinters.Text, Label2
        report.PaperSize = crPaperA4
        report.PaperOrientation = Orient
        report.PaperSource = vBinNumbers(combo1.ListIndex)
    Else
        MsgBox "No paper tray has been selected."
    End If

    Dim nCopy As Integer
    nCopy = Me.txtCopies
Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))

    Set rsRpt = Nothing
End Sub

Sub ProdResultByItem()
Dim application As New CRAXDDRT.application
Dim report As New CRAXDDRT.report
Dim rsRpt As New ADODB.Recordset
Dim sqlResult As String

    Set rsRpt = Db.Execute(sqlprint)
    Set report = application.OpenReport(App.path & "\Reports\rptProdResultByItem.rpt")
    report.Database.Tables(1).SetDataSource rsRpt

    report.FormulaFields(1).Text = "'" & tglAwalRptPrint & "'"
    report.FormulaFields(2).Text = "'" & tglAkhirRptPrint & "'"

    Dim vBinNumbers As Variant
    If combo1.ListIndex >= 0 Then
        vBinNumbers = GetBinNumbers
        report.SelectPrinter Label3, cbxPrinters.Text, Label2
        report.PaperSize = crPaperA3
        report.PaperOrientation = Orient
        report.PaperSource = vBinNumbers(combo1.ListIndex)
    Else
        MsgBox "No paper tray has been selected."
    End If

    Dim nCopy As Integer
    nCopy = Me.txtCopies
    Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))

    Set rsRpt = Nothing
End Sub

Sub pricemaster()
Dim application As New CRAXDDRT.application
Dim report As New CRAXDDRT.report
Dim rsRpt As New ADODB.Recordset

  If rsRpt.State <> adStateClosed Then rsRpt.Close
  rsRpt.Open sqlprint, Db, adOpenDynamic, adLockOptimistic

  If rsRpt.EOF Then Exit Sub

  Set report = application.OpenReport(App.path & "\Reports\rptPriceMaster.rpt")
  report.Database.Tables(1).SetDataSource rsRpt
  '#####################################################################
   '# Price Digit and decimal
    report.FormulaFields(5).Text = "" & gi_decimalDigitPrice & ""
    report.FormulaFields(6).Text = "" & gi_decimalDigitPriceIDR & ""
  '#####################################################################

    Dim vBinNumbers As Variant
    If combo1.ListIndex >= 0 Then
        vBinNumbers = GetBinNumbers
        report.SelectPrinter Label3, cbxPrinters.Text, Label2
        report.PaperSize = crPaperA4
        report.PaperOrientation = Orient
        report.PaperSource = vBinNumbers(combo1.ListIndex)
    Else
        MsgBox "No paper tray has been selected."
    End If

    Dim nCopy As Integer
    nCopy = Me.txtCopies
Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
    Set report = Nothing
    rsRpt.Close
    Set rsRpt = Nothing
End Sub

Sub custexchrate()
  Dim application As New CRAXDDRT.application
  Dim report As New CRAXDDRT.report
  Dim rsRpt As New ADODB.Recordset

  If rsRpt.State <> adStateClosed Then rsRpt.Close
  rsRpt.Open sqlprint, Db, adOpenDynamic, adLockOptimistic

  If rsRpt.EOF Then Exit Sub

  Set report = application.OpenReport(App.path & "\Reports\rptCustExchRate.rpt")
  report.Database.Tables(1).SetDataSource rsRpt
  report.FormulaFields(2).Text = Fbulan

    Dim vBinNumbers As Variant
    If combo1.ListIndex >= 0 Then
        vBinNumbers = GetBinNumbers
        report.SelectPrinter Label3, cbxPrinters.Text, Label2
        report.PaperSize = crPaperA4
        report.PaperOrientation = Orient
        report.PaperSource = vBinNumbers(combo1.ListIndex)
    Else
        MsgBox "No paper tray has been selected."
    End If

    Dim nCopy As Integer
    nCopy = Me.txtCopies
    
    Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
    Set report = Nothing
    rsRpt.Close
    Set rsRpt = Nothing
End Sub

Sub dailyproduction()
    Dim application As New CRAXDDRT.application
    Dim report As New CRAXDDRT.report
    Dim rsRpt As New ADODB.Recordset

    If rsRpt.State <> adStateClosed Then rsRpt.Close
    rsRpt.Open sqlprint, Db, adOpenDynamic, adLockOptimistic

    If rsRpt.EOF Then Exit Sub

    Set report = application.OpenReport(App.path & "\Reports\rptdailyproduction.rpt")
    report.Database.Tables(1).SetDataSource rsRpt
    report.FormulaFields(1).Text = "'" & Format(Fbulan, "dd MMM yyyy") & "'"
    report.FormulaFields(2).Text = "'" & Format(Ftahun, "dd MMM yyyy") & "'"
    '#####################################################################
    '# Qty Digit and decimal
    report.FormulaFields(4).Text = "" & gi_decimalDigitQty & ""
    report.FormulaFields(5).Text = "" & gi_decimalDigitQty & ""
    '#####################################################################
    Dim vBinNumbers As Variant
    If combo1.ListIndex >= 0 Then
        vBinNumbers = GetBinNumbers
        report.SelectPrinter Label3, cbxPrinters.Text, Label2
        report.PaperSize = crPaperA4
        report.PaperOrientation = Orient
        report.PaperSource = vBinNumbers(combo1.ListIndex)
    Else
        MsgBox "No paper tray has been selected."
    End If

    Dim nCopy As Integer
    nCopy = Me.txtCopies
    Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
    Set report = Nothing
    rsRpt.Close
    Set rsRpt = Nothing
End Sub



Sub BookKeepingExchRate()
    Dim AppRpt As New CRAXDDRT.application
    Dim Rpt2 As New CRAXDDRT.report
    Dim rs1 As New ADODB.Recordset, rs2 As New ADODB.Recordset
    Dim sql1 As String, sql2 As String
    'u Formula Title
    Dim Judul1 As String, CurrBook As String, CurrTax As String

    'U/ Formula di Book_exchrate
    Dim Jan1 As String, Feb1 As String, Mar1 As String, Apr1 As String, Mei1 As String, Jun1 As String, Jul1 As String, Aug1 As String, Sep1 As String, Okt1 As String, Nov1 As String, Des1 As String
    Dim Jan2 As String, Feb2 As String, Mar2 As String, Apr2 As String, Mei2 As String, Jun2 As String, Jul2 As String, Aug2 As String, Sep2 As String, Okt2 As String, Nov2 As String, Des2 As String
    'u/ di Tax_exchrate
    Dim Jan3 As String, Feb3 As String, Mar3 As String, Apr3 As String, Mei3 As String, Jun3 As String, Jul3 As String, Aug3 As String, Sep3 As String, Okt3 As String, Nov3 As String, Des3 As String
    Dim Jan4 As String, Feb4 As String, Mar4 As String, Apr4 As String, Mei4 As String, Jun4 As String, Jul4 As String, Aug4 As String, Sep4 As String, Okt4 As String, Nov4 As String, Des4 As String
    Dim Jan5 As String, Feb5 As String, Mar5 As String, Apr5 As String, Mei5 As String, Jun5 As String, Jul5 As String, Aug5 As String, Sep5 As String, Okt5 As String, Nov5 As String, Des5 As String
    Dim Jan6 As String, Feb6 As String, Mar6 As String, Apr6 As String, Mei6 As String, Jun6 As String, Jul6 As String, Aug6 As String, Sep6 As String, Okt6 As String, Nov6 As String, Des6 As String
    Dim Jan7 As String, Feb7 As String, Mar7 As String, Apr7 As String, Mei7 As String, Jun7 As String, Jul7 As String, Aug7 As String, Sep7 As String, Okt7 As String, Nov7 As String, Des7 As String

    Judul1 = "Data Base of Exchange Rate " & vbCrLf & "Book Keeping in Monthly Base"
    Set Rpt2 = AppRpt.OpenReport(App.path & "\Reports\Rpt_Exchange.rpt")

    CurrBook = xbln
    CurrTax = xbln

    Set rs1 = New Recordset

    sql1 = "select * from book_exchangerate  " & _
         " where Exch_year='" & Ftahun & "' and Currency_code='" & Fbulan & "' order by Currency_code,Term_cls"
    rs1.CursorLocation = adUseClient
    rs1.Open sql1, Db, adOpenDynamic, adLockBatchOptimistic

    Jan1 = 0: Feb1 = 0: Mar1 = 0: Apr1 = 0: Mei1 = 0: Jun1 = 0: Jul1 = 0: Aug1 = 0: Sep1 = 0: Okt1 = 0: Nov1 = 0: Des1 = 0
    Jan2 = 0: Feb2 = 0: Mar2 = 0: Apr2 = 0: Mei2 = 0: Jun2 = 0: Jul2 = 0: Aug2 = 0: Sep2 = 0: Okt2 = 0: Nov2 = 0: Des2 = 0
    If rs1.EOF Then

    Else
        While Not rs1.EOF
            If Trim$(rs1!Term_cls) = "1" Then 'Beginning
                If IsNull(rs1!Exch01) Then
                    Jan1 = "-   "
                Else
                    Jan1 = Format(rs1!Exch01, "#,#0.00")
                End If
                If IsNull(rs1!Exch02) Then
                    Feb1 = "-   "
                Else
                    Feb1 = Format(rs1!Exch02, "#,#0.00")
                End If
                If IsNull(rs1!Exch03) Then
                    Mar1 = "-   "
                Else
                    Mar1 = Format(rs1!Exch03, "#,#0.00")
                End If
                If IsNull(rs1!Exch04) Then
                    Apr1 = "-   "
                Else
                    Apr1 = Format(rs1!Exch04, "#,#0.00")
                End If
                If IsNull(rs1!Exch05) Then
                    Mei1 = "-   "
                Else
                    Mei1 = Format(rs1!Exch05, "#,#0.00")
                End If
                If IsNull(rs1!Exch06) Then
                    Jun1 = "-   "
                Else
                    Jun1 = Format(rs1!Exch06, "#,#0.00")
                End If
                If IsNull(rs1!Exch07) Then
                    Jul1 = "-   "
                Else
                    Jul1 = Format(rs1!Exch07, "#,#0.00")
                End If
                If IsNull(rs1!Exch08) Then
                    Aug1 = "-   "
                Else
                    Aug1 = Format(rs1!Exch08, "#,#0.00")
                End If
                If IsNull(rs1!Exch09) Then
                    Sep1 = "-   "
                Else
                    Sep1 = Format(rs1!Exch09, "#,#0.00")
                End If
                If IsNull(rs1!Exch010) Then
                    Okt1 = "-   "
                Else
                    Okt1 = Format(rs1!Exch010, "#,#0.00")
                End If
                If IsNull(rs1!Exch011) Then
                    Nov1 = "-   "
                Else
                    Nov1 = Format(rs1!Exch011, "#,#0.00")
                End If
                If IsNull(rs1!Exch012) Then
                    Des1 = "-   "
                Else
                    Des1 = Format(rs1!Exch012, "#,#0.00")
                End If
            ElseIf Trim$(rs1!Term_cls) = "2" Then 'Beginning
                If IsNull(rs1!Exch01) Then
                    Jan2 = "-   "
                Else
                    Jan2 = Format(rs1!Exch01, "#,#0.00")
                End If
                If IsNull(rs1!Exch02) Then
                    Feb2 = "-   "
                Else
                    Feb2 = Format(rs1!Exch02, "#,#0.00")
                End If
                If IsNull(rs1!Exch03) Then
                    Mar2 = "-   "
                Else
                    Mar2 = Format(rs1!Exch03, "#,#0.00")
                End If
                If IsNull(rs1!Exch04) Then
                    Apr2 = "-   "
                Else
                    Apr2 = Format(rs1!Exch04, "#,#0.00")
                End If
                If IsNull(rs1!Exch05) Then
                    Mei2 = "-   "
                Else
                    Mei2 = Format(rs1!Exch05, "#,#0.00")
                End If
                If IsNull(rs1!Exch06) Then
                    Jun2 = "-   "
                Else
                    Jun2 = Format(rs1!Exch06, "#,#0.00")
                End If
                If IsNull(rs1!Exch07) Then
                    Jul2 = "-   "
                Else
                    Jul2 = Format(rs1!Exch07, "#,#0.00")
                End If
                If IsNull(rs1!Exch08) Then
                    Aug2 = "-   "
                Else
                    Aug2 = Format(rs1!Exch08, "#,#0.00")
                End If
                If IsNull(rs1!Exch09) Then
                    Sep2 = "-   "
                Else
                    Sep2 = Format(rs1!Exch09, "#,#0.00")
                End If
                If IsNull(rs1!Exch010) Then
                    Okt2 = "-   "
                Else
                    Okt2 = Format(rs1!Exch010, "#,#0.00")
                End If
                If IsNull(rs1!Exch011) Then
                    Nov2 = "-   "
                Else
                    Nov2 = Format(rs1!Exch011, "#,#0.00")
                End If
                If IsNull(rs1!Exch012) Then
                    Des2 = "-   "
                Else
                    Des2 = Format(rs1!Exch012, "#,#0.00")
                End If
            End If
        rs1.MoveNext
        Wend
    End If

    sql2 = "Select * from Tax_exchangerate where Exch_year='" & Ftahun & "' and Currency_code='" & Fbulan & "' order by Exch_Month,Week_code,Currency_code"

    Set rs2 = Db.Execute(sql2)

    Jan3 = "-   ": Feb3 = "-   ": Mar3 = "-   ": Apr3 = "-   ": Mei3 = "-   ": Jun3 = "-   ": Jul3 = "-   ": Aug3 = "-   ": Sep3 = "-   ": Okt3 = "-   ": Nov3 = "-   ": Des3 = "-   "
    Jan4 = "-   ": Feb4 = "-   ": Mar4 = "-   ": Apr4 = "-   ": Mei4 = "-   ": Jun4 = "-   ": Jul4 = "-   ": Aug4 = "-   ": Sep4 = "-   ": Okt4 = "-   ": Nov4 = "-   ": Des4 = "-   "
    Jan5 = "-   ": Feb5 = "-   ": Mar5 = "-   ": Apr5 = "-   ": Mei5 = "-   ": Jun5 = "-   ": Jul5 = "-   ": Aug5 = "-   ": Sep5 = "-   ": Okt5 = "-   ": Nov5 = "-   ": Des5 = "-   "
    Jan6 = "-   ": Feb6 = "-   ": Mar6 = "-   ": Apr6 = "-   ": Mei6 = "-   ": Jun6 = "-   ": Jul6 = "-   ": Aug6 = "-   ": Sep6 = "-   ": Okt6 = "-   ": Nov6 = "-   ": Des6 = "-   "
    Jan7 = "-   ": Feb7 = "-   ": Mar7 = "-   ": Apr7 = "-   ": Mei7 = "-   ": Jun7 = "-   ": Jul7 = "-   ": Aug7 = "-   ": Sep7 = "-   ": Okt7 = "-   ": Nov7 = "-   ": Des7 = "-   "

    While Not rs2.EOF
        'if exch
        If rs2!week_code = "1" Then
            If rs2!exch_month = "1" Then Jan3 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "2" Then Feb3 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "3" Then Mar3 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "4" Then Apr3 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "5" Then Mei3 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "6" Then Jun3 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "7" Then Jul3 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "8" Then Aug3 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "9" Then Sep3 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "10" Then Okt3 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "11" Then Nov3 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "12" Then Des3 = Format(rs2!Tax_exchangerate, "#,#0.00")
        ElseIf rs2!week_code = "2" Then
            If rs2!exch_month = "1" Then Jan4 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "2" Then Feb4 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "3" Then Mar4 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "4" Then Apr4 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "5" Then Mei4 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "6" Then Jun4 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "7" Then Jul4 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "8" Then Aug4 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "9" Then Sep4 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "10" Then Okt4 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "11" Then Nov4 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "12" Then Des4 = Format(rs2!Tax_exchangerate, "#,#0.00")
        ElseIf rs2!week_code = "3" Then
            If rs2!exch_month = "1" Then Jan5 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "2" Then Feb5 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "3" Then Mar5 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "4" Then Apr5 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "5" Then Mei5 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "6" Then Jun5 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "7" Then Jul5 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "8" Then Aug5 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "9" Then Sep5 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "10" Then Okt5 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "11" Then Nov5 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "12" Then Des5 = Format(rs2!Tax_exchangerate, "#,#0.00")
        ElseIf rs2!week_code = "4" Then
            If rs2!exch_month = "1" Then Jan6 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "2" Then Feb6 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "3" Then Mar6 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "4" Then Apr6 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "5" Then Mei6 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "6" Then Jun6 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "7" Then Jul6 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "8" Then Aug6 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "9" Then Sep6 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "10" Then Okt6 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "11" Then Nov6 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "12" Then Des6 = Format(rs2!Tax_exchangerate, "#,#0.00")
        ElseIf rs2!week_code = "5" Then
            If rs2!exch_month = "1" Then Jan7 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "2" Then Feb7 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "3" Then Mar7 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "4" Then Apr7 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "5" Then Mei7 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "6" Then Jun7 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "7" Then Jul7 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "8" Then Aug7 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "9" Then Sep7 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "10" Then Okt7 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "11" Then Nov7 = Format(rs2!Tax_exchangerate, "#,#0.00")
            If rs2!exch_month = "12" Then Des7 = Format(rs2!Tax_exchangerate, "#,#0.00")
        End If
    rs2.MoveNext
Wend

    Dim sqlcom As String
    Dim rscom As New Recordset
    sqlcom = "select company_name from company_profile"
    Set rscom = Db.Execute(sqlcom)
    
    Rpt2.Database.Tables(2).SetDataSource rscom


    Rpt2.FormulaFields(2).Text = "'" & CurrBook & "'"
    Rpt2.FormulaFields(3).Text = "'" & CurrTax & "'"
    'beginning
    Rpt2.FormulaFields(4).Text = "'" & Jan1 & "'"
    Rpt2.FormulaFields(5).Text = "'" & Feb1 & "'"
    Rpt2.FormulaFields(6).Text = "'" & Mar1 & "'"
    Rpt2.FormulaFields(7).Text = "'" & Apr1 & "'"
    Rpt2.FormulaFields(8).Text = "'" & Mei1 & "'"
    Rpt2.FormulaFields(9).Text = "'" & Jun1 & "'"
    Rpt2.FormulaFields(10).Text = "'" & Jul1 & "'"
    Rpt2.FormulaFields(11).Text = "'" & Aug1 & "'"
    Rpt2.FormulaFields(12).Text = "'" & Sep1 & "'"
    Rpt2.FormulaFields(13).Text = "'" & Okt1 & "'"
    Rpt2.FormulaFields(14).Text = "'" & Nov1 & "'"
    Rpt2.FormulaFields(15).Text = "'" & Des1 & "'"
    'Ending
    Rpt2.FormulaFields(16).Text = "'" & Jan2 & "'"
    Rpt2.FormulaFields(17).Text = "'" & Feb2 & "'"
    Rpt2.FormulaFields(18).Text = "'" & Mar2 & "'"
    Rpt2.FormulaFields(19).Text = "'" & Apr2 & "'"
    Rpt2.FormulaFields(20).Text = "'" & Mei2 & "'"
    Rpt2.FormulaFields(21).Text = "'" & Jun2 & "'"
    Rpt2.FormulaFields(22).Text = "'" & Jul2 & "'"
    Rpt2.FormulaFields(23).Text = "'" & Aug2 & "'"
    Rpt2.FormulaFields(24).Text = "'" & Sep2 & "'"
    Rpt2.FormulaFields(25).Text = "'" & Okt2 & "'"
    Rpt2.FormulaFields(26).Text = "'" & Nov2 & "'"
    Rpt2.FormulaFields(27).Text = "'" & Des2 & "'"

    'Tax Rate
    'Week 1
    Rpt2.FormulaFields(28).Text = "'" & Jan3 & "'"
    Rpt2.FormulaFields(29).Text = "'" & Feb3 & "'"
    Rpt2.FormulaFields(30).Text = "'" & Mar3 & "'"
    Rpt2.FormulaFields(31).Text = "'" & Apr3 & "'"
    Rpt2.FormulaFields(32).Text = "'" & Mei3 & "'"
    Rpt2.FormulaFields(33).Text = "'" & Jun3 & "'"
    Rpt2.FormulaFields(34).Text = "'" & Jul3 & "'"
    Rpt2.FormulaFields(35).Text = "'" & Aug3 & "'"
    Rpt2.FormulaFields(36).Text = "'" & Sep3 & "'"
    Rpt2.FormulaFields(37).Text = "'" & Okt3 & "'"
    Rpt2.FormulaFields(38).Text = "'" & Nov3 & "'"
    Rpt2.FormulaFields(39).Text = "'" & Des3 & "'"
    'Week 2
    Rpt2.FormulaFields(40).Text = "'" & Jan4 & "'"
    Rpt2.FormulaFields(41).Text = "'" & Feb4 & "'"
    Rpt2.FormulaFields(42).Text = "'" & Mar4 & "'"
    Rpt2.FormulaFields(43).Text = "'" & Apr4 & "'"
    Rpt2.FormulaFields(44).Text = "'" & Mei4 & "'"
    Rpt2.FormulaFields(45).Text = "'" & Jun4 & "'"
    Rpt2.FormulaFields(46).Text = "'" & Jul4 & "'"
    Rpt2.FormulaFields(47).Text = "'" & Aug4 & "'"
    Rpt2.FormulaFields(48).Text = "'" & Sep4 & "'"
    Rpt2.FormulaFields(49).Text = "'" & Okt4 & "'"
    Rpt2.FormulaFields(50).Text = "'" & Nov4 & "'"
    Rpt2.FormulaFields(51).Text = "'" & Des4 & "'"
    'Week 3
    Rpt2.FormulaFields(52).Text = "'" & Jan5 & "'"
    Rpt2.FormulaFields(53).Text = "'" & Feb5 & "'"
    Rpt2.FormulaFields(54).Text = "'" & Mar5 & "'"
    Rpt2.FormulaFields(55).Text = "'" & Apr5 & "'"
    Rpt2.FormulaFields(56).Text = "'" & Mei5 & "'"
    Rpt2.FormulaFields(57).Text = "'" & Jun5 & "'"
    Rpt2.FormulaFields(58).Text = "'" & Jul5 & "'"
    Rpt2.FormulaFields(59).Text = "'" & Aug5 & "'"
    Rpt2.FormulaFields(60).Text = "'" & Sep5 & "'"
    Rpt2.FormulaFields(61).Text = "'" & Okt5 & "'"
    Rpt2.FormulaFields(62).Text = "'" & Nov5 & "'"
    Rpt2.FormulaFields(63).Text = "'" & Des5 & "'"
    'Week 4
    Rpt2.FormulaFields(64).Text = "'" & Jan6 & "'"
    Rpt2.FormulaFields(65).Text = "'" & Feb6 & "'"
    Rpt2.FormulaFields(66).Text = "'" & Mar6 & "'"
    Rpt2.FormulaFields(67).Text = "'" & Apr6 & "'"
    Rpt2.FormulaFields(68).Text = "'" & Mei6 & "'"
    Rpt2.FormulaFields(69).Text = "'" & Jun6 & "'"
    Rpt2.FormulaFields(70).Text = "'" & Jul6 & "'"
    Rpt2.FormulaFields(71).Text = "'" & Aug6 & "'"
    Rpt2.FormulaFields(72).Text = "'" & Sep6 & "'"
    Rpt2.FormulaFields(73).Text = "'" & Okt6 & "'"
    Rpt2.FormulaFields(74).Text = "'" & Nov6 & "'"
    Rpt2.FormulaFields(75).Text = "'" & Des6 & "'"
    'Week 5
    Rpt2.FormulaFields(76).Text = "'" & Jan7 & "'"
    Rpt2.FormulaFields(77).Text = "'" & Feb7 & "'"
    Rpt2.FormulaFields(78).Text = "'" & Mar7 & "'"
    Rpt2.FormulaFields(79).Text = "'" & Apr7 & "'"
    Rpt2.FormulaFields(80).Text = "'" & Mei7 & "'"
    Rpt2.FormulaFields(81).Text = "'" & Jun7 & "'"
    Rpt2.FormulaFields(82).Text = "'" & Jul7 & "'"
    Rpt2.FormulaFields(83).Text = "'" & Aug7 & "'"
    Rpt2.FormulaFields(84).Text = "'" & Sep7 & "'"
    Rpt2.FormulaFields(85).Text = "'" & Okt7 & "'"
    Rpt2.FormulaFields(86).Text = "'" & Nov7 & "'"
    Rpt2.FormulaFields(87).Text = "'" & Des7 & "'"
    Rpt2.FormulaFields(88).Text = "'" & Ftahun & "'"

    Dim vBinNumbers As Variant
    If combo1.ListIndex >= 0 Then
        vBinNumbers = GetBinNumbers
        Rpt2.SelectPrinter Label3, cbxPrinters.Text, Label2
        Rpt2.PaperSize = crPaperA4
        Rpt2.PaperOrientation = Orient
        Rpt2.PaperSource = vBinNumbers(combo1.ListIndex)
    Else
        MsgBox "No paper tray has been selected."
    End If

    Dim nCopy As Integer
    nCopy = Me.txtCopies
    
    Call Printout(Rpt2, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
    Set Rpt2 = Nothing
    rs1.Close
    Set rs1 = Nothing
    rs2.Close
    Set rs2 = Nothing
    Set rscom = Nothing
End Sub

Function cantprintpo(PONO$) As Boolean
Dim rsPO As New Recordset

sql = "select * from purchaseorder_master where po_no = '" & PONO & "' "
Set rsPO = Db.Execute(sql)

If Not rsPO.EOF Then
    If rsPO!fix_cls = 1 Then
        cantprintpo = False
    Else
        cantprintpo = True
    End If
End If

Set rsPO = Nothing
End Function

Function POLocalPrint(PONO$)
  
  Dim application As New CRAXDDRT.application
  Dim report As New CRAXDDRT.report
  Dim rsRpt As New ADODB.Recordset
  
  If rsRpt.State <> adStateClosed Then rsRpt.Close
  rsRpt.Open sqlprint, Db, adOpenDynamic, adLockOptimistic
  
  If rsRpt.EOF Then Exit Function

    Set report = application.OpenReport(App.path & "\Reports\po_lokal.rpt")
    report.Database.Tables(1).SetDataSource rsRpt
    
    report.FormulaFields(1).Text = "" & gi_decimalDigitQty & ""
    report.FormulaFields(2).Text = "" & gi_decimalDigitPrice & ""
    report.FormulaFields(3).Text = "" & gi_decimalDigitAmount & ""
    
    Dim vBinNumbers As Variant
    If combo1.ListIndex >= 0 Then
        vBinNumbers = GetBinNumbers
        report.SelectPrinter Label3, cbxPrinters.Text, Label2
        report.PaperSize = crPaperA4
        report.PaperOrientation = Orient
        report.PaperSource = vBinNumbers(combo1.ListIndex)
    Else
        MsgBox "No paper tray has been selected."
    End If

    Dim nCopy As Integer
    nCopy = Me.txtCopies
    Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
    Set report = Nothing
    rsRpt.Close
    Set rsRpt = Nothing

End Function

Function POImportPrint(PONO$)
  
  Dim application As New CRAXDDRT.application
  Dim report As New CRAXDDRT.report
  Dim rsRpt As New ADODB.Recordset
  Dim RsSubDetail As New Recordset
  
  If rsRpt.State <> adStateClosed Then rsRpt.Close
  rsRpt.Open sqlprint, Db, adOpenDynamic, adLockOptimistic
  
  If rsRpt.EOF Then Exit Function

     Set report = application.OpenReport(App.path & "\Reports\ReportPO.rpt")
    report.Database.Tables(1).SetDataSource rsRpt
    'report.DiscardSavedData
    
        Dim rsSubPO As New ADODB.Recordset
        If rsSubPO.State <> adStateClosed Then rsSubPO.Close
        rsSubPO.Open sqlprint, Db, adOpenKeyset, adLockOptimistic
        If Not rsSubPO.EOF Then
         With report.OpenSubreport("SubReportPO.rpt")
          .Database.Tables(1).SetDataSource rsSubPO
          '.FormulaFields.GetItemByName("SubDetailDecimalQty").Text = gi_decimalDigitQty
          '.FormulaFields.GetItemByName("SubDetailDecimalPrice").Text = gi_decimalDigitPrice
          '.FormulaFields.GetItemByName("SubDetailDecimalPriceIDR").Text = gi_decimalDigitPriceIDR
          '.FormulaFields.GetItemByName("SubDetailDecimalAmount").Text = gi_decimalDigitAmountIDR
         End With
        End If
    
    Dim vBinNumbers As Variant
    If combo1.ListIndex >= 0 Then
        vBinNumbers = GetBinNumbers
        report.SelectPrinter Label3, cbxPrinters.Text, Label2
        'report.PaperSize = crPaperLetter
        report.PaperOrientation = Orient
        report.PaperSource = vBinNumbers(combo1.ListIndex)
    Else
        MsgBox "No paper tray has been selected."
    End If

    Dim nCopy As Integer
    nCopy = Me.txtCopies
    Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
    Set report = Nothing
    rsRpt.Close
    Set rsRpt = Nothing

End Function

Function POSubconPrint(PONO$)
  
  Dim application As New CRAXDDRT.application
  Dim report As New CRAXDDRT.report
  Dim rsRpt As New ADODB.Recordset
  
  If rsRpt.State <> adStateClosed Then rsRpt.Close
  rsRpt.Open sqlprint, Db, adOpenDynamic, adLockOptimistic
  
  If rsRpt.EOF Then Exit Function

    Set report = application.OpenReport(App.path & "\Reports\po_subcon.rpt")
    report.Database.Tables(1).SetDataSource rsRpt
    report.DiscardSavedData
    
'    report.FormulaFields(1).Text = "" & gi_decimalDigitQty & ""
'    report.FormulaFields(2).Text = "" & gi_decimalDigitPrice & ""
'    report.FormulaFields(3).Text = "" & gi_decimalDigitAmount & ""
    
    Dim vBinNumbers As Variant
    If combo1.ListIndex >= 0 Then
        vBinNumbers = GetBinNumbers
        report.SelectPrinter Label3, cbxPrinters.Text, Label2
        report.PaperSize = crPaperA4
        report.PaperOrientation = Orient
        report.PaperSource = vBinNumbers(combo1.ListIndex)
    Else
        MsgBox "No paper tray has been selected."
    End If

    Dim nCopy As Integer
    nCopy = Me.txtCopies
    Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
    Set report = Nothing
    rsRpt.Close
    Set rsRpt = Nothing

End Function

Function POCoilPrint(PONO$)
  Dim application As New CRAXDDRT.application
  Dim report As New CRAXDDRT.report
  Dim rsRpt As New ADODB.Recordset
  
  If rsRpt.State <> adStateClosed Then rsRpt.Close
  rsRpt.Open sqlprint, Db, adOpenDynamic, adLockOptimistic
  
  If rsRpt.EOF Then Exit Function

    Set report = application.OpenReport(App.path & "\Reports\rptPONew.rpt")
    report.Database.Tables(1).SetDataSource rsRpt

'#####################################################################
'# Qty Digit and decimal
report.FormulaFields(3).Text = "" & gi_decimalDigitQty & ""
report.FormulaFields(4).Text = "" & gi_decimalDigitQty & ""
report.FormulaFields(5).Text = "" & gi_decimalDigitPrice & ""
report.FormulaFields(6).Text = "" & gi_decimalDigitPrice & ""
report.FormulaFields(7).Text = "" & gi_decimalDigitAmount & ""
report.FormulaFields(8).Text = "" & gi_decimalDigitAmount & ""

report.FormulaFields(9).Text = "" & gi_decimalDigitThickness & ""
report.FormulaFields(10).Text = "" & gi_decimalDigitThickness & ""
report.FormulaFields(11).Text = "" & gi_decimalDigitWidth & ""
report.FormulaFields(12).Text = "" & gi_decimalDigitWidth & ""
report.FormulaFields(13).Text = "" & gi_decimalDigitLength & ""
report.FormulaFields(14).Text = "" & gi_decimalDigitLength & ""
'#####################################################################

    Dim vBinNumbers As Variant
    If combo1.ListIndex >= 0 Then
        vBinNumbers = GetBinNumbers
        report.SelectPrinter Label3, cbxPrinters.Text, Label2
        report.PaperSize = crPaperA4
        report.PaperOrientation = Orient
        report.PaperSource = vBinNumbers(combo1.ListIndex)
    Else
        MsgBox "No paper tray has been selected."
    End If

    Dim nCopy As Integer
    nCopy = Me.txtCopies
    Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
    Set report = Nothing
    rsRpt.Close
    Set rsRpt = Nothing
End Function
Sub OutstandingMaterial()
  Dim application As New CRAXDDRT.application
  Dim report As New CRAXDDRT.report
  Dim rsRpt As New ADODB.Recordset
  
  If rsRpt.State <> adStateClosed Then rsRpt.Close
  rsRpt.Open sqlprint, Db, adOpenDynamic, adLockOptimistic
  
  If rsRpt.EOF Then Exit Sub
  
  Set report = application.OpenReport(App.path & "\Reports\rptOutstandingMat.rpt")
  report.Database.Tables(1).SetDataSource rsRpt
  report.FormulaFields(4).Text = "'" & Format(Fbulan, "MMMM yyyy") & "'"
  
    Dim vBinNumbers As Variant
    If combo1.ListIndex >= 0 Then
        vBinNumbers = GetBinNumbers
        report.SelectPrinter Label3, cbxPrinters.Text, Label2
        report.PaperSize = crPaperA3
        report.PaperOrientation = Orient
        report.PaperSource = vBinNumbers(combo1.ListIndex)
    Else
        MsgBox "No paper tray has been selected."
    End If

    Dim nCopy As Integer
    nCopy = Me.txtCopies
    Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
    Set report = Nothing
    rsRpt.Close
    Set rsRpt = Nothing
End Sub

Sub PiListDet()

        
              Dim application As New CRAXDDRT.application
              Dim report As New CRAXDDRT.report
              Dim rsRpt As New ADODB.Recordset
                
              If rsRpt.State <> adStateClosed Then rsRpt.Close
              rsRpt.Open sqlprint, Db, adOpenDynamic, adLockOptimistic
            
              If rsRpt.EOF Then Me.MousePointer = vbDefault: Exit Sub
        
              Set report = application.OpenReport(App.path & "\Reports\rpt_pi_list_Details.rpt")
              report.Database.Tables(1).SetDataSource rsRpt
                       
''#####################################################################
''# Qty Digit and decimal
report.FormulaFields(2).Text = "" & gi_decimalDigitQty & ""
report.FormulaFields(3).Text = "" & gi_decimalDigitQty & ""
''#####################################################################
             
             Dim dates As String
            
             
             report.FormulaFields(1).Text = "'" & datePiList & "'"
             report.ReportTitle = "Physical Inventory List ( Detail )"
            
    Dim vBinNumbers As Variant
    If combo1.ListIndex >= 0 Then
        vBinNumbers = GetBinNumbers
        report.SelectPrinter Label3, cbxPrinters.Text, Label2
        report.PaperSize = crPaperA4
        report.PaperOrientation = Orient
        report.PaperSource = vBinNumbers(combo1.ListIndex)
    Else
        MsgBox "No paper tray has been selected."
    End If

    Dim nCopy As Integer
    nCopy = Me.txtCopies
    Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
    Set report = Nothing
    rsRpt.Close
    Set rsRpt = Nothing

            
              Me.MousePointer = vbDefault
              

End Sub


Sub PiList()

        
              Dim application As New CRAXDDRT.application
              Dim report As New CRAXDDRT.report
              Dim rsRpt As New ADODB.Recordset
                
              If rsRpt.State <> adStateClosed Then rsRpt.Close
              rsRpt.Open sqlprint, Db, adOpenDynamic, adLockOptimistic
            
              If rsRpt.EOF Then Me.MousePointer = vbDefault: Exit Sub
        
              Set report = application.OpenReport(App.path & "\Reports\rpt_pi_list.rpt")
              report.Database.Tables(1).SetDataSource rsRpt
                       
''#####################################################################
''# Qty Digit and decimal
report.FormulaFields(2).Text = "" & gi_decimalDigitQty & ""
report.FormulaFields(3).Text = "" & gi_decimalDigitQty & ""
''#####################################################################
               Select Case up_GetDateRange((dtMPList))

                Case 0:
                        report.Sections(4).Suppress = False
                        report.Sections(5).Suppress = True
                        report.Sections(6).Suppress = True

                 Case 1:

                        report.Sections(4).Suppress = True
                        report.Sections(5).Suppress = False
                        report.Sections(6).Suppress = True
                                                                       
                 Case 2:

                        report.Sections(4).Suppress = True
                        report.Sections(5).Suppress = True
                        report.Sections(6).Suppress = False
                                      
              End Select
              Dim dates As String
            
             
             report.FormulaFields(1).Text = "'" & datePiList & "'"
             report.ReportTitle = "Physical Inventory List ( Summary )"
            
    Dim vBinNumbers As Variant
    If combo1.ListIndex >= 0 Then
        vBinNumbers = GetBinNumbers
        report.SelectPrinter Label3, cbxPrinters.Text, Label2
        report.PaperSize = crPaperA4
        report.PaperOrientation = Orient
        report.PaperSource = vBinNumbers(combo1.ListIndex)
    Else
        MsgBox "No paper tray has been selected."
    End If

    Dim nCopy As Integer
    nCopy = Me.txtCopies
    Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
    Set report = Nothing
    rsRpt.Close
    Set rsRpt = Nothing

            
              Me.MousePointer = vbDefault
              

End Sub

Sub piReport()

              Dim application As New CRAXDDRT.application
              Dim report As New CRAXDDRT.report
              Dim rsRpt As New ADODB.Recordset
              Dim Rpt As New FrmRpt3
              Dim sqlControl As String, RsInvControl As New ADODB.Recordset
              
              Me.MousePointer = vbHourglass
            
              If rsRpt.State <> adStateClosed Then rsRpt.Close
              rsRpt.Open sqlprint, Db, adOpenDynamic, adLockOptimistic
            
              If rsRpt.EOF Then Me.MousePointer = vbDefault: Exit Sub

              Set report = application.OpenReport(App.path & "\Reports\rpt_pi_report.rpt")
              report.Database.Tables(1).SetDataSource rsRpt
                
''#####################################################################
''# Qty Digit and decimal
report.FormulaFields(5).Text = "" & gi_decimalDigitQty & ""
report.FormulaFields(6).Text = "" & gi_decimalDigitQty & ""
''#####################################################################
              
               Select Case up_GetDateRange((dtMPList))

                Case 0:
                        report.Sections(4).Suppress = False
                        report.Sections(5).Suppress = True
                        report.Sections(6).Suppress = True

                 Case 1:

                        report.Sections(4).Suppress = True
                        report.Sections(5).Suppress = False
                        report.Sections(6).Suppress = True
                                                                       
                 Case 2:

                        report.Sections(4).Suppress = True
                        report.Sections(5).Suppress = True
                        report.Sections(6).Suppress = False
                                                             
              End Select
         
            
             report.FormulaFields(1).Text = "'" & datePiList & "'"
             report.ReportTitle = "Inventory Report"
            
            
    Dim vBinNumbers As Variant
    If combo1.ListIndex >= 0 Then
        vBinNumbers = GetBinNumbers
        report.SelectPrinter Label3, cbxPrinters.Text, Label2
        report.PaperSize = crPaperA4
        report.PaperOrientation = Orient
        report.PaperSource = vBinNumbers(combo1.ListIndex)
    Else
        MsgBox "No paper tray has been selected."
    End If

    Dim nCopy As Integer
    nCopy = Me.txtCopies
    Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
    Set report = Nothing
    rsRpt.Close
    Set rsRpt = Nothing

    Me.MousePointer = vbDefault
        
End Sub

Sub piReportwh()

    Dim application As New CRAXDDRT.application
    Dim report As New CRAXDDRT.report
    Dim rsRpt As New ADODB.Recordset
    Dim rsrpt2 As New ADODB.Recordset
    Dim Rpt As New FrmRpt3
    Dim intDiffClosing As Integer
    
    Me.MousePointer = vbHourglass
    
    If rsRpt.State <> adStateClosed Then rsRpt.Close
    rsRpt.Open sqlprint, Db, adOpenDynamic, adLockOptimistic
    
    If rsRpt.EOF Then Me.MousePointer = vbDefault: Exit Sub
    
    Set report = application.OpenReport(App.path & "\Reports\rpt_pi_reportwh.rpt")
    report.Database.Tables(1).SetDataSource rsRpt
    
    intDiffClosing = up_GetDateRange(FrmValuationPriceReportWH.DMonth.Value)
    
    report.ReportTitle = "Valuation Price Report Per Warehouse"
    report.FormulaFields(1).Text = "'" & datePiList & "'"
    report.FormulaFields(5).Text = "" & gi_decimalDigitQty & ""
    report.FormulaFields(6).Text = "" & gi_decimalDigitAmountIDR & ""
    report.FormulaFields(11).Text = "" & intDiffClosing & ""
    
    If rsrpt2.State <> adStateClosed Then rsrpt2.Close
    rsrpt2.Open sqlprint2, Db, adOpenDynamic, adLockOptimistic
    
    report.OpenSubreport("summary").Database.Tables(1).SetDataSource rsrpt2
    report.OpenSubreport("summary").FormulaFields(1).Text = "" & intDiffClosing & ""
    
    Dim vBinNumbers As Variant
    If combo1.ListIndex >= 0 Then
        vBinNumbers = GetBinNumbers
        report.SelectPrinter Label3, cbxPrinters.Text, Label2
        report.PaperSize = crPaperA4
        report.PaperOrientation = Orient
        report.PaperSource = vBinNumbers(combo1.ListIndex)
    Else
        MsgBox "No paper tray has been selected."
    End If

    Dim nCopy As Integer
    nCopy = Me.txtCopies
    Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
    Set report = Nothing
    rsRpt.Close
    Set rsRpt = Nothing

    Me.MousePointer = vbDefault
        
End Sub

Sub ResinRequest()
    Dim application As New CRAXDDRT.application
    Dim report As New CRAXDDRT.report
    Dim rsRpt As New ADODB.Recordset
    
    If rsRpt.State <> adStateClosed Then rsRpt.Close
    rsRpt.Open sqlprint, Db, adOpenDynamic, adLockOptimistic
    
    If rsRpt.EOF Then Exit Sub
    
    Set report = application.OpenReport(App.path & "\Reports\rptresinreq.rpt")
    report.Database.Tables(1).SetDataSource rsRpt
    
    Dim vBinNumbers As Variant
    If combo1.ListIndex >= 0 Then
        vBinNumbers = GetBinNumbers
        report.SelectPrinter Label3, cbxPrinters.Text, Label2
        report.PaperSize = crPaperA4
        report.PaperOrientation = Orient
        report.PaperSource = vBinNumbers(combo1.ListIndex)
    Else
        MsgBox "No paper tray has been selected."
    End If

    Dim nCopy As Integer
    nCopy = Me.txtCopies
    Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
    Set report = Nothing
    rsRpt.Close
    Set rsRpt = Nothing
End Sub

Sub MaterialRequest()
    Dim application As New CRAXDDRT.application
    Dim report As New CRAXDDRT.report
    Dim rsRpt As New ADODB.Recordset
    
    If rsRpt.State <> adStateClosed Then rsRpt.Close
    rsRpt.Open sqlprint, Db, adOpenDynamic, adLockOptimistic
    
    If rsRpt.EOF Then Exit Sub
    
    Set report = application.OpenReport(App.path & "\Reports\rptmaterialreq.rpt")
    report.Database.Tables(1).SetDataSource rsRpt
    
    Dim vBinNumbers As Variant
    If combo1.ListIndex >= 0 Then
        vBinNumbers = GetBinNumbers
        report.SelectPrinter Label3, cbxPrinters.Text, Label2
        report.PaperSize = crPaperA4
        report.PaperOrientation = Orient
        report.PaperSource = vBinNumbers(combo1.ListIndex)
    Else
        MsgBox "No paper tray has been selected."
    End If

    Dim nCopy As Integer
    nCopy = Me.txtCopies
    Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
    Set report = Nothing
    rsRpt.Close
    Set rsRpt = Nothing
End Sub
Sub MaterialConsumption()
  Dim application As New CRAXDDRT.application
  Dim report As New CRAXDDRT.report
  Dim rsRpt As New ADODB.Recordset
  
  If rsRpt.State <> adStateClosed Then rsRpt.Close
  rsRpt.Open sqlprint, Db, adOpenDynamic, adLockOptimistic
  
  If rsRpt.EOF Then Exit Sub
  
    Set report = application.OpenReport(App.path & "\Reports\rptmaterialconsumption.rpt")
    report.Database.Tables(1).SetDataSource rsRpt
    
    report.FormulaFields(34).Text = "'" & sqlprint2 & "'"
    report.FormulaFields(35).Text = "'" & F_Factory & "'"
    report.ReportTitle = "Material Consumption Report"
         
  
    Dim vBinNumbers As Variant
    If combo1.ListIndex >= 0 Then
        vBinNumbers = GetBinNumbers
        report.SelectPrinter Label3, cbxPrinters.Text, Label2
        report.PaperSize = crPaperA3
        report.PaperOrientation = crLandscape
        report.PaperSource = vBinNumbers(combo1.ListIndex)
    Else
        MsgBox "No paper tray has been selected."
    End If

    Dim nCopy As Integer
    nCopy = Me.txtCopies
    Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
    Set report = Nothing
    rsRpt.Close
    Set rsRpt = Nothing
End Sub

Sub AccPay()
    Dim application As New CRAXDDRT.application
    Dim report As New CRAXDDRT.report
    Dim rsRpt As New ADODB.Recordset
    
    If rsRpt.State <> adStateClosed Then rsRpt.Close
    rsRpt.Open sqlprint, Db, adOpenDynamic, adLockOptimistic
    
    If rsRpt.EOF Then Exit Sub
    
    Set report = application.OpenReport(App.path & "\Reports\rpt_Accpay.rpt")
    report.Database.Tables(1).SetDataSource rsRpt
    
    report.FormulaFields(1).Text = "'" & Format(frm_accPay.DMonth, "dd-MMM-yyyy") & " to " & Format(frm_accPay.Dmonth2, "dd-MMM-yyyy") & "'"
              
'    report.FormulaFields(9).Text = "" & gi_decimalDigitQty & ""
'    report.FormulaFields(10).Text = "" & gi_decimalDigitQty & ""
'    report.FormulaFields(11).Text = "" & gi_decimalDigitPrice & ""
'    report.FormulaFields(12).Text = "" & gi_decimalDigitPrice & ""
'    report.FormulaFields(13).Text = "" & gi_decimalDigitAmount & ""
'    report.FormulaFields(14).Text = "" & gi_decimalDigitAmount & "x"
    
    Dim vBinNumbers As Variant
    If combo1.ListIndex >= 0 Then
        vBinNumbers = GetBinNumbers
        report.SelectPrinter Label3, cbxPrinters.Text, Label2
        report.PaperSize = crPaperA4
        report.PaperOrientation = Orient
        report.PaperSource = vBinNumbers(combo1.ListIndex)
    Else
        MsgBox "No paper tray has been selected."
    End If
    
    Dim nCopy As Integer
    nCopy = Me.txtCopies
    
    Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
    Set report = Nothing
    rsRpt.Close
    Set rsRpt = Nothing
End Sub

Private Sub RecSupInquiry()

              Dim application As New CRAXDDRT.application
              Dim report As New CRAXDDRT.report
              Dim rsRpt As New ADODB.Recordset
                            
  
              Me.MousePointer = vbHourglass
                
                                    
              If rsRpt.State <> adStateClosed Then rsRpt.Close
              rsRpt.Open sqlprint, Db, adOpenDynamic, adLockOptimistic
            
              If rsRpt.EOF Then Me.MousePointer = vbDefault: Exit Sub
            
              Set report = application.OpenReport(App.path & "\Reports\rptRecSupInq.rpt")
              report.Database.Tables(1).SetDataSource rsRpt
              
              
              '#RecordSet COmpany Profile
              Dim rsCP As New ADODB.Recordset
              If rsCP.State <> adStateClosed Then rsCP.Close
              rsCP.Open "select * from company_profile", Db, adOpenKeyset, adLockOptimistic
              report.Database.Tables(2).SetDataSource rsCP
                                     
              Dim dates As String
                         
             report.FormulaFields(4).Text = "'" & datePiList & "'"
             report.ReportTitle = "Receipt / Supply Inquiry"
             
            '#Set Data Header (Stock Master)
             report.FormulaFields(5).Text = "'" & MonthPre & "'" '#Pre Month
             report.FormulaFields(6).Text = "'" & MonthReceipt & "'" '#Receipt
             report.FormulaFields(7).Text = "'" & MonthSupply & "'" '#Supply
             report.FormulaFields(8).Text = "'" & MonthCurrent & "'" '#Current
             
            
    Dim vBinNumbers As Variant
    If combo1.ListIndex >= 0 Then
        vBinNumbers = GetBinNumbers
        report.SelectPrinter Label3, cbxPrinters.Text, Label2
        report.PaperSize = crPaperA4
        report.PaperOrientation = Orient
        report.PaperSource = vBinNumbers(combo1.ListIndex)
    Else
        MsgBox "No paper tray has been selected."
    End If

    Dim nCopy As Integer
    nCopy = Me.txtCopies
    Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
    Set report = Nothing
    rsRpt.Close
    Set rsRpt = Nothing
            
              Me.MousePointer = vbDefault
End Sub

Sub PurchasingMaterial()
  Dim application As New CRAXDDRT.application
  Dim report As New CRAXDDRT.report
  Dim rsRpt As New ADODB.Recordset
  
  If rsRpt.State <> adStateClosed Then rsRpt.Close
  rsRpt.Open sqlprint, Db, adOpenDynamic, adLockOptimistic
  
  If rsRpt.EOF Then Exit Sub
  
  Set report = application.OpenReport(App.path & "\Reports\rptPurchasingMat.rpt")
  report.Database.Tables(1).SetDataSource rsRpt
  report.FormulaFields(2).Text = "'" & Format(Fbulan, "MMMM yyyy") & "'"
  
    Dim vBinNumbers As Variant
    If combo1.ListIndex >= 0 Then
        vBinNumbers = GetBinNumbers
        report.SelectPrinter Label3, cbxPrinters.Text, Label2
        report.PaperSize = crPaperA3
        report.PaperOrientation = Orient
        report.PaperSource = vBinNumbers(combo1.ListIndex)
    Else
        MsgBox "No paper tray has been selected."
    End If

    Dim nCopy As Integer
    nCopy = Me.txtCopies
    Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
    Set report = Nothing
    rsRpt.Close
    Set rsRpt = Nothing
End Sub
Sub itemMaster()
  Dim application As New CRAXDDRT.application
  Dim report As New CRAXDDRT.report
  Dim rsRpt As New ADODB.Recordset
  Dim rssubrpt As New ADODB.Recordset

  If rsRpt.State <> adStateClosed Then rsRpt.Close
  rsRpt.Open sqlprint, Db, adOpenDynamic, adLockOptimistic
  
  If rsRpt.EOF Then Exit Sub
  Set report = application.OpenReport(App.path & "\Reports\rpt_itemMaster2.rpt")
  report.Database.Tables(1).SetDataSource rsRpt
  
  If rssubrpt.State <> adStateClosed Then rssubrpt.Close
  rssubrpt.Open sqlprint2, Db, adOpenKeyset, adLockOptimistic
  report.OpenSubreport("SubRptItemMaster").Database.Tables(1).SetDataSource rssubrpt
  
  report.ReportTitle = "Item Master"

    Dim vBinNumbers As Variant
    If combo1.ListIndex >= 0 Then
        vBinNumbers = GetBinNumbers
        report.SelectPrinter Label3, cbxPrinters.Text, Label2
        report.PaperSize = crPaperA4
        report.PaperOrientation = Orient
        report.PaperSource = vBinNumbers(combo1.ListIndex)
    Else
        MsgBox "No paper tray has been selected."
    End If

    Dim nCopy As Integer
    nCopy = Me.txtCopies
    Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
    Set report = Nothing
    Set rsRpt = Nothing
    Set rssubrpt = Nothing
  
End Sub


Private Sub rawMaterial()

              Dim application As New CRAXDDRT.application
              Dim report As New CRAXDDRT.report
              Dim rsRpt As New ADODB.Recordset
                                             
              If rsRpt.State <> adStateClosed Then rsRpt.Close
              rsRpt.Open sqlprint, Db, adOpenDynamic, adLockOptimistic
            
              If rsRpt.EOF Then Exit Sub
            
              Set report = application.OpenReport(App.path & "\Reports\rptrawMaterialStock.rpt")
              report.Database.Tables(1).SetDataSource rsRpt
              
              '#RecordSet COmpany Profile
              Dim rsCP As New ADODB.Recordset
              If rsCP.State <> adStateClosed Then rsCP.Close
              rsCP.Open "select * from company_profile", Db, adOpenKeyset, adLockOptimistic
              report.Database.Tables(2).SetDataSource rsCP
                       
              Dim dates As String
                         
             report.FormulaFields(8).Text = "'" & datePiList & "'"
             report.ReportTitle = "Raw Material Stock"
            

    Dim vBinNumbers As Variant
    If combo1.ListIndex >= 0 Then
        vBinNumbers = GetBinNumbers
        report.SelectPrinter Label3, cbxPrinters.Text, Label2
        report.PaperSize = crPaperA4
        report.PaperOrientation = Orient
        report.PaperSource = vBinNumbers(combo1.ListIndex)
    Else
        MsgBox "No paper tray has been selected."
    End If

    Dim nCopy As Integer
    nCopy = Me.txtCopies
    Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
    Set report = Nothing
    rsRpt.Close
    Set rsRpt = Nothing
    
End Sub


Sub Formula()
 
Dim application As New CRAXDDRT.application
Dim report As New CRAXDDRT.report
Dim rdb As CRAXDDRT.Database
Dim rdbtables As CRAXDDRT.DatabaseTables
Dim rdbtable As CRAXDDRT.DatabaseTable
Dim rsections As CRAXDDRT.Sections
Dim rsection As CRAXDDRT.Section
Dim robjs As CRAXDDRT.ReportObjects
Dim robj As CRAXDDRT.SubreportObject
Dim sreport As CRAXDDRT.report
Dim Y As Long, i As Long
Dim rs1 As Recordset, rs2 As Recordset
    
    Set rs1 = New Recordset
    rs1.Open sqlprint, Db, adOpenDynamic, adLockOptimistic
    If Not rs1.EOF Then
        Set rs2 = New Recordset
        rs2.Open sqlprint2, Db, adOpenDynamic, adLockOptimistic
        
    If Not rs2.EOF Then
        Set report = application.OpenReport(App.path & "\REPORTs\Load.rpt")
        report.Database.Tables(1).SetDataSource rs1
        Set rdb = report.Database
        Set rdbtables = rdb.Tables
        Set rdbtable = rdbtables.Item(1)
        rdbtable.SetDataSource rs1, 3
        
        Set rsections = report.Sections
        For i = 1 To rsections.Count
            Set rsection = rsections.Item(i)
            Set robjs = rsection.ReportObjects
            For Y = 1 To robjs.Count
                If robjs.Item(Y).Kind = crSubreportObject Then
                    Set robj = robjs.Item(Y)
                    Set sreport = robj.OpenSubreport
                    Set rdb = sreport.Database
                    Set rdbtables = rdb.Tables
                    Set rdbtable = rdbtables.Item(1)
                    rdbtable.SetDataSource rs2, 3
                End If
            Next
        Next
        
        Dim vBinNumbers As Variant
        If combo1.ListIndex >= 0 Then
            vBinNumbers = GetBinNumbers
    
            report.SelectPrinter Label3, cbxPrinters.Text, Label2
            report.PaperOrientation = Orient
            report.PaperSource = vBinNumbers(combo1.ListIndex)
            report.PaperSize = crPaperA4
        Else
            MsgBox "No paper tray has been selected."
        End If

        Dim nCopy As Integer
        nCopy = Me.txtCopies
        Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
        Set report = Nothing

    End If
    End If
End Sub

Sub ExchangeList()
 Dim application As New CRAXDDRT.application
  Dim report As New CRAXDDRT.report
  Dim rsRpt As New ADODB.Recordset

  If rsRpt.State <> adStateClosed Then rsRpt.Close
  rsRpt.Open sqlprint, Db, adOpenDynamic, adLockOptimistic
  
  Set report = application.OpenReport(App.path & "\Reports\list_NoRate.rpt")
  report.Database.Tables(1).SetDataSource rsRpt
  
    report.FormulaFields(1).Text = "'" & Fbulan & "'"
    
    Dim vBinNumbers As Variant
    If combo1.ListIndex >= 0 Then
        vBinNumbers = GetBinNumbers
        report.SelectPrinter Label3, cbxPrinters.Text, Label2
        report.PaperSize = crPaperA4
        report.PaperOrientation = Orient
        report.PaperSource = vBinNumbers(combo1.ListIndex)
    Else
        MsgBox "No paper tray has been selected."
    End If

    Dim nCopy As Integer
    nCopy = Me.txtCopies
    Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
    Set report = Nothing
    rsRpt.Close
    Set rsRpt = Nothing
End Sub

Sub pocust()
  Dim application As New CRAXDDRT.application
  Dim report As New CRAXDDRT.report
  Dim rsRpt As New ADODB.Recordset

  If rsRpt.State <> adStateClosed Then rsRpt.Close
  rsRpt.Open sqlprint, Db, adOpenDynamic, adLockOptimistic

  If rsRpt.EOF Then Exit Sub

  Set report = application.OpenReport(App.path & "\Reports\Rpt_PO_Cust.rpt")
  report.Database.Tables(1).SetDataSource rsRpt

    report.FormulaFields(2).Text = "'Period " & Format(Fbulan, "MMMM yyyy") & "'"
    report.FormulaFields(4).Text = "'" & Ftahun & "'"

    Dim vBinNumbers As Variant
    If combo1.ListIndex >= 0 Then
        vBinNumbers = GetBinNumbers
        report.SelectPrinter Label3, cbxPrinters.Text, Label2
        report.PaperSize = crPaperA4
        report.PaperOrientation = Orient
        report.PaperSource = vBinNumbers(combo1.ListIndex)
    Else
        MsgBox "No paper tray has been selected."
    End If

    Dim nCopy As Integer
    nCopy = Me.txtCopies
    Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
    Set report = Nothing
    rsRpt.Close
    Set rsRpt = Nothing
End Sub

Private Sub RecSupScheduleInquiry()


Dim application As New CRAXDDRT.application
Dim report As New CRAXDDRT.report
Dim rsRpt As New ADODB.Recordset
Dim Rpt As New FrmRpt3
 

If rsRpt.State <> adStateClosed Then rsRpt.Close
rsRpt.Open sqlprint, Db, adOpenDynamic, adLockOptimistic


Set report = application.OpenReport(App.path & "\Reports\rpt_recSupSchedule.rpt")
report.Database.Tables(1).SetDataSource rsRpt

'#RecordSet COmpany Profile
Dim rsCP As New ADODB.Recordset
If rsCP.State <> adStateClosed Then rsCP.Close
rsCP.Open "select isnull(company_name,'')company_name,isnull(address1,'')address1,isnull(address2,'')address2,isnull(phone1,'')phone1,isnull(fax,'')fax from company_profile", Db, adOpenKeyset, adLockOptimistic
 
If rsCP.EOF = False Then
    report.FormulaFields(23).Text = "'" & Trim(rsCP!company_name) & "'" 'companyName
    report.FormulaFields(24).Text = "'" & Trim(rsCP!address1) & "'" 'address1
    report.FormulaFields(25).Text = "'" & Trim(rsCP!address2) & "'" 'address2
    report.FormulaFields(26).Text = "'" & IIf(Trim(rsCP!phone1) = "", "", "Telp : " & Trim(rsCP!phone1)) & "'" 'telp
    report.FormulaFields(27).Text = "'" & IIf(Trim(rsCP!fax) = "", "", "Fax : " & Trim(rsCP!fax)) & "'"  'fax
End If

''#####################################################################
''# Qty Digit and decimal
report.FormulaFields(28).Text = "" & gi_decimalDigitQty & ""
report.FormulaFields(29).Text = "" & gi_decimalDigitQty & ""
report.FormulaFields(30).Text = "" & gi_decimalDigitPrice & ""
report.FormulaFields(31).Text = "" & gi_decimalDigitPrice & ""
report.FormulaFields(32).Text = "" & gi_decimalDigitAmount & ""
report.FormulaFields(33).Text = "" & gi_decimalDigitAmount & ""
''#####################################################################

report.FormulaFields(16).Text = "'" & dtMPList & "'" 'Item Description
report.FormulaFields(17).Text = "'" & datePiList & "'"  'Period
report.FormulaFields(5).Text = "" & gd_StockMaster & "" 'Stock
report.FormulaFields(22).Text = "'" & gs_unitDesc & "'" 'UnitDesc
report.ReportTitle = "Receipt Supply Schedule Inquiry"
  

    Dim vBinNumbers As Variant
    If combo1.ListIndex >= 0 Then
        vBinNumbers = GetBinNumbers
        report.SelectPrinter Label3, cbxPrinters.Text, Label2
        report.PaperSize = crPaperA4
        report.PaperOrientation = Orient
        report.PaperSource = vBinNumbers(combo1.ListIndex)
    Else
        MsgBox "No paper tray has been selected."
    End If

    Dim nCopy As Integer
    nCopy = Me.txtCopies
    Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
    Set report = Nothing
    rsRpt.Close
    Set rsRpt = Nothing

End Sub

Sub RptRequestAuto()

    Dim application As New CRAXDDRT.application
    Dim report As New CRAXDDRT.report
    Dim rsRpt As New ADODB.Recordset
    
    If rsRpt.State <> adStateClosed Then rsRpt.Close
    rsRpt.Open sqlprint, Db, adOpenDynamic, adLockOptimistic
    
    If rsRpt.EOF Then Exit Sub
    
    Set report = application.OpenReport(App.path & "\Reports\RequestAuto.rpt")
    report.Database.Tables(1).SetDataSource rsRpt
    report.FormulaFields(1).Text = "" & gi_decimalDigitQtyBOM & ""
    
    Dim vBinNumbers As Variant
    
    If combo1.ListIndex >= 0 Then
        vBinNumbers = GetBinNumbers
        report.SelectPrinter Label3, cbxPrinters.Text, Label2
        report.PaperSize = crPaperA4
        report.PaperOrientation = crLandscape
        report.PaperSource = vBinNumbers(combo1.ListIndex)
    Else
        MsgBox "No paper tray has been selected."
    End If
    
    Dim nCopy As Integer
    nCopy = Me.txtCopies
    Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
    Set report = Nothing
    rsRpt.Close
    Set rsRpt = Nothing

End Sub

Sub AlarmlistOP()
  Dim application As New CRAXDDRT.application
  Dim report As New CRAXDDRT.report
  Dim rsRpt As New ADODB.Recordset

  If rsRpt.State <> adStateClosed Then rsRpt.Close
  rsRpt.Open sqlprint, Db, adOpenDynamic, adLockOptimistic

  If rsRpt.EOF Then Exit Sub

  Set report = application.OpenReport(App.path & "\Reports\RptAlarmlistOP.rpt")
  report.Database.Tables(1).SetDataSource rsRpt
  report.FormulaFields(2).Text = "'" & Format(FrmAlarmList.dtpDate, "dd MMMM  yyyy") & "'"
'#####################################################################
'# Qty Digit and decimal
report.FormulaFields(4).Text = "" & gi_decimalDigitQty & ""
report.FormulaFields(5).Text = "" & gi_decimalDigitQty & ""
'#####################################################################
    Dim vBinNumbers As Variant
    If combo1.ListIndex >= 0 Then
        vBinNumbers = GetBinNumbers
        report.SelectPrinter Label3, cbxPrinters.Text, Label2
        report.PaperSize = crPaperA4
        report.PaperOrientation = Orient
        report.PaperSource = vBinNumbers(combo1.ListIndex)
    Else
        MsgBox "No paper tray has been selected."
    End If

    Dim nCopy As Integer
    nCopy = Me.txtCopies
    Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
    Set report = Nothing
    rsRpt.Close
    Set rsRpt = Nothing
End Sub


Sub Alarmlist()
  Dim application As New CRAXDDRT.application
  Dim report As New CRAXDDRT.report
  Dim rsRpt As New ADODB.Recordset

  If rsRpt.State <> adStateClosed Then rsRpt.Close
  rsRpt.Open sqlprint, Db, adOpenDynamic, adLockOptimistic

  If rsRpt.EOF Then Exit Sub

  Set report = application.OpenReport(App.path & "\Reports\RptAlarmlist.rpt")
  report.Database.Tables(1).SetDataSource rsRpt
  report.FormulaFields(2).Text = "'" & Format(FrmAlarmList.dtpDate, "dd MMMM  yyyy") & "'"
'#####################################################################
'# Qty Digit and decimal
report.FormulaFields(4).Text = "" & gi_decimalDigitQty & ""
report.FormulaFields(5).Text = "" & gi_decimalDigitQty & ""
'#####################################################################
    Dim vBinNumbers As Variant
    If combo1.ListIndex >= 0 Then
        vBinNumbers = GetBinNumbers
        report.SelectPrinter Label3, cbxPrinters.Text, Label2
        report.PaperSize = crPaperA4
        report.PaperOrientation = Orient
        report.PaperSource = vBinNumbers(combo1.ListIndex)
    Else
        MsgBox "No paper tray has been selected."
    End If

    Dim nCopy As Integer
    nCopy = Me.txtCopies
    Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
    Set report = Nothing
    rsRpt.Close
    Set rsRpt = Nothing
End Sub

Sub ForecastPart()
  Dim application As New CRAXDDRT.application
  Dim report As New CRAXDDRT.report
  Dim rsRpt As New ADODB.Recordset

  If rsRpt.State <> adStateClosed Then rsRpt.Close
  rsRpt.Open sqlprint, Db, adOpenDynamic, adLockOptimistic

  If rsRpt.EOF Then Exit Sub
    
  Set report = application.OpenReport(App.path & "\Reports\Rpt_forecastparts.rpt")
    report.Database.Tables(1).SetDataSource rsRpt
    
    report.FormulaFields(6).Text = "'" & Format(DateAdd("m", 0, Format(Fbulan, "YYYY-MM-DD")), "MMM  YYYY") & "'"
    report.FormulaFields(5).Text = "'" & Format(DateAdd("m", 1, Format(Fbulan, "YYYY-MM-DD")), "MMM  YYYY") & "'"
    report.FormulaFields(4).Text = "'" & Format(DateAdd("m", 2, Format(Fbulan, "YYYY-MM-DD")), "MMM  YYYY") & "'"
    report.FormulaFields(3).Text = "'" & Format(DateAdd("m", 3, Format(Fbulan, "YYYY-MM-DD")), "MMM  YYYY") & "'"
    report.FormulaFields(2).Text = "'" & Format(DateAdd("m", 4, Format(Fbulan, "YYYY-MM-DD")), "MMM  YYYY") & "'"
    report.FormulaFields(1).Text = "'" & Format(DateAdd("m", 5, Format(Fbulan, "YYYY-MM-DD")), "MMM  YYYY") & "'"
    '
    report.FormulaFields(13).Text = "'" & Format(Fbulan, "MMMM YYYY") & " to " & Format(Ftahun, "MMMM YYYY") & "'"  'Periode
    report.FormulaFields(14).Text = "'" & Trim$(F_Factory) & "'"   'Supplier Code
    report.FormulaFields(15).Text = "'" & F_Cust_Name & "'"  'Nama Perusahaan
        '#####################################################################
        '# Qty Digit and decimal
        report.FormulaFields(16).Text = "" & gi_decimalDigitQty & ""
        report.FormulaFields(17).Text = "" & gi_decimalDigitQty & ""
        '#####################################################################

    Dim vBinNumbers As Variant
    If combo1.ListIndex >= 0 Then
        vBinNumbers = GetBinNumbers
        report.SelectPrinter Label3, cbxPrinters.Text, Label2
        report.PaperSize = crPaperA4
        report.PaperOrientation = Orient
        report.PaperSource = vBinNumbers(combo1.ListIndex)
    Else
        MsgBox "No paper tray has been selected."
    End If

    Dim nCopy As Integer
    nCopy = Me.txtCopies
    Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
    Set report = Nothing
    rsRpt.Close
    Set rsRpt = Nothing
End Sub

Sub ForecastMaterial()
  Dim application As New CRAXDDRT.application
  Dim report As New CRAXDDRT.report
  Dim rsRpt As New ADODB.Recordset

  If rsRpt.State <> adStateClosed Then rsRpt.Close
  rsRpt.Open sqlprint, Db, adOpenDynamic, adLockOptimistic

  If rsRpt.EOF Then Exit Sub
    
  Set report = application.OpenReport(App.path & "\Reports\Rpt_forecastmaterial.rpt")
  report.Database.Tables(1).SetDataSource rsRpt
  
    report.FormulaFields(6).Text = "'" & Format(DateAdd("m", 0, Format(Fbulan, "YYYY-MM-DD")), "MMM  YYYY") & "'"
    report.FormulaFields(5).Text = "'" & Format(DateAdd("m", 1, Format(Fbulan, "YYYY-MM-DD")), "MMM  YYYY") & "'"
    report.FormulaFields(4).Text = "'" & Format(DateAdd("m", 2, Format(Fbulan, "YYYY-MM-DD")), "MMM  YYYY") & "'"
    report.FormulaFields(3).Text = "'" & Format(DateAdd("m", 3, Format(Fbulan, "YYYY-MM-DD")), "MMM  YYYY") & "'"
    report.FormulaFields(2).Text = "'" & Format(DateAdd("m", 4, Format(Fbulan, "YYYY-MM-DD")), "MMM  YYYY") & "'"
    report.FormulaFields(1).Text = "'" & Format(DateAdd("m", 5, Format(Fbulan, "YYYY-MM-DD")), "MMM  YYYY") & "'"
    '
    report.FormulaFields(13).Text = "'" & Format(Fbulan, "MMMM YYYY") & " to " & Format(Ftahun, "MMMM YYYY") & "'"  'Periode
    report.FormulaFields(14).Text = "'" & Trim$(F_Factory) & "'"   'Supplier Code
    report.FormulaFields(15).Text = "'" & F_Cust_Name & "'"  'Nama Perusahaan
        '#####################################################################
        '# Qty Digit and decimal
        report.FormulaFields(16).Text = "" & gi_decimalDigitQty & ""
        report.FormulaFields(17).Text = "" & gi_decimalDigitQty & ""
        report.FormulaFields(18).Text = "" & gi_decimalDigitThickness & ""
        report.FormulaFields(19).Text = "" & gi_decimalDigitThickness & ""
        report.FormulaFields(20).Text = "" & gi_decimalDigitWidth & ""
        report.FormulaFields(21).Text = "" & gi_decimalDigitWidth & ""
        report.FormulaFields(22).Text = "" & gi_decimalDigitLength & ""
        report.FormulaFields(23).Text = "" & gi_decimalDigitLength & ""
        '#####################################################################
    Dim vBinNumbers As Variant
    If combo1.ListIndex >= 0 Then
        vBinNumbers = GetBinNumbers
        report.SelectPrinter Label3, cbxPrinters.Text, Label2
        report.PaperSize = crPaperA4
        report.PaperOrientation = Orient
        report.PaperSource = vBinNumbers(combo1.ListIndex)
    Else
        MsgBox "No paper tray has been selected."
    End If

    Dim nCopy As Integer
    nCopy = Me.txtCopies
    Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
    Set report = Nothing
    rsRpt.Close
    Set rsRpt = Nothing
End Sub


Private Sub WIP()

Dim lrs_sql As New ADODB.Recordset
Dim application As New CRAXDDRT.application
Dim report As New CRAXDDRT.report

Me.MousePointer = vbHourglass

If lrs_sql.State <> adStateClosed Then lrs_sql.Close
lrs_sql.Open sqlprint, Db, adOpenKeyset, adLockOptimistic
                                

Set report = application.OpenReport(App.path & "\Reports\rpt_WIP.rpt")
report.Database.Tables(1).SetDataSource lrs_sql

Dim lrs_companyProfile As New ADODB.Recordset
If lrs_companyProfile.State <> adStateClosed Then lrs_companyProfile.Close
lrs_companyProfile.Open "select * from company_profile", Db, adOpenKeyset, adLockOptimistic

If lrs_companyProfile.EOF = False Then
    report.FormulaFields(5).Text = "'" & Trim(lrs_companyProfile!company_name) & "'"
    report.FormulaFields(6).Text = "'" & Trim(lrs_companyProfile!address1) & "'"
    report.FormulaFields(7).Text = "'" & Trim(lrs_companyProfile!address2) & "'"
    report.FormulaFields(8).Text = "'" & Trim(lrs_companyProfile!City) & "'"
    report.FormulaFields(9).Text = "'" & Trim(lrs_companyProfile!postal_code) & "'"
    report.FormulaFields(10).Text = "'" & Trim(lrs_companyProfile!phone1) & "'"
    report.FormulaFields(11).Text = "'" & Trim(lrs_companyProfile!fax) & "'"
End If

report.FormulaFields(2).Text = "'" & dtMPList & "'"
report.FormulaFields(3).Text = "'" & datePiList & "'"
report.FormulaFields(4).Text = "'" & tglAwalRptPrint & "'"
report.ReportTitle = "WIP Report"

Dim vBinNumbers As Variant
If combo1.ListIndex >= 0 Then
    vBinNumbers = GetBinNumbers
    report.SelectPrinter Label3, cbxPrinters.Text, Label2
    report.PaperSize = crPaperA4
    report.PaperOrientation = Orient
    report.PaperSource = vBinNumbers(combo1.ListIndex)
Else
    MsgBox "No paper tray has been selected."
End If

Dim nCopy As Integer
nCopy = Me.txtCopies
Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
Set report = Nothing
lrs_companyProfile.Close
lrs_sql.Close

End Sub


Private Sub MatrixDaily()

Dim lrs_sql As New ADODB.Recordset
Dim application As New CRAXDDRT.application
Dim report As New CRAXDDRT.report

Me.MousePointer = vbHourglass

If lrs_sql.State <> adStateClosed Then lrs_sql.Close
lrs_sql.Open sqlprint, Db, adOpenKeyset, adLockOptimistic
                                

Set report = application.OpenReport(App.path & "\Reports\MatrixDaily.rpt")
report.Database.Tables(1).SetDataSource lrs_sql

report.FormulaFields(1).Text = "'" & tglAwalRptPrint & "'"

Dim zi As Byte
For zi = 0 To 30
    report.FormulaFields(zi + 2).Text = "'" & xdays(zi) & "'"
Next

'#####################################################################
'# Qty Digit and decimal
report.FormulaFields(33).Text = "" & gi_decimalDigitQty & ""
report.FormulaFields(34).Text = "" & gi_decimalDigitQty & ""
'#####################################################################

Dim vBinNumbers As Variant
If combo1.ListIndex >= 0 Then
    vBinNumbers = GetBinNumbers
    report.SelectPrinter Label3, cbxPrinters.Text, Label2
    report.PaperSize = crPaperA4
    report.PaperOrientation = Orient
    report.PaperSource = vBinNumbers(combo1.ListIndex)
Else
    MsgBox "No paper tray has been selected."
End If

Dim nCopy As Integer
nCopy = Me.txtCopies
Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
Set report = Nothing
'lrs_companyProfile.Close
lrs_sql.Close
Set lrs_sql = Nothing
End Sub

Private Sub rptWorksheet()

    Dim lrs_sql As New ADODB.Recordset
    Dim application As New CRAXDDRT.application
    Dim report As New CRAXDDRT.report
    
    Me.MousePointer = vbHourglass
    
    If lrs_sql.State <> adStateClosed Then lrs_sql.Close
    lrs_sql.Open sqlprint, Db, adOpenKeyset, adLockOptimistic
                                    
    
    Set report = application.OpenReport(App.path & "\Reports\rptWorksheet.rpt")
    report.Database.Tables(1).SetDataSource lrs_sql
    
    '#####################################################################
    '# Qty Digit and decimal
    report.FormulaFields(2).Text = "" & gi_decimalDigitQty & ""
    report.FormulaFields(3).Text = "" & gi_decimalDigitQty & ""
    report.FormulaFields(4).Text = "" & gi_decimalDigitBox & ""
    report.FormulaFields(5).Text = "" & gi_decimalDigitBox & ""
    '#####################################################################
    
    Dim vBinNumbers As Variant
    If combo1.ListIndex >= 0 Then
        vBinNumbers = GetBinNumbers
        report.SelectPrinter Label3, cbxPrinters.Text, Label2
        report.PaperSize = crPaperA4
        report.PaperOrientation = Orient
        report.PaperSource = vBinNumbers(combo1.ListIndex)
    Else
        MsgBox "No paper tray has been selected."
    End If
    
    Dim nCopy As Integer
    nCopy = Me.txtCopies
    Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
    Set report = Nothing
    lrs_sql.Close
    Set lrs_sql = Nothing
End Sub

Private Sub rptProdResultInquiry2()

    Dim lrs_sql As New ADODB.Recordset
    Dim application As New CRAXDDRT.application
    Dim report As New CRAXDDRT.report
    
    Me.MousePointer = vbHourglass
    If lrs_sql.State <> adStateClosed Then lrs_sql.Close
    lrs_sql.Open sqlprint, Db, adOpenKeyset, adLockOptimistic
                                        
    Set report = application.OpenReport(App.path & "\Reports\rptProdResultInquiry.rpt")
    report.Database.Tables(1).SetDataSource lrs_sql
    
    report.FormulaFields(1).Text = "'" & tglAwalRptPrint & "'"
    report.FormulaFields(2).Text = "'" & tglAkhirRptPrint & "'"

    '#####################################################################
    '# Qty Digit and decimal
    report.FormulaFields(4).Text = "" & gi_decimalDigitQty & ""
    report.FormulaFields(5).Text = "" & gi_decimalDigitQty & ""
    report.FormulaFields(6).Text = "" & gi_decimalDigitBox & ""
    report.FormulaFields(7).Text = "" & gi_decimalDigitBox & ""
    '#####################################################################
    
    Dim vBinNumbers As Variant
    If combo1.ListIndex >= 0 Then
        vBinNumbers = GetBinNumbers
        report.SelectPrinter Label3, cbxPrinters.Text, Label2
        report.PaperSize = crPaperA4
        report.PaperOrientation = Orient
        report.PaperSource = vBinNumbers(combo1.ListIndex)
    Else
        MsgBox "No paper tray has been selected."
    End If
    
    Dim nCopy As Integer
    nCopy = Me.txtCopies
    Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
    Set report = Nothing
    lrs_sql.Close
    Set lrs_sql = Nothing
End Sub

Private Sub MatrixDaily22()

Dim lrs_sql As New ADODB.Recordset
Dim application As New CRAXDDRT.application
Dim report As New CRAXDDRT.report
Dim zi As Byte
Me.MousePointer = vbHourglass

If lrs_sql.State <> adStateClosed Then lrs_sql.Close
lrs_sql.Open sqlprint, Db, adOpenKeyset, adLockOptimistic
                                

Set report = application.OpenReport(App.path & "\Reports\MatrixDaily22.rpt")
report.Database.Tables(1).SetDataSource lrs_sql

report.FormulaFields(1).Text = "'" & tglAwalRptPrint & "'"
For zi = 0 To 21
    report.FormulaFields(zi + 2).Text = "'" & xdays(zi) & "'"
Next

'#####################################################################
'# Qty Digit and decimal
report.FormulaFields(24).Text = "" & gi_decimalDigitQty & ""
report.FormulaFields(25).Text = "" & gi_decimalDigitQty & ""
'#####################################################################

Dim vBinNumbers As Variant
If combo1.ListIndex >= 0 Then
    vBinNumbers = GetBinNumbers
    report.SelectPrinter Label3, cbxPrinters.Text, Label2
    report.PaperSize = crPaperA4
    report.PaperOrientation = Orient
    report.PaperSource = vBinNumbers(combo1.ListIndex)
Else
    MsgBox "No paper tray has been selected."
End If

Dim nCopy As Integer
nCopy = Me.txtCopies
Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
Set report = Nothing
'lrs_companyProfile.Close
lrs_sql.Close
Set lrs_sql = Nothing
End Sub

Private Sub SupplyList()
Dim application As New CRAXDDRT.application
Dim report As New CRAXDDRT.report
Dim rsRpt As New ADODB.Recordset

Me.MousePointer = vbHourglass

If rsRpt.State <> adStateClosed Then rsRpt.Close
rsRpt.Open sqlprint, Db, adOpenDynamic, adLockOptimistic
         
Set report = application.OpenReport(App.path & "\Reports\rptSupplyList.rpt")

report.Database.Tables(1).SetDataSource rsRpt
report.FormulaFields(2).Text = "'" & Format(frmRptSupplyList.dtAwal.Value, "dd MMM yyyy") & "'"
report.FormulaFields(3).Text = "'" & Format(frmRptSupplyList.dtAkhir.Value, "dd MMM yyyy") & "'"
report.FormulaFields(5).Text = "'" & frmRptSupplyList.cbo_towarehouse.Column(0) & "'"
report.FormulaFields(6).Text = "'" & frmRptSupplyList.cbo_towarehouse.Column(1) & "'"
report.FormulaFields(7).Text = "'" & frmRptSupplyList.cbo_supply.Column(0) & "'"
report.FormulaFields(8).Text = "'" & frmRptSupplyList.cbo_supply.Column(1) & "'"
report.FormulaFields(9).Text = gi_decimalDigitQtyBOM
report.FormulaFields(10).Text = "'" & frmRptSupplyList.cbo_frwarehouse.Column(0) & "'"
report.FormulaFields(11).Text = "'" & frmRptSupplyList.cbo_frwarehouse.Column(1)

    
    report.ReportTitle = "Parts (Material) Supply List Report"

Dim vBinNumbers As Variant
If combo1.ListIndex >= 0 Then
    vBinNumbers = GetBinNumbers
    report.SelectPrinter Label3, cbxPrinters.Text, Label2
    report.PaperSize = crPaperA4
    report.PaperOrientation = Orient
    report.PaperSource = vBinNumbers(combo1.ListIndex)
Else
    MsgBox "No paper tray has been selected."
End If

Dim nCopy As Integer
nCopy = Me.txtCopies
Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
Set report = Nothing
rsRpt.Close
Set rsRpt = Nothing
 
Me.MousePointer = vbDefault
End Sub

Private Sub SupplyListValue()
Dim application As New CRAXDDRT.application
Dim report As New CRAXDDRT.report
Dim rsRpt As New ADODB.Recordset

Me.MousePointer = vbHourglass

If rsRpt.State <> adStateClosed Then rsRpt.Close
rsRpt.Open sqlprint, Db, adOpenDynamic, adLockOptimistic
         
Set report = application.OpenReport(App.path & "\Reports\rptSupplyListValue.rpt")

report.Database.Tables(1).SetDataSource rsRpt
report.FormulaFields(2).Text = "'" & Format(FrmValuationPriceReportCls.dtAwal.Value, "dd MMM yyyy") & "'"
report.FormulaFields(3).Text = "'" & Format(FrmValuationPriceReportCls.dtAkhir.Value, "dd MMM yyyy") & "'"
report.FormulaFields(5).Text = "'" & FrmValuationPriceReportCls.cbo_towarehouse.Column(0) & "'"
report.FormulaFields(6).Text = "'" & FrmValuationPriceReportCls.cbo_towarehouse.Column(1) & "'"
report.FormulaFields(7).Text = "'" & FrmValuationPriceReportCls.cbo_supply.Column(0) & "'"
report.FormulaFields(8).Text = "'" & FrmValuationPriceReportCls.cbo_supply.Column(1) & "'"
report.FormulaFields(9).Text = gi_decimalDigitQty
report.FormulaFields(10).Text = "'" & FrmValuationPriceReportCls.cbo_frwarehouse.Column(0) & "'"
report.FormulaFields(11).Text = "'" & FrmValuationPriceReportCls.cbo_frwarehouse.Column(1) & "'"

Dim vBinNumbers As Variant
If combo1.ListIndex >= 0 Then
    vBinNumbers = GetBinNumbers
    report.SelectPrinter Label3, cbxPrinters.Text, Label2
    report.PaperSize = crPaperA4
    report.PaperOrientation = crLandscape
    report.PaperSource = vBinNumbers(combo1.ListIndex)
Else
    MsgBox "No paper tray has been selected."
End If

Dim nCopy As Integer
nCopy = Me.txtCopies
Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
Set report = Nothing
rsRpt.Close
Set rsRpt = Nothing
 
Me.MousePointer = vbDefault
End Sub

Private Sub pmsBOM()
Dim application As New CRAXDDRT.application
Dim report As New CRAXDDRT.report
Dim rsRpt As New ADODB.Recordset

If rsRpt.State <> adStateClosed Then rsRpt.Close
rsRpt.Open sqlprint, Db, adOpenDynamic, adLockOptimistic

Set report = application.OpenReport(App.path & "\Reports\rptPartMaterialSupplyBOM.rpt")
report.Database.Tables(1).SetDataSource rsRpt
report.FormulaFields(1).Text = gi_decimalDigitQtyBOM
report.FormulaFields(4).Text = gi_decimalDigitQty

Dim vBinNumbers As Variant
If combo1.ListIndex >= 0 Then
    vBinNumbers = GetBinNumbers
    report.SelectPrinter Label3, cbxPrinters.Text, Label2
    report.PaperSize = crPaperA4
    report.PaperOrientation = Orient
    report.PaperSource = vBinNumbers(combo1.ListIndex)
Else
    MsgBox "No paper tray has been selected."
End If

Dim nCopy As Integer
nCopy = Me.txtCopies
Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
Set report = Nothing
rsRpt.Close
Set rsRpt = Nothing
 
Me.MousePointer = vbDefault
End Sub
Private Sub salesreport()
Dim application As New CRAXDDRT.application
Dim report As New CRAXDDRT.report
Dim rsRpt As New ADODB.Recordset

If rsRpt.State <> adStateClosed Then rsRpt.Close
rsRpt.Open sqlprint, Db, adOpenDynamic, adLockOptimistic

Set report = application.OpenReport(App.path & "\Reports\SalesReport.rpt")
report.Database.Tables(1).SetDataSource rsRpt
report.FormulaFields(1).Text = "'" & Format(FrmSalesReport.dtAwal.Value, "dd mmmm yyyy") & "'"
report.FormulaFields(2).Text = "'" & Format(FrmSalesReport.dtAkhir.Value, "dd mmmm yyyy") & "'"
report.FormulaFields(3).Text = "'" & FrmSalesReport.cbo_trade.Column(0) & "'"
report.FormulaFields(4).Text = "'" & FrmSalesReport.cbo_trade.Column(1) & "'"
report.FormulaFields(5).Text = "'" & FrmSalesReport.cbo_group.Column(0) & "'"
report.FormulaFields(8).Text = "'" & FrmSalesReport.cbo_group.Column(1) & "'"
report.FormulaFields(6).Text = "'" & FrmSalesReport.cbo_item.Column(0) & "'"
report.FormulaFields(7).Text = "'" & FrmSalesReport.cbo_item.Column(1) & "'"
report.FormulaFields(14).Text = gi_decimalDigitQty

Dim vBinNumbers As Variant
If combo1.ListIndex >= 0 Then
    vBinNumbers = GetBinNumbers
    report.SelectPrinter Label3, cbxPrinters.Text, Label2
    report.PaperSize = crPaperA4
    report.PaperOrientation = Orient
    report.PaperSource = vBinNumbers(combo1.ListIndex)
Else
    MsgBox "No paper tray has been selected."
End If

Dim nCopy As Integer
nCopy = Me.txtCopies
Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
Set report = Nothing
rsRpt.Close
Set rsRpt = Nothing
 
Me.MousePointer = vbDefault
End Sub

Private Sub PrintBC40()
    
  Dim application As New CRAXDDRT.application
  Dim report As New CRAXDDRT.report
  Dim rsRpt As New ADODB.Recordset
  Dim rsrptsub1 As New ADODB.Recordset
  Dim rsrptsub2 As New ADODB.Recordset
  
  
  rsRpt.CursorLocation = adUseClient
  If rsRpt.State <> adStateClosed Then rsRpt.Close
  rsRpt.Open sqlprint, Db, adOpenDynamic, adLockOptimistic
  
  If rsRpt.EOF Then Exit Sub

    Set report = application.OpenReport(App.path & "\Reports\bc_40.rpt")
    report.Database.Tables(1).SetDataSource rsRpt
    
    report.FormulaFields(1).Text = "" & gi_decimalDigitQty & ""
    report.FormulaFields(2).Text = "" & gi_decimalDigitAmount & ""
    report.FormulaFields(3).Text = "" & gi_decimalDigitAmountIDR & ""
    report.FormulaFields(6).Text = "" & gi_decimalDigitBox & ""
    report.FormulaFields(7).Text = "" & rsRpt.RecordCount & ""
    
    Set rsrptsub1 = New Recordset
    rsrptsub1.CursorLocation = adUseClient
    rsrptsub1.Open sqlprint2, Db, adOpenKeyset, adLockOptimistic
    report.OpenSubreport("po").Database.Tables(1).SetDataSource rsrptsub1
        
    Set rsrptsub2 = New Recordset
    rsrptsub2.CursorLocation = adUseClient
    rsrptsub2.Open sqlprint3, Db, adOpenKeyset, adLockOptimistic
    report.OpenSubreport("amount").Database.Tables(1).SetDataSource rsrptsub2
    report.OpenSubreport("amount").FormulaFields(3).Text = "" & gi_decimalDigitAmount & ""
    report.OpenSubreport("amount").FormulaFields(4).Text = "" & gi_decimalDigitAmountIDR & ""
    
    Dim vBinNumbers As Variant
    If combo1.ListIndex >= 0 Then
        vBinNumbers = GetBinNumbers
        report.SelectPrinter Label3, cbxPrinters.Text, Label2
        report.PaperSize = crPaperLegal
        report.PaperOrientation = Orient
        report.PaperSource = vBinNumbers(combo1.ListIndex)
    Else
        MsgBox "No paper tray has been selected."
    End If

    Dim nCopy As Integer
    nCopy = Me.txtCopies
    Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
    Set report = Nothing
    rsRpt.Close
    Set rsRpt = Nothing
        
End Sub

Sub InvoiceDetailList()

    Dim application As New CRAXDDRT.application
    Dim report As New CRAXDDRT.report
    Dim rs1 As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim rs3 As New ADODB.Recordset

    rs1.Open sqlprint, Db, adOpenKeyset
    If Not rs1.EOF Then
    
        Set report = application.OpenReport(App.path & "\Reports\DetailOfSalesReport.rpt")
        report.Database.Tables(1).SetDataSource rs1
        report.FormulaFields(1).Text = "'" & Format(F_DetailSalesReport.MDate, "dd MMMM  yyyy") & " to " & Format(F_DetailSalesReport.MDate1, "dd MMMM  yyyy") & "'"
        
        If rs2.State <> adStateClosed Then rs2.Close
        rs2.Open sqlprint2, Db, adOpenDynamic
        report.OpenSubreport("summary_cust").Database.Tables(1).SetDataSource rs2
        
        If rs3.State <> adStateClosed Then rs3.Close
        rs3.Open sqlprint3, Db, adOpenDynamic
        report.OpenSubreport("summary_country").Database.Tables(1).SetDataSource rs3
        
        Dim vBinNumbers As Variant
        If combo1.ListIndex >= 0 Then
            vBinNumbers = GetBinNumbers
            report.SelectPrinter Label3, cbxPrinters.Text, Label2
            report.PaperOrientation = Orient
            report.PaperSource = vBinNumbers(combo1.ListIndex)
        Else
            MsgBox "No paper tray has been selected."
        End If
        
        Dim nCopy As Integer
        nCopy = Me.txtCopies
        Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
        Set report = Nothing
        
        If rs1.State <> adStateClosed Then rs1.Close
        Set rs1 = Nothing
        
        If rs2.State <> adStateClosed Then rs2.Close
        Set rs2 = Nothing
        
        If rs3.State <> adStateClosed Then rs3.Close
        Set rs3 = Nothing

    End If

End Sub

Sub InvoiceSummaryList()

    Dim application As New CRAXDDRT.application
    Dim report As New CRAXDDRT.report
    Dim rs1 As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset

    rs1.Open sqlprint, Db, adOpenKeyset
    If Not rs1.EOF Then
    
        Set report = application.OpenReport(App.path & "\Reports\SummaryOfSalesReport.rpt")
        report.Database.Tables(1).SetDataSource rs1
        report.FormulaFields(1).Text = "'" & Format(F_SummarySalesReport.MDate, "dd MMMM  yyyy") & " to " & Format(F_SummarySalesReport.MDate1, "dd MMMM  yyyy") & "'"
                
        If rs2.State <> adStateClosed Then rs2.Close
        rs2.Open sqlprint2, Db, adOpenDynamic
        report.OpenSubreport("summary_country").Database.Tables(1).SetDataSource rs2
        
        Dim vBinNumbers As Variant
        If combo1.ListIndex >= 0 Then
            vBinNumbers = GetBinNumbers
            report.SelectPrinter Label3, cbxPrinters.Text, Label2
            report.PaperOrientation = Orient
            report.PaperSource = vBinNumbers(combo1.ListIndex)
        Else
            MsgBox "No paper tray has been selected."
        End If
        
        Dim nCopy As Integer
        nCopy = Me.txtCopies
        Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
        Set report = Nothing
        
        If rs1.State <> adStateClosed Then rs1.Close
        Set rs1 = Nothing
        
        If rs2.State <> adStateClosed Then rs2.Close
        Set rs2 = Nothing
        
    End If

End Sub

Sub RptPORequestPrint()
Dim application As New CRAXDDRT.application
Dim report As New CRAXDDRT.report
Dim rs1 As Recordset

Set rs1 = New Recordset
rs1.Open sqlprint, Db, adOpenKeyset, adLockOptimistic
If Not rs1.EOF Then

    Set report = application.OpenReport(App.path & "\REPORTs\RptPOReqReport.rpt")
    report.Database.Tables(1).SetDataSource rs1
    report.FormulaFields(1).Text = tglAwalRptPrint
    report.FormulaFields(2).Text = Kt1
    report.FormulaFields(3).Text = Kt2
    report.FormulaFields(4).Text = Kt3

    Dim vBinNumbers As Variant
    If combo1.ListIndex >= 0 Then
    vBinNumbers = GetBinNumbers

    report.SelectPrinter Label3, cbxPrinters.Text, Label2
    report.PaperOrientation = crLandscape
    report.PaperSource = vBinNumbers(combo1.ListIndex)
    'report.PaperSize = crPaperA4
    Else
    MsgBox "No paper tray has been selected."
    End If

    Dim nCopy As Integer
    nCopy = Me.txtCopies
    Call Printout(report, nCopy, True, Val(txtRange(0)), Val(txtRange(1)))
    Set report = Nothing
    rs1.Close
    Set rs1 = Nothing
End If
End Sub
