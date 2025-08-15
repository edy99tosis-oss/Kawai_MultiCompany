VERSION 5.00
Begin VB.Form FrmEmailConfig 
   BackColor       =   &H00FDDFE3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Email Configuration"
   ClientHeight    =   8880
   ClientLeft      =   6255
   ClientTop       =   4065
   ClientWidth     =   11160
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmEmailConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   11160
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "send"
      Height          =   375
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   38
      Tag             =   "FFTT*/"
      Top             =   8280
      Width           =   1140
   End
   Begin VB.CommandButton CmdClear 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Clear"
      Height          =   375
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   15
      Tag             =   "FFTT*/"
      Top             =   8280
      Width           =   1140
   End
   Begin VB.CommandButton CmdSubmit 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Submit"
      Height          =   375
      Left            =   9795
      Style           =   1  'Graphical
      TabIndex        =   14
      Tag             =   "FFTT*/"
      Top             =   8280
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Caption         =   " Login Information "
      Height          =   3915
      Left            =   240
      TabIndex        =   21
      Tag             =   "TTTT*/"
      Top             =   3120
      Width           =   10695
      Begin VB.TextBox TxtCC 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   12
         Tag             =   "TTFF*/"
         Top             =   3075
         Width           =   3660
      End
      Begin VB.TextBox TxtBCC 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2160
         MaxLength       =   500
         TabIndex        =   13
         Tag             =   "TTFF*/"
         Top             =   3435
         Width           =   8460
      End
      Begin VB.TextBox TxtMailSign 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   3120
         MaxLength       =   20
         TabIndex        =   11
         Tag             =   "TTFF*/"
         Top             =   4140
         Width           =   3660
      End
      Begin VB.TextBox TxtMailFooter 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2160
         MaxLength       =   150
         TabIndex        =   10
         Tag             =   "TTFF*/"
         Top             =   2730
         Width           =   3660
      End
      Begin VB.TextBox TxtMailContent 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2160
         MaxLength       =   150
         TabIndex        =   9
         Tag             =   "TTFF*/"
         Top             =   2400
         Width           =   3660
      End
      Begin VB.TextBox TxtMailHeader 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   8
         Tag             =   "TTFF*/"
         Top             =   2070
         Width           =   3660
      End
      Begin VB.TextBox TxtSubject 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2160
         MaxLength       =   75
         TabIndex        =   7
         Tag             =   "TTFF*/"
         Top             =   1740
         Width           =   3660
      End
      Begin VB.TextBox TxtTimer 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2160
         MaxLength       =   9
         TabIndex        =   6
         Tag             =   "TTFF*/"
         Top             =   1410
         Width           =   3660
      End
      Begin VB.TextBox TxtEmail 
         Height          =   285
         Left            =   2160
         MaxLength       =   30
         TabIndex        =   3
         Tag             =   "TTFF*/"
         Top             =   360
         Width           =   3660
      End
      Begin VB.TextBox TxtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2160
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   5
         Tag             =   "TTFF*/"
         Top             =   1080
         Width           =   3660
      End
      Begin VB.TextBox TxtUserName 
         Height          =   285
         Left            =   2160
         MaxLength       =   30
         TabIndex        =   4
         Tag             =   "TTFF*/"
         Top             =   720
         Width           =   3660
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "To Email Address"
         Height          =   225
         Left            =   240
         TabIndex        =   35
         Tag             =   "TTFF*/"
         Top             =   3105
         Width           =   1785
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "CC Email Address"
         Height          =   225
         Left            =   240
         TabIndex        =   34
         Tag             =   "TTFF*/"
         Top             =   3435
         Width           =   1785
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Mail Sign"
         Height          =   225
         Left            =   1200
         TabIndex        =   33
         Tag             =   "TTFF*/"
         Top             =   4140
         Width           =   1785
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Mail Footer"
         Height          =   225
         Left            =   240
         TabIndex        =   32
         Tag             =   "TTFF*/"
         Top             =   2730
         Width           =   1785
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Mail Content"
         Height          =   225
         Left            =   240
         TabIndex        =   31
         Tag             =   "TTFF*/"
         Top             =   2400
         Width           =   1785
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Mail Header"
         Height          =   225
         Left            =   240
         TabIndex        =   30
         Tag             =   "TTFF*/"
         Top             =   2070
         Width           =   1785
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Subject"
         Height          =   225
         Left            =   240
         TabIndex        =   29
         Tag             =   "TTFF*/"
         Top             =   1740
         Width           =   1785
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Timer"
         Height          =   225
         Left            =   240
         TabIndex        =   28
         Tag             =   "TTFF*/"
         Top             =   1410
         Width           =   1785
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Email Address"
         Height          =   225
         Left            =   240
         TabIndex        =   27
         Tag             =   "TTFF*/"
         Top             =   360
         Width           =   1785
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         Height          =   225
         Left            =   240
         TabIndex        =   26
         Tag             =   "TTFF*/"
         Top             =   1080
         Width           =   1785
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
         Height          =   225
         Left            =   240
         TabIndex        =   25
         Tag             =   "TTFF*/"
         Top             =   720
         Width           =   1785
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Caption         =   " SMTP Server "
      Height          =   1875
      Left            =   240
      TabIndex        =   20
      Tag             =   "TTTF*/"
      Top             =   1080
      Width           =   10695
      Begin VB.TextBox TxtSmtpTimeout 
         Height          =   285
         Left            =   2160
         MaxLength       =   9
         TabIndex        =   36
         Tag             =   "TTFF*/"
         Top             =   1380
         Width           =   3660
      End
      Begin VB.TextBox TxtPortName 
         Height          =   285
         Left            =   2160
         MaxLength       =   9
         TabIndex        =   2
         Tag             =   "TTFF*/"
         Top             =   1020
         Width           =   3660
      End
      Begin VB.TextBox TxtDesc 
         Height          =   285
         Left            =   2160
         MaxLength       =   30
         TabIndex        =   1
         Tag             =   "TTFF*/"
         Top             =   660
         Width           =   3660
      End
      Begin VB.TextBox TxtServerName 
         Height          =   285
         Left            =   2160
         MaxLength       =   30
         TabIndex        =   0
         Tag             =   "TTFF*/"
         Top             =   300
         Width           =   3660
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Smtp Timeout"
         Height          =   225
         Left            =   240
         TabIndex        =   37
         Tag             =   "TTFF*/"
         Top             =   1380
         Width           =   1785
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Port Number"
         Height          =   225
         Left            =   240
         TabIndex        =   24
         Tag             =   "TTFF*/"
         Top             =   1020
         Width           =   1785
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   225
         Left            =   240
         TabIndex        =   23
         Tag             =   "TTFF*/"
         Top             =   660
         Width           =   1785
      End
      Begin VB.Label label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Server Name"
         Height          =   225
         Left            =   240
         TabIndex        =   22
         Tag             =   "TTFF*/"
         Top             =   300
         Width           =   1785
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FDDFE3&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   240
      TabIndex        =   18
      Tag             =   "TFTT*/"
      Top             =   7560
      Width           =   10695
      Begin VB.Label LblErrMsg 
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
         Left            =   105
         TabIndex        =   19
         Tag             =   "TFTF*/"
         Top             =   195
         Width           =   10410
      End
   End
   Begin VB.CommandButton cmdsubmenu 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Sub &Menu"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   16
      Tag             =   "TFFT*/"
      Top             =   8280
      Width           =   1140
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Email Configuration"
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
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Tag             =   "TTTF*/"
      Top             =   300
      Width           =   10725
   End
End
Attribute VB_Name = "FrmEmailConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'##################################
'Created By     : Anak Toss
'Created Date   : 19 Agustus 2010
'Last Update By :
'Last Update    :
'##################################

Option Explicit
Option Compare Text

Dim li_Special As Byte
Dim ls_ServerName As String
Private gvFSO As Scripting.FileSystemObject

Private Sub cmdClear_Click()
up_ClearScreen
End Sub

Private Sub CmdSubMenu_Click()
   Dim ls_Answer As String
    
    ls_Answer = MsgBox("Are you sure want to close this menu ?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation")
    If ls_Answer = vbYes Then
        If cmdsubmenu.Caption = "Sub &Menu" Then
            DoEvents
            frmMainMenu.Show
            DoEvents
            Unload Me
        End If
    End If

End Sub

Private Sub CmdSubmit_Click()
If uf_Validate = False Then Exit Sub
    up_SaveData
End Sub

Private Sub Command1_Click()
Me.MousePointer = vbHourglass
up_SendEmail
Me.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
  

  Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"
'  kosong
  up_LoadData
    
  
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then Cancel = True
End Sub

Private Sub up_ClearScreen()
    txtServerName = ""
    txtDesc = ""
    TxtPortName = ""
    TxtSmtpTimeout = ""
    TxtEmail = ""
    TxtUserName = ""
    txtPassword = ""
    TxtTimer = ""
    TxtSubject = ""
    TxtMailHeader = ""
    TxtMailContent = ""
    TxtMailFooter = ""
    TxtMailSign = ""
    TxtCC = ""
    TxtBCC = ""
End Sub

Private Sub up_LoadData()
Dim ls_sql As String
Dim RS As New Recordset


ls_sql = "Select * From Email_Config"
RS.Open ls_sql, Db, adOpenDynamic, adLockOptimistic

If Not RS.EOF Then
    txtServerName = Trim(RS!smtp_server & "")
    txtDesc = Trim(RS!smtp_desc & "")
    TxtPortName = Val(RS!port_number)
    TxtSmtpTimeout = Val(RS!smtp_timeout)
    TxtEmail = Trim(RS!Email_Address & "")
    TxtUserName = Trim(RS!user_email & "")
    txtPassword = Trim(RS!pass_email & "")
    TxtTimer = Val(RS!Timer)
    TxtSubject = Trim(RS!Subject & "")
    TxtMailHeader = Trim(RS!mail_header & "")
    TxtMailContent = Trim(RS!Mail_Content & "")
    TxtMailFooter = Trim(RS!Mail_Footer & "")
    TxtMailSign = Trim(RS!mail_sign & "")
    TxtCC = Trim(RS!ccemail_address & "")
    TxtBCC = Trim(RS!bccemail_address & "")
End If

RS.Close

End Sub

Private Sub up_SaveData()
Dim ls_sql As String
Dim li_RowAff As Integer

ls_sql = " update email_config set smtp_server = '" & Trim(txtServerName) & "', smtp_desc = '" & Trim(txtDesc) & "',  " & vbCrLf & _
        " port_number = " & Val(Trim(TxtPortName)) & ", smtp_timeout = " & Val(Trim(TxtSmtpTimeout)) & ", email_address = '" & Trim(TxtEmail) & "', user_email = '" & Trim(TxtUserName) & "', " & vbCrLf & _
        " pass_email = '" & Trim(txtPassword) & "', timer = " & Val(Trim(TxtTimer)) & ", subject = '" & Trim(TxtSubject) & "',  " & vbCrLf & _
        " mail_header = '" & Trim(TxtMailHeader) & "', mail_content = '" & Trim(TxtMailContent) & "', mail_footer = '" & Trim(TxtMailFooter) & "', mail_sign = '" & Trim(TxtMailSign) & "', " & vbCrLf & _
        " ccemail_address = '" & Trim(TxtCC.Text) & "', bccemail_address = '" & Trim(TxtBCC.Text) & "' "
Db.Execute ls_sql, li_RowAff

If li_RowAff = 0 Then
    ls_sql = "Insert Into Email_Config (smtp_server, smtp_desc, port_number, smtp_timeout, email_address, user_email, pass_email, timer, subject, mail_header, mail_content, mail_footer, mail_sign, ccemail_address, bccemail_address)" & vbCrLf & _
             "Values ('" & Trim(txtServerName) & "', '" & Trim(txtDesc) & "', " & Val(Trim(TxtPortName)) & ", " & Val(Trim(TxtSmtpTimeout)) & ", '" & Trim(TxtEmail) & "', '" & Trim(TxtUserName) & "', '" & Trim(txtPassword) & "', " & vbCrLf & _
             " " & Val(Trim(TxtTimer)) & ", '" & Trim(TxtSubject) & "', '" & Trim(TxtMailHeader) & "', '" & Trim(TxtMailContent) & "', '" & Trim(TxtMailFooter) & "', '" & Trim(TxtMailSign) & "', " & vbCrLf & _
             " '" & Trim(TxtCC) & "', '" & Trim(TxtBCC) & "' ) "
             
    Db.Execute ls_sql
End If

LblErrMsg = DisplayMsg(1000)

End Sub

Private Function uf_Validate() As Boolean
    
    uf_Validate = True
    If txtServerName = "" Then
        uf_Validate = False
        txtServerName.SetFocus
        LblErrMsg.Caption = "Please input Server Name!"
    ElseIf txtDesc = "" Then
        uf_Validate = False
        txtDesc.SetFocus
        LblErrMsg.Caption = "Please input Description!"
    ElseIf TxtPortName = "" Then
        uf_Validate = False
        TxtPortName.SetFocus
        LblErrMsg.Caption = "Please input Port Name!"
    ElseIf IsNumeric(TxtPortName.Text) = False Then
        uf_Validate = False
        TxtPortName.SetFocus
        LblErrMsg.Caption = "Please input valid Port Name!"
    ElseIf TxtSmtpTimeout = "" Then
        uf_Validate = False
        TxtSmtpTimeout.SetFocus
        LblErrMsg.Caption = "Please input Smtp Timeout!"
    ElseIf IsNumeric(TxtSmtpTimeout.Text) = False Then
        uf_Validate = False
        TxtSmtpTimeout.SetFocus
        LblErrMsg.Caption = "Please input valid Smtp Timeout!"
    ElseIf TxtEmail = "" Then
        uf_Validate = False
        TxtEmail.SetFocus
        LblErrMsg.Caption = "Please input Email Address!"
    ElseIf TxtUserName = "" Then
        uf_Validate = False
        TxtUserName.SetFocus
        LblErrMsg.Caption = "Please input User Name!"
    ElseIf txtPassword = "" Then
        uf_Validate = False
        txtPassword.SetFocus
        LblErrMsg.Caption = "Please input Password!"
    ElseIf TxtTimer = "" Then
        uf_Validate = False
        TxtTimer.SetFocus
        LblErrMsg.Caption = "Please input Timer!"
    ElseIf IsNumeric(TxtTimer.Text) = False Then
        uf_Validate = False
        TxtTimer.SetFocus
        LblErrMsg.Caption = "Please input valid Timer!"
    ElseIf TxtSubject = "" Then
        uf_Validate = False
        TxtSubject.SetFocus
        LblErrMsg.Caption = "Please input Subject!"
    ElseIf TxtMailHeader = "" Then
        uf_Validate = False
        TxtMailHeader.SetFocus
        LblErrMsg.Caption = "Please input Mail Header!"
    ElseIf TxtMailContent = "" Then
        uf_Validate = False
        TxtMailContent.SetFocus
        LblErrMsg.Caption = "Please input Mail Content!"
    ElseIf TxtMailFooter = "" Then
        uf_Validate = False
        TxtMailFooter.SetFocus
        LblErrMsg.Caption = "Please input Mail Footer!"
'    ElseIf TxtMailSign = "" Then
'        uf_Validate = False
'        TxtMailSign.SetFocus
'        LblErrMsg.Caption = "Please input Mail Sign!"
    End If
    
End Function

Private Sub up_SendEmail()
    'Recordset
    Dim rsmailconfig    As New ADODB.Recordset
    Dim rsHeader        As New ADODB.Recordset
    Dim rsdetail        As New ADODB.Recordset
    
    'Excel Application
    Dim xlapp As New Excel.application
    
    Dim ls_SupplierCodeGrid As String
    Dim ls_TicketDateGrid   As String
    
    'Variable
    Dim ls_sql As String
    Dim li_Row As Integer
    Dim ls_smtpserver As String
    Dim li_smtpport     As Double
    Dim li_smtptimeout  As Double
    Dim ls_sender       As String
    Dim ls_username     As String
    Dim ls_password     As String
    Dim li_timer        As Double
    Dim ls_subject      As String
    Dim ls_cc           As String
    Dim ls_bcc          As String
    Dim ls_message      As String
    Dim ls_recipient    As String
    Dim ls_SupplierCode As String
    Dim ls_TicketNo     As String
    Dim ls_SampleNo     As String
    Dim ls_ticketdate   As String
    Dim ls_ItemCode     As String
    Dim li_wet          As Double
    Dim li_drc          As Double
    Dim li_dry          As Double
    Dim li_price        As Double
    Dim li_amount       As Double
    Dim ls_log          As String
    
    Dim Idx                 As Integer
    Dim li_totalwet         As Double, li_totaldry          As Double
    Dim li_totalamout       As Double, li_totaladvanced     As Double
    Dim li_totaltaxamount   As Double, li_totalfinalpaid    As Double
    
    Dim strFile         As String
    
    
    Dim strAttachFile   As String
    Const clrGrey = 15
    
    Dim ls_desc As String
    
    On Local Error GoTo errHandler
    
        'Set Timer
    ls_sql = " select smtp_server, port_number, smtp_timeout, email_address, user_email, pass_email, timer, subject, " & vbCrLf & _
            "       ccemail_address, bccemail_address,mail_header,Mail_Content,Mail_Footer " & vbCrLf & _
            " from email_config "
            
    If rsmailconfig.State <> adStateClosed Then rsmailconfig.Close
    Set rsmailconfig = Db.Execute(ls_sql)
    
    If Not rsmailconfig.EOF Then
        ls_smtpserver = Trim(rsmailconfig!smtp_server & "")
        li_smtpport = IIf(IsNull(rsmailconfig!port_number), 0, rsmailconfig!port_number)
        li_smtptimeout = IIf(IsNull(rsmailconfig!smtp_timeout), 0, rsmailconfig!smtp_timeout)
        ls_sender = Trim(rsmailconfig!Email_Address & "")
        ls_username = Trim(rsmailconfig!user_email & "")
        ls_password = Trim(rsmailconfig!pass_email & "")
        li_timer = IIf(IsNull(rsmailconfig!Timer), 0, rsmailconfig!Timer)
        ls_subject = Trim(rsmailconfig!Subject & "")
        ls_cc = Trim(rsmailconfig!ccemail_address & "")
        ls_bcc = Trim(rsmailconfig!bccemail_address & "")
        ls_message = Trim(rsmailconfig!mail_header & "") & vbCrLf & vbCrLf
        ls_message = ls_message & Trim(rsmailconfig!Mail_Content & "") & vbCrLf & vbCrLf
        ls_message = ls_message & "Part No          :" & vbCrLf
        ls_message = ls_message & "Part Name        :" & vbCrLf
        ls_message = ls_message & "Qty              :" & vbCrLf
        ls_message = ls_message & "Vendor           :" & vbCrLf
        ls_message = ls_message & "Surat Jalan      :" & vbCrLf & vbCrLf
        ls_message = ls_message & "Kepada yang berkepentingan Mohon Segera Pastikan Barang Tersebut." & vbCrLf
        ls_message = ls_message & "Terima Kasih" & vbCrLf & vbCrLf & vbCrLf
        ls_message = ls_message & "Note :" & vbCrLf
        ls_message = ls_message & Trim(rsmailconfig!Mail_Footer & "")
        
        
        
    End If
    
    rsmailconfig.Close
    Set rsmailconfig = Nothing
    
    
      '========================================================================
      'start send email function
      '========================================================================

      Call up_cdoSendEmail(ls_smtpserver, li_smtpport, li_smtptimeout, ls_username, ls_password, ls_sender, ls_recipient, ls_subject & ls_TicketNo, ls_message, ls_cc, ls_bcc, strAttachFile)

      '========================================================================
      'end send email function
      '========================================================================
      
    
ErrExit:
    Exit Sub
errHandler:
    'Call SaveToTxt("SendEmail_ErrorLog", " Ticket No : " & ls_TicketNo & " ( " & ls_ticketdate & " ) " & err.Description, CurrDate)
    err.clear
    Resume ErrExit
    
End Sub



Public Sub up_cdoSendEmail(ByVal ls_smtpserver As String, ByVal li_smtpport As Double, ByVal li_smtptimeout As Double, _
                    ByVal ls_username As String, ByVal ls_password As String, ByVal ls_sender As String, _
                    ByVal ls_recipient As String, ByVal ls_subject As String, ByVal ls_body As String, _
                    Optional ByVal ls_cc As String, Optional ByVal ls_bcc As String, Optional ByVal ls_attachment As String)
    
    Dim cdoMsg As New CDO.message
    Dim cdoConf As New CDO.Configuration
    Dim Flds
    Dim attachment
   
    On Error GoTo errHandler
    
    Set Flds = cdoConf.Fields
       
    With Flds
        .Item(cdoSendUsingMethod) = cdoSendUsingPort
        .Item(cdoSMTPServer) = ls_smtpserver
        '.Item(cdoSMTPServerPort) = IIf(li_smtpport = 0, 25, li_smtpport)
        .Item(cdoSMTPServerPort) = li_smtpport
        .Item(cdoSMTPConnectionTimeout) = IIf(li_smtptimeout = 0, 10, li_smtptimeout)
        .Item(cdoSMTPAuthenticate) = 1 'default
        .Item(cdoSendUserName) = ls_username
        .Item(cdoSendPassword) = ls_password
        .Item(cdoSMTPUseSSL) = 1
        
        
        
        
        
        
    .update
    End With
   
   
    ' Apply the settings to the message.
    With cdoMsg
        Set .Configuration = cdoConf
        .To = ls_cc
        .From = ls_sender
        .Subject = ls_subject
        .TextBody = ls_body
        
'        If Not colAttachments Is Nothing Then
'            For Each attachment In colAttachments
'                .AddAttachment attachment
'            Next
'        End If
        'If ls_cc <> "" Then .CC = ls_cc
        If ls_bcc <> "" Then .BCC = ls_bcc
        .Send
    End With
    LblErrMsg.Caption = "Send Successful"
    
   
ErrExit:
    Set cdoMsg = Nothing
    Set cdoConf = Nothing
    Set Flds = Nothing
    
    Exit Sub

errHandler:
LblErrMsg.Caption = err.Description
'Call SaveToTxt("SendEmail_ErrorLog", "SEND MAIL - " & err.number & " : " & err.Description, CurrDate)
GoTo ErrExit
    
End Sub
