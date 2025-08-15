VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EZ Runner ver.3"
   ClientHeight    =   7350
   ClientLeft      =   2580
   ClientTop       =   2670
   ClientWidth     =   9270
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmLogin.frx":0E42
   MousePointer    =   99  'Custom
   ScaleHeight     =   7350
   ScaleWidth      =   9270
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMenu 
      BackColor       =   &H00FFFFFF&
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
      IMEMode         =   3  'DISABLE
      Left            =   7065
      MaxLength       =   6
      TabIndex        =   2
      Top             =   5115
      Width           =   840
   End
   Begin VB.TextBox txtUser 
      BackColor       =   &H00FFFFFF&
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
      Left            =   7065
      MaxLength       =   15
      TabIndex        =   0
      Text            =   "Admin"
      Top             =   4275
      Width           =   1500
   End
   Begin VB.TextBox txtPass 
      BackColor       =   &H00FFFFFF&
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
      IMEMode         =   3  'DISABLE
      Left            =   7065
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "adm"
      Top             =   4695
      Width           =   1500
   End
   Begin VB.CommandButton cmdLogin 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Login"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6450
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5655
      Width           =   1005
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0E0FF&
      Cancel          =   -1  'True
      Caption         =   "&Exit"
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
      Left            =   7565
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5670
      Width           =   1000
   End
   Begin VB.Label Label3 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   4
      Left            =   885
      TabIndex        =   10
      Top             =   6285
      Width           =   435
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   3
      Left            =   8115
      TabIndex        =   9
      Top             =   6285
      Width           =   435
   End
   Begin VB.Image Image2 
      Height          =   1305
      Left            =   705
      Picture         =   "frmLogin.frx":114C
      Top             =   4995
      Width           =   4920
   End
   Begin VB.Image Image4 
      Height          =   270
      Left            =   7575
      Picture         =   "frmLogin.frx":15FF6
      Top             =   195
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Image Image1 
      Height          =   3630
      Left            =   735
      Picture         =   "frmLogin.frx":1650C
      Top             =   405
      Width           =   7335
   End
   Begin VB.Label LblErrMsg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      Height          =   450
      Left            =   825
      TabIndex        =   8
      Top             =   6555
      Width           =   7740
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Menu ID   :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   5835
      TabIndex        =   7
      Top             =   5115
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Id     :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   5835
      TabIndex        =   6
      Top             =   4275
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   5835
      TabIndex        =   5
      Top             =   4695
      Width           =   1095
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sql As String
Dim rsUser As New ADODB.Recordset

Public strServer     As String
Public strDatabase   As String
Public strUserID     As String
Public strPassword   As String
Dim DbTimeout        As String
Dim CommandTimeout   As String

Sub Kosong()
    txtUser.Text = ""
    txtPass.Text = ""
    txtmenu.Text = ""
    LblErrMsg = ""
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim Ret As String, NC As Long, TempPWD As String, dbmaster As String

    err.clear
    IniFile = App.path & "\config.ini"
    Label3(4) = "Build : " & App.Major & "." & App.Minor & "." & App.Revision

    'CONNSTR
    'Ret = String(255, 0)
    'NC = GetPrivateProfileString("Database", "ConnStr", "", Ret, 255, IniFile)
    'If NC <> 0 Then Ret = Left$(Ret, NC)
    'ConnStr = Ret
    
   'Connection String
   Ret = String(1500, 0)
   NC = GetPrivateProfileString("Database", "Provider", "", Ret, 1500, IniFile)
   If NC <> 0 Then Ret = Left$(Ret, NC)
   ConnStr = fc_Decrypt(Ret)
   
   Ret = String(1500, 0)
   NC = GetPrivateProfileString("Database", "Server", "", Ret, 1500, IniFile)
   If NC <> 0 Then Ret = Left$(Ret, NC)
   strServer = fc_Decrypt(Ret)
    
   Ret = String(1500, 0)
   NC = GetPrivateProfileString("Database", "Database", "", Ret, 1500, IniFile)
   If NC <> 0 Then Ret = Left$(Ret, NC)
   strDatabase = fc_Decrypt(Ret)
    
   Ret = String(1500, 0)
   NC = GetPrivateProfileString("Database", "UserID", "", Ret, 1500, IniFile)
   If NC <> 0 Then Ret = Left$(Ret, NC)
   strUserID = fc_Decrypt(Ret)
    
   Ret = String(1500, 0)
   NC = GetPrivateProfileString("Database", "Password", "", Ret, 1500, IniFile)
   If NC <> 0 Then Ret = Left$(Ret, NC)
   strPassword = fc_Decrypt(Ret)
       
   ConnStr = ConnStr & ";User ID=" & strUserID & ";" & "Initial Catalog=" & strDatabase & ";" & "Data Source=" & strServer & ";" & "pwd=" & strPassword & ";"
    
    'Db.ConnectionTimeout = 120
    'Db.CommandTimeout = 120
    
   'DbTimeout
   Ret = String(255, 0)
   NC = GetPrivateProfileString("Database", "DbTimeout", "PT. YAMAHA INDONESIA MOTOR MFG.", Ret, 255, IniFile)
   If NC <> 0 Then Ret = Left$(Ret, NC)
   DbTimeout = fc_Decrypt(Ret)
   
   'CommandTimeout
   Ret = String(255, 0)
   NC = GetPrivateProfileString("Database", "CommandTimeout", "PT. YAMAHA INDONESIA MOTOR MFG.", Ret, 255, IniFile)
   If NC <> 0 Then Ret = Left$(Ret, NC)
   CommandTimeout = fc_Decrypt(Ret)
    
   ' Database backup path server
   Ret = String(255, 0)
   NC = GetPrivateProfileString("PATH", "DBBackupServer", "", Ret, 255, IniFile)
   If NC <> 0 Then Ret = Left$(Ret, NC)
   gvDBBackupServer = fc_Decrypt(Ret)
    
   ' Database backup path local
   Ret = String(255, 0)
   NC = GetPrivateProfileString("PATH", "DBBackupLocal", "", Ret, 255, IniFile)
   If NC <> 0 Then Ret = Left$(Ret, NC)
   gvDBBackupLocal = fc_Decrypt(Ret)
    
   Db.ConnectionTimeout = DbTimeout
   Db.CommandTimeout = CommandTimeout

    Db.Open ConnStr
    Db.IsolationLevel = adXactIsolated
    If err.number <> 0 Then
       MsgBox err.Description, vbCritical, "Error"
       Unload Me
       End
    End If
    
    txtUser.Text = ""
    txtPass.Text = ""
'    txtmenu.Text = "J10"
    
    txtUser.SetFocus
    'Regional Setting
    Call OpenReg
End Sub

Private Sub cmdLogin_Click()
Dim lockOut As Integer, InvalidLogin As Integer
Dim passLogin As String
    
    Me.MousePointer = vbHourglass
    If txtUser = "" Then
        txtUser.SetFocus
        LblErrMsg = DisplayMsg(1002)
    ElseIf txtPass = "" Then
        txtPass.SetFocus
        LblErrMsg = DisplayMsg(1004)
    Else
        sql = "select * from user_Setup where userName = '" & txtUser & "'"
        If rsUser.State <> adStateClosed Then rsUser.Close
        rsUser.Open sql, Db, adOpenStatic, adLockOptimistic
        
        If (rsUser.EOF And rsUser.BOF) Then 'jika salah user
            txtUser.SetFocus
            LblErrMsg = DisplayMsg(3000)
            
        Else 'jika usernya sesuai
        
            userLogin = Trim(rsUser("UserName"))
            passLogin = fc_Decrypt(Trim(rsUser("Password")))
            StatusAdmin = Trim(rsUser("status_Admin"))
            lockOut = rsUser("locked")
            InvalidLogin = rsUser("InvalidLogin")
            UserInitPO = IIf(IsNull(rsUser("InitPO")), "", rsUser("InitPO"))
            
            'DBName
            Dim Rct As String, CC As Long
            Rct = String(255, 0)
            CC = GetPrivateProfileString("Database", "Database", "", Rct, 255, IniFile)
            If CC <> 0 Then Rct = Left$(Rct, CC)
            gs_DBName = fc_Decrypt(Rct)
            
            If lockOut = 1 Then 'cek dulu udah dilockout atau belum
                LblErrMsg = DisplayMsg(3002)
            ElseIf txtPass <> passLogin Then 'jika salah pass
                LblErrMsg = DisplayMsg(3001)
                
                sql = "update User_Setup set InvalidLogin=" & InvalidLogin + 1 & ", Last_Update = getdate(), Last_User = '" & userLogin & "' " & _
                    "where UserName ='" & txtUser & "'"  ' 'tambah jml invalidlogin
                Db.Execute sql
                
                rsUser.filter = ""
                rsUser.Requery
                
                If rsUser("invalidlogin") = 3 And StatusAdmin <> 1 Then  'jika 3 kale salah pass update Lock = 1
                    sql = "update user_Setup set locked = '1', Last_Update = getdate(), Last_User = '" & userLogin & "' " & _
                        "where userName='" & txtUser & "'"
                    Db.Execute sql
                End If
                
            Else 'jika input user dan pass bener
                Dim rscekuser As New Recordset
                Dim waktu
                waktu = Format(Now(), "MM/dd/yyyy hh:mm:ss")
                sql = "update user_Setup set Last_Login ='" & waktu & "', InvalidLogin = 0, Last_Update = getdate(), Last_User = '" & userLogin & "' " & _
                    "where UserName ='" & txtUser & "'"
                Db.Execute sql
                
                sql = "Insert Into log_history values('" & Format(Now, "yyyy-mm-dd") & "','" & Trim(txtUser.Text) & "','Login',Getdate()) "

                Db.Execute sql
                
                
                rsUser.Requery
                
                DoEvents
                On Error GoTo errHandler
                
                If txtmenu = "" Then
                    DoEvents
                    frmMainMenu.loadtree
                    frmMainMenu.Show
                    DoEvents
                    Me.Hide
                Else
                    If panggilForm(txtmenu) = 0 Then
                        DoEvents: Me.Hide
                    Else
                        LblErrMsg = DisplayMsg(3006)
                    End If
                End If
            End If
        End If
    End If
    
    Me.MousePointer = vbCustom
Exit Sub

errHandler:
If err.number = 440 Then
   MsgBox "Windows out of memory!" & vbCrLf & "Please close other application to free some memory!. This application will be closed! ", vbExclamation, "Kawai"
   End
Else
   MsgBox err.Description & vbCrLf & "This application will be closed! ", vbExclamation, "Kawai"
   End
End If
Screen.MousePointer = vbDefault
If rsUser.State <> adStateClosed Then rsUser.Close
End Sub

Private Sub cmdExit_Click()
   DoEvents
   Unload Me
   Call CloseReg
   End
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    sql = "update user_Setup set InvalidLogin = 0, Last_Update = getdate(), Last_User = '" & userLogin & "' " & _
        "where UserName ='" & userLogin & "'"
    Db.Execute sql
    If rsUser.State <> adStateClosed Then rsUser.Close: Set rsUser = Nothing
    If Db.State <> adStateClosed Then Db.Close: Set Db = Nothing
End Sub

Private Sub TxtMenu_KeyPress(KeyAscii As Integer)
    LblErrMsg = ""
    If KeyAscii = Asc("'") Then KeyAscii = 0
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    LblErrMsg = ""
    If KeyAscii = Asc("'") Then KeyAscii = 0
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub txtUser_KeyPress(KeyAscii As Integer)
    LblErrMsg = ""
    If KeyAscii = Asc("'") Then KeyAscii = 0
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub OpenReg()
    Dim SShortDate As String, SThousand As String, SDecimal As String, SMonThousand As String, SMonDecimal As String
    'Jika Terakhir kali abnormal Termination maka kembalikan Setting
    If GetSetting("NIC", "Setting", "Shutdown", "1") = "0" Then
        SShortDate = GetSetting("NIC", "Setting", "ShortDate", Get_locale(LOCALE_SSHORTDATE))
        SThousand = GetSetting("NIC", "Setting", "Thousand", Get_locale(LOCALE_STHOUSAND))
        SDecimal = GetSetting("NIC", "Setting", "Decimal", Get_locale(LOCALE_SDECIMAL))
        SMonDecimal = GetSetting("NIC", "Setting", "MonDecimal", Get_locale(LOCALE_SMONDECIMALSEP))
        SMonThousand = GetSetting("NIC", "Setting", "MonThousand", Get_locale(LOCALE_SMONTHOUSANDSEP))
        Set_locale LOCALE_SSHORTDATE, SShortDate
        Set_locale LOCALE_STHOUSAND, SThousand
        Set_locale LOCALE_SDECIMAL, SDecimal
        Set_locale LOCALE_SMONDECIMALSEP, SMonDecimal
        Set_locale LOCALE_SMONTHOUSANDSEP, SMonThousand
     End If

     'Jika Indonesia maka setLocale
     'If Get_locale(LOCALE_IDEFAULTCOUNTRY) = "62" Then
     If Get_locale(LOCALE_IDEFAULTCOUNTRY) <> "10" Then
        SaveSetting "NIC", "Setting", "ShortDate", Get_locale(LOCALE_SSHORTDATE)
        SaveSetting "NIC", "Setting", "Thousand", Get_locale(LOCALE_STHOUSAND)
        SaveSetting "NIC", "Setting", "Decimal", Get_locale(LOCALE_SDECIMAL)
        SaveSetting "NIC", "Setting", "MonDecimal", Get_locale(LOCALE_SMONDECIMALSEP)
        SaveSetting "NIC", "Setting", "MonThousand", Get_locale(LOCALE_SMONTHOUSANDSEP)

        SShortDate = "MM/dd/yyyy"
        SThousand = ","
        SDecimal = "."
        SMonDecimal = "."
        SMonThousand = ","

        Set_locale LOCALE_SSHORTDATE, SShortDate
        Set_locale LOCALE_STHOUSAND, SThousand
        Set_locale LOCALE_SDECIMAL, SDecimal
        Set_locale LOCALE_SMONDECIMALSEP, SMonDecimal
        Set_locale LOCALE_SMONTHOUSANDSEP, SMonThousand
        
        SaveSetting "NIC", "Setting", "Shutdown", "0"
      End If
End Sub

Private Sub CloseReg()
   'Regional Setting
   Dim SShortDate As String, SThousand As String, SDecimal As String, SMonThousand As String, SMonDecimal As String
   If Get_locale(LOCALE_IDEFAULTCOUNTRY) <> "10" Then
      SShortDate = GetSetting("NIC", "Setting", "ShortDate", Get_locale(LOCALE_SSHORTDATE))
      SThousand = GetSetting("NIC", "Setting", "Thousand", Get_locale(LOCALE_STHOUSAND))
      SDecimal = GetSetting("NIC", "Setting", "Decimal", Get_locale(LOCALE_SDECIMAL))
      SMonDecimal = GetSetting("NIC", "Setting", "MonDecimal", Get_locale(LOCALE_SMONDECIMALSEP))
      SMonThousand = GetSetting("NIC", "Setting", "MonThousand", Get_locale(LOCALE_SMONTHOUSANDSEP))

      Set_locale LOCALE_SSHORTDATE, SShortDate
      Set_locale LOCALE_STHOUSAND, SThousand
      Set_locale LOCALE_SDECIMAL, SDecimal
      Set_locale LOCALE_SMONDECIMALSEP, SMonDecimal
      Set_locale LOCALE_SMONTHOUSANDSEP, SMonThousand
   End If
   SaveSetting "NIC", "Setting", "Shutdown", "1"
End Sub
