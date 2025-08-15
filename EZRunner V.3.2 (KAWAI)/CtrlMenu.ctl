VERSION 5.00
Begin VB.UserControl CtrlMenu 
   BackColor       =   &H00FDDFE3&
   BackStyle       =   0  'Transparent
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1845
   LockControls    =   -1  'True
   ScaleHeight     =   405
   ScaleWidth      =   1845
   ToolboxBitmap   =   "CtrlMenu.ctx":0000
   Begin VB.TextBox TxtMenu 
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
      Left            =   900
      MaxLength       =   6
      TabIndex        =   0
      Top             =   60
      Width           =   855
   End
   Begin VB.Label LblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Menu ID"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   9
      Left            =   90
      TabIndex        =   1
      Top             =   105
      Width           =   720
   End
   Begin VB.Shape ShpMenu 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      Height          =   405
      Left            =   0
      Top             =   0
      Width           =   1845
   End
End
Attribute VB_Name = "CtrlMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim sFormName As String
Dim SFormDescription As String

Public Event ErrMessage(ErrMsg As String)

Public Property Get FormName() As String
    FormName = sFormName
End Property

Public Property Let FormName(ByVal vNewValue As String)
    sFormName = vNewValue
End Property

Public Property Get FormDescription() As String
    FormDescription = SFormDescription
End Property

Public Property Let FormDescription(ByVal vNewValue As String)
    SFormDescription = vNewValue
End Property

Public Sub OpenMenu(menuID As String, nmForm As String)
Dim stForm As Integer
    If txtmenu <> "" Then
        stForm = panggilForm(menuID, nmForm)
        If stForm = 0 Then 'Berhasil dipanggil
            RaiseEvent ErrMessage("")
        ElseIf stForm = 1 Then 'Invalid Menu
            RaiseEvent ErrMessage("Invalid Menu ID")
        ElseIf stForm = 2 Then 'Manggil diri sendiri
            RaiseEvent ErrMessage("This Form's Menu ID is " & menuID)
        End If
    End If
End Sub

Public Function GetMenuID(frmName As String) As String
    Dim sql As String
    Dim rst As Recordset
    
    If Db.State = adStateClosed Then Exit Function
    
    sql = "select * from user_menu where menu_name='" & frmName & "'"
    Set rst = New Recordset
    rst.Open sql, Db, adOpenKeyset, adLockOptimistic
    If rst.EOF = False Then GetMenuID = Trim(rst!menu_id)
End Function

Public Property Get MenuText() As String
    MenuText = GetMenuID(sFormName)
End Property

Public Property Let MenuText(ByVal vNewValue As String)
    txtmenu.Text = vNewValue
End Property

Private Sub TxtMenu_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = Asc("'") Then KeyAscii = 0
    If KeyAscii = 13 Then Call OpenMenu(txtmenu.Text, sFormName)
End Sub

