Attribute VB_Name = "MdlSetting"
Option Explicit
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'========================================================================================
'Setting System
'========================================================================================
'System
            Public Const gb_Simulation = False
            
'Part Supply [Unschedule]
            Public Const gb_AllowClearInputArea_PartSupplyUnschedule = False
            
'Part Receipt [Schedule]
            Public Const gb_AllowInputWithoutFix_PartReceiptSchedule = False
            
'Material Supply Automatic
            Public Const gb_AllowSetDefaultSupplyQty_MaterialSupplyAutomatic = True
            
'Material Supply Request Automatic
            Public Const gb_WarehouseReferToItem_MaterialSupplyRequestAutomatic = False
            
'Invoice Create
            Public Const gb_AllowMultipleDO_InvoiceCreate As Boolean = True
            Public Const gb_InvoiceReferToDO_InvoiceCreate As Boolean = True
'========================================================================================
            
'========================================================================================
'Setting Region
'========================================================================================
Declare Function GetLocaleInfo Lib "kernel32" Alias _
"GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, _
ByVal lpLCData As String, ByVal cchData As Long) As Long
            
Declare Function SetLocaleInfo Lib "kernel32" Alias _
"SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, _
ByVal lpLCData As String) As Boolean
        
Declare Function GetUserDefaultLCID% Lib "kernel32" ()
        
Public Const LOCALE_ICOUNTRY = &H5
Public Const LOCALE_IDEFAULTCOUNTRY = &HA
Public Const LOCALE_IDEFAULTLANGUAGE = &H9
Public Const LOCALE_SCOUNTRY = &H6
Public Const LOCALE_SDATE = &H1D
Public Const LOCALE_SDECIMAL = &HE
Public Const LOCALE_SENGCOUNTRY = &H1002
Public Const LOCALE_SENGLANGUAGE = &H1001
Public Const LOCALE_SGROUPING = &H10
Public Const LOCALE_SLANGUAGE = &H2
Public Const LOCALE_SSHORTDATE = &H1F
Public Const LOCALE_STHOUSAND = &HF
Public Const LOCALE_STIME = &H1E
Public Const LOCALE_STIMEFORMAT = &H1003
Public Const LOCALE_SMONTHOUSANDSEP = &H17
Public Const LOCALE_SMONDECIMALSEP = &H16
Public LightGreen As ColorConstants
        
Public Sub up_InitSimulation(frm As Form)
LightGreen = RGB(204, 255, 204)
Dim obj
frm.BackColor = LightGreen
For Each obj In frm.Controls
    
   If TypeOf obj Is Frame Then
        If obj.BackColor = &HFDDFE3 Then obj.BackColor = LightGreen
    End If
    If TypeOf obj Is Label Then
        If obj.BackColor = &HFDDFE3 Then obj.BackColor = LightGreen
    End If
    If TypeOf obj Is SSTab Then
        obj.BackColor = LightGreen
    End If
    If TypeOf obj Is TextBox Then
        If obj.BackColor = &HFDDFE3 Then obj.BackColor = LightGreen
    End If
    If TypeOf obj Is CheckBox Then
        If obj.BackColor = &HFDDFE3 Then obj.BackColor = LightGreen
    End If
    If TypeOf obj Is OptionButton Then
        If obj.BackColor = &HFDDFE3 Then obj.BackColor = LightGreen
    End If
Next
End Sub


Public Function Get_locale(Lc_Type As Long) As String  ' Retrieve the regional setting

      Dim Symbol As String
      Dim iRet1 As Long
      Dim iRet2 As Long
      Dim lpLCDataVar As String
      Dim Pos As Integer
      Dim Locale As Long
      
      Locale = GetUserDefaultLCID()

      
      iRet1 = GetLocaleInfo(Locale, Lc_Type, lpLCDataVar, 0)
      Symbol = String$(iRet1, 0)
      
      iRet2 = GetLocaleInfo(Locale, Lc_Type, Symbol, iRet1)
      Pos = InStr(Symbol, Chr$(0))
      If Pos > 0 Then
           Symbol = Left$(Symbol, Pos - 1)
           Get_locale = Symbol
      End If

End Function

Public Sub Set_locale(Lc_Type As Long, LcData As String)   'Change the regional setting

      Dim iRet As Long
      Dim Locale As Long

      Locale = GetUserDefaultLCID() 'Get user Locale ID
      'iRet = SetLocaleInfo(Locale, LOCALE_SSHORTDATE, "MM/dd/yyyy")
      'iRet = SetLocaleInfo(Locale, LOCALE_STHOUSAND, ",")
      'iRet = SetLocaleInfo(Locale, LOCALE_SDECIMAL, ".")
      iRet = SetLocaleInfo(Locale, Lc_Type, LcData)

End Sub

Public Function fc_Decrypt(psFile As String) As String
   On Local Error GoTo err_handler
   Screen.MousePointer = vbHourglass
   
   Dim strCode As String
   Dim strPos As Integer
   Dim strChar As String
    
   fc_Decrypt = ""
    
   Do
      strPos = Left(psFile, 1)
      psFile = Mid(psFile, 2)
      strCode = Left(psFile, strPos)
      psFile = Mid(psFile, Len(strCode) + 1)
      fc_Decrypt = fc_Decrypt & Chr(strCode)
   Loop Until psFile = ""

err_exit:
   Screen.MousePointer = vbDefault
   Exit Function
err_handler:
   Screen.MousePointer = vbDefault
   MsgBox "Err. Number : " & err.number & vbNewLine & "Err. Description : " & err.Description, vbCritical, "Error decrypting file"
   err.clear
   Resume err_exit
End Function

Public Function fc_Encrypt(psFile As String) As String
   On Local Error GoTo err_handler
   Screen.MousePointer = vbHourglass
   
   Dim strChar As String
   Dim intCount   As Integer
    
   fc_Encrypt = ""
    
   For intCount = 1 To Len(psFile)
      strChar = Asc(Mid(psFile, intCount, 1))
      fc_Encrypt = fc_Encrypt & Len(strChar) & strChar
   Next intCount
    
err_exit:
   Screen.MousePointer = vbDefault
   Exit Function
err_handler:
   Screen.MousePointer = vbDefault
   MsgBox "Err. Number : " & err.number & vbNewLine & "Err. Description : " & err.Description, vbCritical, "Error encrypting file"
   err.clear
   Resume err_exit
End Function
