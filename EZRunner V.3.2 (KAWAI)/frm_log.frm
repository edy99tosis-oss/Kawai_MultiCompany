VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_log 
   BackColor       =   &H00FDDFE3&
   Caption         =   "Part/Material Receipt List"
   ClientHeight    =   4455
   ClientLeft      =   1335
   ClientTop       =   2775
   ClientWidth     =   8220
   Icon            =   "frm_log.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   8220
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "E&xcel"
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
      Left            =   6780
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3675
      Width           =   1035
   End
   Begin EZRunnerv3.CtrlMenu CtrlMenu1 
      Height          =   420
      Left            =   5989
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   285
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   741
   End
   Begin VB.CommandButton Cmd_Save 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Sub &Menu"
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
      Index           =   8
      Left            =   379
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3675
      Width           =   1125
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   379
      TabIndex        =   5
      Top             =   2880
      Width           =   7470
      Begin VB.Label LblErrMsg 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "LblErrMsg"
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
         TabIndex        =   6
         Top             =   195
         Width           =   7260
      End
   End
   Begin MSComCtl2.DTPicker DMonth 
      Height          =   315
      Left            =   2535
      TabIndex        =   0
      Top             =   1860
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
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
      Format          =   145227779
      CurrentDate     =   37798
   End
   Begin MSComCtl2.DTPicker Dmonth2 
      Height          =   315
      Left            =   4605
      TabIndex        =   1
      Top             =   1860
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
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
      Format          =   145227779
      CurrentDate     =   37798
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date (Month)"
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
      Left            =   930
      TabIndex        =   9
      Top             =   1920
      Width           =   1125
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "to"
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
      Left            =   4335
      TabIndex        =   8
      Top             =   1920
      Width           =   165
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Log User History"
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
      Height          =   390
      Left            =   375
      TabIndex        =   7
      Top             =   570
      Width           =   7470
   End
End
Attribute VB_Name = "frm_log"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Amounti As Double
Dim PPni As Double
Dim grandI As Double

Dim bteHakPrice As Byte



Private Sub Cmd_Save_Click(Index As Integer)
    frmMainMenu.Show
    Unload Me
End Sub

Private Sub Command1_Click() 'EXCEL
    Dim xlapp As New Excel.application
    Dim rsCek As New Recordset, Idx As Long, tempi As String, tempcls As String
    Dim bolcls As Boolean, bolcur As Boolean
    Dim rsCompany As New Recordset
    Dim AmountCls As Double, PPnCls As Double, GrandCls As Double
    
    MousePointer = vbHourglass
        sql = " select * from log_history where Tanggal between '" & Format(DMonth.Value, "yyyy-mm-dd") & "' and '" & Format(Dmonth2.Value, "yyyy-mm-dd") & "' " & _
              " order by Tanggal,UserID"
              
    If rsCek.State <> adStateClosed Then rsCek.Close
    'rsCek.CursorLocation = adUseClient
    rsCek.Open sql, Db, adOpenDynamic, adLockOptimistic

    If rsCek.EOF Then
        LblErrMsg.Caption = DisplayMsg(4006)
    Else
            
        With xlapp
            
            sql = "select rtrim(company_name) company_name, rtrim(address1) Address1, rtrim(Address2) Address2, rtrim(Province) Province, rtrim(city) City, Rtrim(Postal_Code) POstal_Code, Rtrim(phone1) Phone1, Rtrim(phone2) Phone2,rtrim(fax) Fax  From company_profile "
            If rsCompany.State <> adStateClosed Then rsCompany.Close
            rsCompany.Open sql, Db, adOpenDynamic, adLockOptimistic
            If rsCompany.EOF Then MousePointer = vbDefault: Exit Sub
            .Workbooks.Add
            
            .Range("a2", "d2").Merge
            .Range("a2") = "Log User History"
            .Range("a4", "j4").Merge
            .Range("a4") = rsCompany!company_name
            .Range("a5", "j5").Merge
            .Range("a5") = rsCompany!address1 & " " & rsCompany!address2 & " " & rsCompany!City
            .Range("a6", "j6").Merge
            .Range("a6") = "Phone : " & rsCompany!phone1 & " " & rsCompany!phone2
            .Range("a7", "j7").Merge
            .Range("a7") = "Fax : " & rsCompany!fax
            
            .Range("b9", "d9").Merge
            .Range("a9") = "Date"
            .Range("b9") = ": " & Format(DMonth, "dd MMMM YYYY") & " to " & Format(Dmonth2, "dd MMMM YYYY")
            
            Idx = 11
            
            Do While Not rsCek.EOF
               
        
                If Idx = 11 Then
                                           
                    Idx = Idx + 1
                    .Range("a" & Idx).HorizontalAlignment = xlCenter
                    .Range("a" & Idx) = "Tanggal"
                    .Range("b" & Idx) = "User Name"
                    .Range("c" & Idx) = "Menu"
                    .Range("d" & Idx) = "Last Update"
                    .Range("a" & Idx, "D" & Idx).Borders(xlEdgeTop).LineStyle = xlContinuous
                    .Range("a" & Idx, "D" & Idx).Borders(xlEdgeBottom).LineStyle = xlContinuous
                    Idx = Idx + 1
                End If
        
                Idx = Idx
                'Content
                .Range("a" & Idx).HorizontalAlignment = xlCenter
                .Range("a" & Idx) = Format(rsCek!Tanggal, "DD-MMM-YYYY")
                .Range("b" & Idx) = Trim(rsCek!UserID)
                .Range("c" & Idx) = Trim(rsCek!MenuDesc)
                .Range("d" & Idx) = "'" & Trim(rsCek!Last_Update)
                
                Idx = Idx + 1
                rsCek.MoveNext
            Loop

            
            .Range("a2", "j2").Columns.Font.Name = "Arial"
            .Range("a2", "j2").Columns.Font.Size = "10"
            .Range("a2", "j2").Columns.Font.Bold = True
            .Range("a2", "j2").HorizontalAlignment = xlCenter
            
            .Range("a1", "j" & Idx + 3).Columns.AutoFit
            .Range("A1").Select
            
            .ActiveSheet.PageSetup.LeftMargin = application.InchesToPoints(0.4)
            .ActiveSheet.PageSetup.RightMargin = application.InchesToPoints(0.04)
            .ActiveSheet.PageSetup.Orientation = 2
            .WindowState = xlMaximized
            .Visible = True
        
        End With
    End If
    MousePointer = vbDefault
    
End Sub

Private Sub CtrlMenu1_ErrMessage(ErrMsg As String)
If ErrMsg = "" Then
    Unload Me
Else
    LblErrMsg.Caption = ErrMsg
End If
End Sub

Private Sub DMonth_Change()
If CDate(Dmonth2) < CDate(DMonth) Then
      LblErrMsg.Caption = DisplayMsg(4068)
      Exit Sub
   Else
      LblErrMsg.Caption = ""
   End If
End Sub

Private Sub DMonth_Click()
If CDate(Dmonth2) < CDate(DMonth) Then
      LblErrMsg.Caption = DisplayMsg(4068)
      Exit Sub
   Else
      LblErrMsg.Caption = ""
   End If
End Sub


Private Sub Dmonth2_Change()
If CDate(Dmonth2) < CDate(DMonth) Then
      LblErrMsg.Caption = DisplayMsg(4066)
      Exit Sub
   Else
      LblErrMsg.Caption = ""
   End If
End Sub

Private Sub Dmonth2_Click()
If CDate(Dmonth2) < CDate(DMonth) Then
      LblErrMsg.Caption = DisplayMsg(4066)
      Exit Sub
   Else
      LblErrMsg.Caption = ""
   End If
End Sub

Private Sub Form_Load()
Dim RsW As New Recordset
Dim ir As Integer
If gb_Simulation = True Then Call up_InitSimulation(Me)
LblErrMsg = ""
CtrlMenu1.FormName = Me.Name
Me.Caption = Me.Caption & " (Menu ID : " & frmcode(Me.Name) & ")"

DMonth = Format(Date, "dd MMM yyyy")
Dmonth2 = Format(Date, "dd MMM yyyy")

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then Cancel = 1
End Sub
