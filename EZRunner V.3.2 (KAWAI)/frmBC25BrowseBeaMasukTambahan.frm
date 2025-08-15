VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmBC25BrowseBeaMasukTambahan 
   BackColor       =   &H00FDDFE3&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bea Masuk Tambahan"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9465
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBC25BrowseBeaMasukTambahan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   9465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNoSeriBahanBaku 
      Appearance      =   0  'Flat
      Height          =   350
      Left            =   2520
      TabIndex        =   59
      Top             =   7560
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDDFE3&
      Caption         =   "Identitas Barang"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   50
      Top             =   120
      Width           =   9255
      Begin VB.TextBox txtNomorHS 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   54
         Top             =   360
         Width           =   2265
      End
      Begin VB.TextBox txtUraianBarang 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   53
         Top             =   840
         Width           =   7305
      End
      Begin VB.TextBox txtCIF 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   52
         Top             =   1320
         Width           =   2265
      End
      Begin VB.TextBox txtCIFRupiah 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   6720
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   51
         Top             =   1320
         Width           =   2265
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor HS"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   58
         Top             =   435
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Uraian Barang"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   57
         Top             =   915
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Harga CIF"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   56
         Top             =   1395
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CIF Rp."
         Height          =   195
         Index           =   3
         Left            =   5640
         TabIndex        =   55
         Top             =   1395
         Width           =   645
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDDFE3&
      Caption         =   "Bea Masuk"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   34
      Top             =   2160
      Width           =   9255
      Begin VB.TextBox txtBMBayar 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   350
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   38
         Text            =   "0.00"
         Top             =   1320
         Width           =   2625
      End
      Begin VB.TextBox txtBesarTarif 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   350
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   37
         Text            =   "0.00"
         Top             =   840
         Width           =   705
      End
      Begin VB.TextBox txtTarifFasilitas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   350
         Left            =   7320
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   36
         Text            =   "0.00"
         Top             =   840
         Width           =   705
      End
      Begin VB.TextBox txtBMFasilitas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   350
         Left            =   6360
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   35
         Text            =   "0"
         Top             =   1320
         Width           =   2625
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tarif Fasilitas"
         Height          =   195
         Index           =   4
         Left            =   4560
         TabIndex        =   49
         Top             =   435
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BM Bayar"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   48
         Top             =   1395
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Besar Tarif"
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   47
         Top             =   915
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Tarif BM"
         Height          =   195
         Index           =   7
         Left            =   240
         TabIndex        =   46
         Top             =   435
         Width           =   1185
      End
      Begin MSForms.ComboBox cboTarifFasilitas 
         Height          =   345
         Left            =   4560
         TabIndex        =   45
         Tag             =   "TTFF*/"
         Top             =   840
         Width           =   2655
         VariousPropertyBits=   746604575
         BackColor       =   -2147483648
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4683;609"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboJenisTarifBM 
         Height          =   345
         Left            =   1800
         TabIndex        =   44
         Tag             =   "TTFF*/"
         Top             =   360
         Width           =   2655
         VariousPropertyBits=   746604575
         BackColor       =   -2147483648
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4683;609"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rp."
         Height          =   300
         Index           =   8
         Left            =   1440
         TabIndex        =   43
         Top             =   1395
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Index           =   9
         Left            =   2640
         TabIndex        =   42
         Top             =   915
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Index           =   10
         Left            =   8160
         TabIndex        =   41
         Top             =   915
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BM Fasilitas"
         Height          =   195
         Index           =   11
         Left            =   4560
         TabIndex        =   40
         Top             =   1395
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rp."
         Height          =   255
         Index           =   12
         Left            =   6000
         TabIndex        =   39
         Top             =   1368
         Width           =   285
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FDDFE3&
      Caption         =   "Tambahan Bea Masuk"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   6
      Top             =   4200
      Width           =   9255
      Begin VB.CheckBox chkBMADs 
         BackColor       =   &H00FDDFE3&
         Caption         =   "BMADs"
         Height          =   255
         Left            =   1920
         TabIndex        =   23
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtBMADs 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   3240
         MaxLength       =   5
         TabIndex        =   22
         Text            =   "0.00"
         Top             =   312
         Width           =   1065
      End
      Begin VB.TextBox txtBMADsRupiah 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   350
         Left            =   5520
         MaxLength       =   20
         TabIndex        =   21
         Text            =   "0"
         Top             =   312
         Width           =   2745
      End
      Begin VB.CommandButton cmdBMADs 
         BackColor       =   &H0080FFFF&
         Caption         =   "X"
         Height          =   375
         Left            =   8520
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   300
         Width           =   375
      End
      Begin VB.CheckBox chkBMIS 
         BackColor       =   &H00FDDFE3&
         Caption         =   "BMIs"
         Height          =   255
         Left            =   1920
         TabIndex        =   19
         Top             =   828
         Width           =   1215
      End
      Begin VB.TextBox txtBMIs 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   3240
         MaxLength       =   5
         TabIndex        =   18
         Text            =   "0.00"
         Top             =   780
         Width           =   1065
      End
      Begin VB.TextBox txtBMIsRupiah 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   350
         Left            =   5520
         MaxLength       =   20
         TabIndex        =   17
         Text            =   "0"
         Top             =   780
         Width           =   2745
      End
      Begin VB.CommandButton cmdBMIs 
         BackColor       =   &H0080FFFF&
         Caption         =   "X"
         Height          =   375
         Left            =   8520
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   768
         Width           =   375
      End
      Begin VB.CheckBox chkBMTPs 
         BackColor       =   &H00FDDFE3&
         Caption         =   "BMTPs"
         Height          =   255
         Left            =   1920
         TabIndex        =   15
         Top             =   1298
         Width           =   1215
      End
      Begin VB.TextBox txtBMTPs 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   3240
         MaxLength       =   5
         TabIndex        =   14
         Text            =   "0.00"
         Top             =   1250
         Width           =   1065
      End
      Begin VB.TextBox txtBMTPsRupiah 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   350
         Left            =   5520
         MaxLength       =   20
         TabIndex        =   13
         Text            =   "0"
         Top             =   1250
         Width           =   2745
      End
      Begin VB.CommandButton cmdBMTPs 
         BackColor       =   &H0080FFFF&
         Caption         =   "X"
         Height          =   375
         Left            =   8520
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1238
         Width           =   375
      End
      Begin VB.CheckBox chkBMPs 
         BackColor       =   &H00FDDFE3&
         Caption         =   "BMPs"
         Height          =   255
         Left            =   1920
         TabIndex        =   11
         Top             =   1740
         Width           =   1215
      End
      Begin VB.TextBox txtBMPs 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   3240
         MaxLength       =   5
         TabIndex        =   10
         Text            =   "0.00"
         Top             =   1680
         Width           =   1065
      End
      Begin VB.TextBox txtBMPsRupiah 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   350
         Left            =   5520
         MaxLength       =   20
         TabIndex        =   9
         Text            =   "0"
         Top             =   1695
         Width           =   2745
      End
      Begin VB.CommandButton cmdBMPs 
         BackColor       =   &H0080FFFF&
         Caption         =   "X"
         Height          =   375
         Left            =   8520
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox txtTotalBM 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   350
         Left            =   5520
         MaxLength       =   20
         TabIndex        =   7
         Text            =   "0"
         Top             =   2160
         Width           =   2745
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BMAD"
         Height          =   195
         Index           =   14
         Left            =   240
         TabIndex        =   33
         Top             =   360
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "% = Rp"
         Height          =   195
         Index           =   15
         Left            =   4560
         TabIndex        =   32
         Top             =   390
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BMI"
         Height          =   195
         Index           =   13
         Left            =   240
         TabIndex        =   31
         Top             =   858
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "% = Rp"
         Height          =   195
         Index           =   16
         Left            =   4560
         TabIndex        =   30
         Top             =   858
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BMTP"
         Height          =   195
         Index           =   17
         Left            =   240
         TabIndex        =   29
         Top             =   1328
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "% = Rp"
         Height          =   195
         Index           =   18
         Left            =   4560
         TabIndex        =   28
         Top             =   1328
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BMP"
         Height          =   195
         Index           =   19
         Left            =   240
         TabIndex        =   27
         Top             =   1740
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "% = Rp"
         Height          =   195
         Index           =   20
         Left            =   4560
         TabIndex        =   26
         Top             =   1770
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total BM Tambahan"
         Height          =   195
         Index           =   21
         Left            =   240
         TabIndex        =   25
         Top             =   2235
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rp"
         Height          =   195
         Index           =   23
         Left            =   4995
         TabIndex        =   24
         Top             =   2205
         Width           =   225
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FDDFE3&
      Height          =   555
      Left            =   120
      TabIndex        =   4
      Tag             =   "TFTT*/"
      Top             =   6960
      Width           =   9315
      Begin VB.Label LblerrMsg 
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
         TabIndex        =   5
         Tag             =   "TTFF*/"
         Top             =   195
         Width           =   9090
      End
   End
   Begin VB.CommandButton cmdSubmit 
      BackColor       =   &H0080FFFF&
      Caption         =   "Submit"
      Height          =   375
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7680
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H0080FFFF&
      Caption         =   "Close"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7680
      Width           =   975
   End
   Begin VB.TextBox txtNoPengajuan 
      Appearance      =   0  'Flat
      Height          =   350
      Left            =   4920
      MaxLength       =   50
      TabIndex        =   1
      Top             =   7560
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.TextBox txtNoSeri 
      Appearance      =   0  'Flat
      Height          =   350
      Left            =   3720
      MaxLength       =   5
      TabIndex        =   0
      Top             =   7560
      Visible         =   0   'False
      Width           =   1065
   End
End
Attribute VB_Name = "frmBC25BrowseBeaMasukTambahan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Sub up_LoadData(pNoPengajuan As String, pNoSeri As Integer)
    Dim RS As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim NomorHS As String
    Dim cekBox As String
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC25DetailBeaMasukTambahan_Sel"
    
    cmd.Parameters.append cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, Replace(pNoPengajuan, "-", ""))
    cmd.Parameters.append cmd.CreateParameter("NoSeri", adInteger, adParamInput, 5, pNoSeri)
        
    Set RS = cmd.Execute
    
    If Not RS.EOF Then
        NomorHS = IIf(IsNull(RS.Fields("POS_TARIF")), "", RS.Fields("POS_TARIF"))
        
        txtNomorHS = Replace(NomorHS, ".", "")
        txtNomorHS = Mid(txtNomorHS.Text, 1, 10)
        If txtNomorHS <> "" Then
            txtNomorHS = Left(txtNomorHS.Text, 4) & "." & Mid(txtNomorHS.Text, 5, 2) & "." & Mid(txtNomorHS.Text, 7, 2) & "." & Mid(txtNomorHS.Text, 9, 2)
        End If
        
        txtUraianBarang = IIf(IsNull(RS.Fields("URAIAN")), "", RS.Fields("URAIAN"))
        txtCIF = Format(IIf(IsNull(RS.Fields("CIF")), 0, RS.Fields("CIF")), "#,0.00")
        txtCIFRupiah = Format(IIf(IsNull(RS.Fields("CIF_Rupiah")), 0, RS.Fields("CIF_Rupiah")), "#,0.00")
        
        cboJenisTarifBM = IIf(IsNull(RS.Fields("JENISTARIFBM")), "", RS.Fields("JENISTARIFBM"))
        cboTarifFasilitas = IIf(IsNull(RS.Fields("JENISFASILITASBM")), "", RS.Fields("JENISFASILITASBM"))
        txtBesarTarif = Format(IIf(IsNull(RS.Fields("TARIFBM")), 0, RS.Fields("TARIFBM")), "#,0.00")
        txtBMBayar = Format(IIf(IsNull(RS.Fields("NILAIFASILITASBM")), 0, RS.Fields("NILAIFASILITASBM")), "#,0.00")
        txtBMFasilitas = Format(IIf(IsNull(RS.Fields("NILAIBAYARBM")), 0, RS.Fields("NILAIBAYARBM")), "#,0.00")
        
        cekBox = IIf(IsNull(RS.Fields("FLAGBMAD")), "", RS.Fields("FLAGBMAD"))
        If cekBox = "Y" Then
            chkBMADs.Value = True
        Else
            chkBMADs.Value = False
        End If
        txtBMADs = Format(IIf(IsNull(RS.Fields("TARIFBMAD")), 0, RS.Fields("TARIFBMAD")), "#,0.00")
        txtBMADsRupiah = Format(IIf(IsNull(RS.Fields("NILAIFASILITASBMAD")), 0, RS.Fields("NILAIFASILITASBMAD")), "#,0.00")
        
        cekBox = IIf(IsNull(RS.Fields("FLAGBMI")), "", RS.Fields("FLAGBMI"))
        If cekBox = "Y" Then
            chkBMIS.Value = True
        Else
            chkBMIS.Value = False
        End If
        txtBMIs = Format(IIf(IsNull(RS.Fields("TARIFBMI")), 0, RS.Fields("TARIFBMI")), "#,0.00")
        txtBMIsRupiah = Format(IIf(IsNull(RS.Fields("NILAIFASILITASBMI")), 0, RS.Fields("NILAIFASILITASBMI")), "#,0.00")
        
        cekBox = IIf(IsNull(RS.Fields("FLAGBMTP")), "", RS.Fields("FLAGBMTP"))
        If cekBox = "Y" Then
            chkBMTPs.Value = True
        Else
            chkBMTPs.Value = False
        End If
        txtBMTPs = Format(IIf(IsNull(RS.Fields("TARIFBMTP")), 0, RS.Fields("TARIFBMTP")), "#,0.00")
        txtBMTPsRupiah = Format(IIf(IsNull(RS.Fields("NILAIFASILITASBMTP")), 0, RS.Fields("NILAIFASILITASBMTP")), "#,0.00")
        
        cekBox = IIf(IsNull(RS.Fields("FLAGBMP")), "", RS.Fields("FLAGBMP"))
        If cekBox = "Y" Then
            chkBMPs.Value = True
        Else
            chkBMPs.Value = False
        End If
        txtBMPs = Format(IIf(IsNull(RS.Fields("TARIFBMP")), 0, RS.Fields("TARIFBMP")), "#,0.00")
        txtBMPsRupiah = Format(IIf(IsNull(RS.Fields("NILAIFASILITASBMP")), 0, RS.Fields("NILAIFASILITASBMP")), "#,0.00")
        
        txtTotalBM = Format(CDbl(txtBMADsRupiah) + CDbl(txtBMIsRupiah) + CDbl(txtBMTPsRupiah) + CDbl(txtBMPsRupiah), "#,0.00")
    End If
End Sub

Public Sub up_LoadDataBahanBaku(pNoPengajuan As String, pNoSeriBarang As Integer, pNoSeriBahanBaku As Integer, pKodeAsal As Integer)
    Dim RS As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim NomorHS As String
    Dim cekBox As String
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC25DetailBeaMasukTambahanBahanBaku_Sel"
    
    cmd.Parameters.append cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, Replace(pNoPengajuan, "-", ""))
    cmd.Parameters.append cmd.CreateParameter("NoSeriBarang", adInteger, adParamInput, 5, pNoSeriBarang)
    cmd.Parameters.append cmd.CreateParameter("NoSeriBahanBaku", adInteger, adParamInput, 5, pNoSeriBarang)
    cmd.Parameters.append cmd.CreateParameter("KodeAsal", adInteger, adParamInput, 5, pKodeAsal)
        
    Set RS = cmd.Execute
    
    If Not RS.EOF Then
        NomorHS = IIf(IsNull(RS.Fields("POS_TARIF")), "", RS.Fields("POS_TARIF"))
        
        txtNomorHS = Replace(NomorHS, ".", "")
        txtNomorHS = Mid(txtNomorHS.Text, 1, 10)
        If txtNomorHS <> "" Then
            txtNomorHS = Left(txtNomorHS.Text, 4) & "." & Mid(txtNomorHS.Text, 5, 2) & "." & Mid(txtNomorHS.Text, 7, 2) & "." & Mid(txtNomorHS.Text, 9, 2)
        End If
        
        txtUraianBarang = IIf(IsNull(RS.Fields("URAIAN")), "", RS.Fields("URAIAN"))
        txtCIF = Format(IIf(IsNull(RS.Fields("CIF")), 0, RS.Fields("CIF")), "#,0.00")
        txtCIFRupiah = Format(IIf(IsNull(RS.Fields("CIF_Rupiah")), 0, RS.Fields("CIF_Rupiah")), "#,0.00")
        
        cboJenisTarifBM = IIf(IsNull(RS.Fields("JENISTARIFBM")), "", RS.Fields("JENISTARIFBM"))
        cboTarifFasilitas = IIf(IsNull(RS.Fields("JENISFASILITASBM")), "", RS.Fields("JENISFASILITASBM"))
        txtBesarTarif = Format(IIf(IsNull(RS.Fields("TARIFBM")), 0, RS.Fields("TARIFBM")), "#,0.00")
        
        
        txtBMBayar = Format(IIf(IsNull(RS.Fields("NILAIFASILITASBM")), 0, RS.Fields("NILAIFASILITASBM")), "#,0.00")
        txtBMFasilitas = Format(IIf(IsNull(RS.Fields("NILAIBAYARBM")), 0, RS.Fields("NILAIBAYARBM")), "#,0.00")
        
        cekBox = IIf(IsNull(RS.Fields("FLAGBMAD")), "", RS.Fields("FLAGBMAD"))
        If cekBox = "Y" Then
            chkBMADs.Value = True
        Else
            chkBMADs.Value = False
        End If
        txtBMADs = Format(IIf(IsNull(RS.Fields("TARIFBMAD")), 0, RS.Fields("TARIFBMAD")), "#,0.00")
        txtBMADsRupiah = Format(IIf(IsNull(RS.Fields("NILAIFASILITASBMAD")), 0, RS.Fields("NILAIFASILITASBMAD")), "#,0.00")
        
        cekBox = IIf(IsNull(RS.Fields("FLAGBMI")), "", RS.Fields("FLAGBMI"))
        If cekBox = "Y" Then
            chkBMIS.Value = True
        Else
            chkBMIS.Value = False
        End If
        txtBMIs = Format(IIf(IsNull(RS.Fields("TARIFBMI")), 0, RS.Fields("TARIFBMI")), "#,0.00")
        txtBMIsRupiah = Format(IIf(IsNull(RS.Fields("NILAIFASILITASBMI")), 0, RS.Fields("NILAIFASILITASBMI")), "#,0.00")
        
        cekBox = IIf(IsNull(RS.Fields("FLAGBMTP")), "", RS.Fields("FLAGBMTP"))
        If cekBox = "Y" Then
            chkBMTPs.Value = True
        Else
            chkBMTPs.Value = False
        End If
        txtBMTPs = Format(IIf(IsNull(RS.Fields("TARIFBMTP")), 0, RS.Fields("TARIFBMTP")), "#,0.00")
        txtBMTPsRupiah = Format(IIf(IsNull(RS.Fields("NILAIFASILITASBMTP")), 0, RS.Fields("NILAIFASILITASBMTP")), "#,0.00")
        
        cekBox = IIf(IsNull(RS.Fields("FLAGBMP")), "", RS.Fields("FLAGBMP"))
        If cekBox = "Y" Then
            chkBMPs.Value = True
        Else
            chkBMPs.Value = False
        End If
        txtBMPs = Format(IIf(IsNull(RS.Fields("TARIFBMP")), 0, RS.Fields("TARIFBMP")), "#,0.00")
        txtBMPsRupiah = Format(IIf(IsNull(RS.Fields("NILAIFASILITASBMP")), 0, RS.Fields("NILAIFASILITASBMP")), "#,0.00")
        
        txtTotalBM = Format(CDbl(txtBMADsRupiah) + CDbl(txtBMIsRupiah) + CDbl(txtBMTPsRupiah) + CDbl(txtBMPsRupiah), "#,0.00")
    End If
End Sub


Private Sub up_Delete(pKodeTarif As String)
    Dim RS As ADODB.Recordset
    Dim cmd As ADODB.Command
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC25DetailBeaMasukTambahan_Del"
    
    cmd.Parameters.append cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
    cmd.Parameters.append cmd.CreateParameter("NoSeri", adInteger, adParamInput, , txtNoSeri)
    cmd.Parameters.append cmd.CreateParameter("JenisTarif", adVarChar, adParamInput, 10, pKodeTarif)
    cmd.Execute
    
    up_LoadData txtNoPengajuan, txtNoSeri
    
    LblerrMsg.Caption = DisplayMsg(1201)
End Sub

Private Sub up_DeleteBahanBaku(pKodeTarif As String)
    Dim RS As ADODB.Recordset
    Dim cmd As ADODB.Command
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC25DetailBeaMasukTambahanBahanBaku_Del"
    
    cmd.Parameters.append cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
    cmd.Parameters.append cmd.CreateParameter("NoSeriBarang", adInteger, adParamInput, , txtNoSeri)
    cmd.Parameters.append cmd.CreateParameter("NoSeriBahanBaku", adInteger, adParamInput, , txtNoSeriBahanBaku)
    cmd.Parameters.append cmd.CreateParameter("KodeAsal", adVarChar, adParamInput, 10, 0)
    cmd.Parameters.append cmd.CreateParameter("JenisTarif", adVarChar, adParamInput, 10, pKodeTarif)
    cmd.Execute
    
    up_LoadDataBahanBaku txtNoPengajuan, txtNoSeri, txtNoSeriBahanBaku, 0
    
    LblerrMsg.Caption = DisplayMsg(1201)
End Sub


Private Sub up_SaveDataBahanBaku()
Dim RS As ADODB.Recordset
Dim cmd As ADODB.Command
Dim Y As Integer
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim prm3 As ADODB.Parameter
Dim prm4 As ADODB.Parameter
Dim prm5 As ADODB.Parameter
Dim prm6 As ADODB.Parameter
Dim prm7 As ADODB.Parameter
Dim prm8 As ADODB.Parameter
Dim prm9 As ADODB.Parameter
Dim prm10 As ADODB.Parameter
Dim prm11 As ADODB.Parameter
Dim prm12 As ADODB.Parameter

Dim cekFlag As String

'###################################################################################################
'1. BMAD

Dim NilaiFasilitas As Double
Dim NilaiBayar As Double

If Left(cboTarifFasilitas, 1) = "0" Then
    NilaiFasilitas = 0
    NilaiBayar = CDbl(txtBMADsRupiah)
ElseIf Left(cboTarifFasilitas, 1) = "6" Then
    NilaiFasilitas = CDbl(txtBMADsRupiah)
    NilaiBayar = 0
End If

Set cmd = New ADODB.Command
cmd.CommandType = adCmdStoredProc
cmd.CommandTimeout = 0
cmd.ActiveConnection = Db
cmd.CommandText = "sp_BC25DetailBeaMasukTambahanBahanBaku_Upd"
    
Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
cmd.Parameters.append prm1
Set prm2 = cmd.CreateParameter("NoSeri", adInteger, adParamInput, 5, txtNoSeri)
cmd.Parameters.append prm2
Set prm3 = cmd.CreateParameter("JenisTarif", adVarChar, adParamInput, 10, "BMAD")
cmd.Parameters.append prm3
Set prm4 = cmd.CreateParameter("KodeFasilitas", adVarChar, adParamInput, 5, Left(cboTarifFasilitas, 1))
cmd.Parameters.append prm4
Set prm5 = cmd.CreateParameter("KodeTarif", adVarChar, adParamInput, 5, Null)
cmd.Parameters.append prm5
Set prm6 = cmd.CreateParameter("NilaiBayar", adDecimal, adParamInput, , NilaiBayar)
prm6.Precision = 38
prm6.NumericScale = 2
cmd.Parameters.append prm6
Set prm7 = cmd.CreateParameter("NilaiFasilitas", adDecimal, adParamInput, , NilaiFasilitas)
prm7.Precision = 38
prm7.NumericScale = 2
cmd.Parameters.append prm7
Set prm8 = cmd.CreateParameter("Tarif", adDecimal, adParamInput, , CDbl(txtBMADs))
prm8.Precision = 38
prm8.NumericScale = 2
cmd.Parameters.append prm8
Set prm9 = cmd.CreateParameter("TarifFasilitas", adDecimal, adParamInput, , 100)
prm9.Precision = 38
prm9.NumericScale = 2
cmd.Parameters.append prm9
Set prm10 = cmd.CreateParameter("NoSeriBahanBaku", adInteger, adParamInput, 5, txtNoSeriBahanBaku)
cmd.Parameters.append prm10
Set prm11 = cmd.CreateParameter("KodeAsal", adInteger, adParamInput, 5, 0)
cmd.Parameters.append prm11

cmd.Execute Y

    If Y = 0 Then
        Set cmd = New ADODB.Command
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 0
        cmd.ActiveConnection = Db
        cmd.CommandText = "sp_BC25DetailBeaMasukTambahanBahanBaku_Ins"
                    
        Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
        cmd.Parameters.append prm1
        Set prm2 = cmd.CreateParameter("NoSeri", adInteger, adParamInput, 5, txtNoSeri)
        cmd.Parameters.append prm2
        Set prm3 = cmd.CreateParameter("JenisTarif", adVarChar, adParamInput, 10, "BMAD")
        cmd.Parameters.append prm3
        Set prm4 = cmd.CreateParameter("KodeFasilitas", adVarChar, adParamInput, 5, Left(cboTarifFasilitas, 1))
        cmd.Parameters.append prm4
        Set prm5 = cmd.CreateParameter("KodeTarif", adVarChar, adParamInput, 5, Null)
        cmd.Parameters.append prm5
        Set prm6 = cmd.CreateParameter("NilaiBayar", adDecimal, adParamInput, , NilaiBayar)
        prm6.Precision = 38
        prm6.NumericScale = 2
        cmd.Parameters.append prm6
        Set prm7 = cmd.CreateParameter("NilaiFasilitas", adDecimal, adParamInput, , NilaiFasilitas)
        prm7.Precision = 38
        prm7.NumericScale = 2
        cmd.Parameters.append prm7
        Set prm8 = cmd.CreateParameter("Tarif", adDecimal, adParamInput, , CDbl(txtBMADs))
        prm8.Precision = 38
        prm8.NumericScale = 2
        cmd.Parameters.append prm8
        Set prm9 = cmd.CreateParameter("TarifFasilitas", adDecimal, adParamInput, , 100)
        prm9.Precision = 38
        prm9.NumericScale = 2
        cmd.Parameters.append prm9
        Set prm10 = cmd.CreateParameter("NoSeriBahanBaku", adInteger, adParamInput, 5, txtNoSeriBahanBaku)
        cmd.Parameters.append prm10
        Set prm11 = cmd.CreateParameter("KodeAsal", adInteger, adParamInput, 5, 0)
        cmd.Parameters.append prm11
        
        cmd.Execute Y
    End If
'###################################################################################################
'###################################################################################################
    
'###################################################################################################
'2. BMI

If Left(cboTarifFasilitas, 1) = "0" Then
    NilaiFasilitas = 0
    NilaiBayar = CDbl(txtBMIsRupiah)
ElseIf Left(cboTarifFasilitas, 1) = "6" Then
    NilaiFasilitas = CDbl(txtBMIsRupiah)
    NilaiBayar = 0
End If

Y = 0
Set cmd = New ADODB.Command
cmd.CommandType = adCmdStoredProc
cmd.CommandTimeout = 0
cmd.ActiveConnection = Db
cmd.CommandText = "sp_BC25DetailBeaMasukTambahanBahanBaku_Upd"
    
Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
cmd.Parameters.append prm1
Set prm2 = cmd.CreateParameter("NoSeri", adInteger, adParamInput, 5, txtNoSeri)
cmd.Parameters.append prm2
Set prm3 = cmd.CreateParameter("JenisTarif", adVarChar, adParamInput, 10, "BMI")
cmd.Parameters.append prm3
Set prm4 = cmd.CreateParameter("KodeFasilitas", adVarChar, adParamInput, 5, Left(cboTarifFasilitas, 1))
cmd.Parameters.append prm4
Set prm5 = cmd.CreateParameter("KodeTarif", adVarChar, adParamInput, 5, Null)
cmd.Parameters.append prm5
Set prm6 = cmd.CreateParameter("NilaiBayar", adDecimal, adParamInput, , NilaiBayar)
prm6.Precision = 38
prm6.NumericScale = 2
cmd.Parameters.append prm6
Set prm7 = cmd.CreateParameter("NilaiFasilitas", adDecimal, adParamInput, , NilaiFasilitas)
prm7.Precision = 38
prm7.NumericScale = 2
cmd.Parameters.append prm7
Set prm8 = cmd.CreateParameter("Tarif", adDecimal, adParamInput, , CDbl(txtBMIs))
prm8.Precision = 38
prm8.NumericScale = 2
cmd.Parameters.append prm8
Set prm9 = cmd.CreateParameter("TarifFasilitas", adDecimal, adParamInput, , 100)
prm9.Precision = 38
prm9.NumericScale = 2
cmd.Parameters.append prm9
Set prm10 = cmd.CreateParameter("NoSeriBahanBaku", adInteger, adParamInput, 5, txtNoSeriBahanBaku)
cmd.Parameters.append prm10
Set prm11 = cmd.CreateParameter("KodeAsal", adInteger, adParamInput, 5, 0)
cmd.Parameters.append prm11

cmd.Execute Y

    If Y = 0 Then
        Set cmd = New ADODB.Command
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 0
        cmd.ActiveConnection = Db
        cmd.CommandText = "sp_BC25DetailBeaMasukTambahanBahanBaku_Ins"
            
        
        Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
        cmd.Parameters.append prm1
        Set prm2 = cmd.CreateParameter("NoSeri", adInteger, adParamInput, 5, txtNoSeri)
        cmd.Parameters.append prm2
        Set prm3 = cmd.CreateParameter("JenisTarif", adVarChar, adParamInput, 10, "BMI")
        cmd.Parameters.append prm3
        Set prm4 = cmd.CreateParameter("KodeFasilitas", adVarChar, adParamInput, 5, Left(cboTarifFasilitas, 1))
        cmd.Parameters.append prm4
        Set prm5 = cmd.CreateParameter("KodeTarif", adVarChar, adParamInput, 5, Null)
        cmd.Parameters.append prm5
        Set prm6 = cmd.CreateParameter("NilaiBayar", adDecimal, adParamInput, , NilaiBayar)
        prm6.Precision = 38
        prm6.NumericScale = 2
        cmd.Parameters.append prm6
        Set prm7 = cmd.CreateParameter("NilaiFasilitas", adDecimal, adParamInput, , NilaiFasilitas)
        prm7.Precision = 38
        prm7.NumericScale = 2
        cmd.Parameters.append prm7
        Set prm8 = cmd.CreateParameter("Tarif", adDecimal, adParamInput, , CDbl(txtBMIs))
        prm8.Precision = 38
        prm8.NumericScale = 2
        cmd.Parameters.append prm8
        Set prm9 = cmd.CreateParameter("TarifFasilitas", adDecimal, adParamInput, , 100)
        prm9.Precision = 38
        prm9.NumericScale = 2
        cmd.Parameters.append prm9
        Set prm10 = cmd.CreateParameter("NoSeriBahanBaku", adInteger, adParamInput, 5, txtNoSeriBahanBaku)
        cmd.Parameters.append prm10
        Set prm11 = cmd.CreateParameter("KodeAsal", adInteger, adParamInput, 5, 0)
        cmd.Parameters.append prm11
        
        cmd.Execute Y
    End If
'###################################################################################################
'###################################################################################################
        
'###################################################################################################
'3. BMTP

If Left(cboTarifFasilitas, 1) = "0" Then
    NilaiFasilitas = 0
    NilaiBayar = CDbl(txtBMTPsRupiah)
ElseIf Left(cboTarifFasilitas, 1) = "6" Then
    NilaiFasilitas = CDbl(txtBMTPsRupiah)
    NilaiBayar = 0
End If

Y = 0
Set cmd = New ADODB.Command
cmd.CommandType = adCmdStoredProc
cmd.CommandTimeout = 0
cmd.ActiveConnection = Db
cmd.CommandText = "sp_BC25DetailBeaMasukTambahanBahanBaku_Upd"
    
Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
cmd.Parameters.append prm1
Set prm2 = cmd.CreateParameter("NoSeri", adInteger, adParamInput, 5, txtNoSeri)
cmd.Parameters.append prm2
Set prm3 = cmd.CreateParameter("JenisTarif", adVarChar, adParamInput, 10, "BMTP")
cmd.Parameters.append prm3
Set prm4 = cmd.CreateParameter("KodeFasilitas", adVarChar, adParamInput, 5, Left(cboTarifFasilitas, 1))
cmd.Parameters.append prm4
Set prm5 = cmd.CreateParameter("KodeTarif", adVarChar, adParamInput, 5, Null)
cmd.Parameters.append prm5
Set prm6 = cmd.CreateParameter("NilaiBayar", adDecimal, adParamInput, , NilaiBayar)
prm6.Precision = 38
prm6.NumericScale = 2
cmd.Parameters.append prm6
Set prm7 = cmd.CreateParameter("NilaiFasilitas", adDecimal, adParamInput, , NilaiFasilitas)
prm7.Precision = 38
prm7.NumericScale = 2
cmd.Parameters.append prm7
Set prm8 = cmd.CreateParameter("Tarif", adDecimal, adParamInput, , CDbl(txtBMTPs))
prm8.Precision = 38
prm8.NumericScale = 2
cmd.Parameters.append prm8
Set prm9 = cmd.CreateParameter("TarifFasilitas", adDecimal, adParamInput, , 100)
prm9.Precision = 38
prm9.NumericScale = 2
cmd.Parameters.append prm9
Set prm10 = cmd.CreateParameter("NoSeriBahanBaku", adInteger, adParamInput, 5, txtNoSeriBahanBaku)
cmd.Parameters.append prm10
Set prm11 = cmd.CreateParameter("KodeAsal", adInteger, adParamInput, 5, 0)
cmd.Parameters.append prm11

cmd.Execute Y

    If Y = 0 Then
        Set cmd = New ADODB.Command
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 0
        cmd.ActiveConnection = Db
        cmd.CommandText = "sp_BC25DetailBeaMasukTambahanBahanBaku_Ins"
            
        
        Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
        cmd.Parameters.append prm1
        Set prm2 = cmd.CreateParameter("NoSeri", adInteger, adParamInput, 5, txtNoSeri)
        cmd.Parameters.append prm2
        Set prm3 = cmd.CreateParameter("JenisTarif", adVarChar, adParamInput, 10, "BMTP")
        cmd.Parameters.append prm3
        Set prm4 = cmd.CreateParameter("KodeFasilitas", adVarChar, adParamInput, 5, Left(cboTarifFasilitas, 1))
        cmd.Parameters.append prm4
        Set prm5 = cmd.CreateParameter("KodeTarif", adVarChar, adParamInput, 5, Null)
        cmd.Parameters.append prm5
        Set prm6 = cmd.CreateParameter("NilaiBayar", adDecimal, adParamInput, , NilaiBayar)
        prm6.Precision = 38
        prm6.NumericScale = 2
        cmd.Parameters.append prm6
        Set prm7 = cmd.CreateParameter("NilaiFasilitas", adDecimal, adParamInput, , NilaiFasilitas)
        prm7.Precision = 38
        prm7.NumericScale = 2
        cmd.Parameters.append prm7
        Set prm8 = cmd.CreateParameter("Tarif", adDecimal, adParamInput, , CDbl(txtBMTPs))
        prm8.Precision = 38
        prm8.NumericScale = 2
        cmd.Parameters.append prm8
        Set prm9 = cmd.CreateParameter("TarifFasilitas", adDecimal, adParamInput, , 100)
        prm9.Precision = 38
        prm9.NumericScale = 2
        cmd.Parameters.append prm9
        Set prm10 = cmd.CreateParameter("NoSeriBahanBaku", adInteger, adParamInput, 5, txtNoSeriBahanBaku)
        cmd.Parameters.append prm10
        Set prm11 = cmd.CreateParameter("KodeAsal", adInteger, adParamInput, 5, 0)
        cmd.Parameters.append prm11

        
        cmd.Execute Y
    End If
'###################################################################################################
'###################################################################################################

'###################################################################################################
'4. BMP

If Left(cboTarifFasilitas, 1) = "0" Then
    NilaiFasilitas = 0
    NilaiBayar = CDbl(txtBMPsRupiah)
ElseIf Left(cboTarifFasilitas, 1) = "6" Then
    NilaiFasilitas = CDbl(txtBMPsRupiah)
    NilaiBayar = 0
End If

Y = 0
Set cmd = New ADODB.Command
cmd.CommandType = adCmdStoredProc
cmd.CommandTimeout = 0
cmd.ActiveConnection = Db
cmd.CommandText = "sp_BC25DetailBeaMasukTambahanBahanBaku_Upd"
    
Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
cmd.Parameters.append prm1
Set prm2 = cmd.CreateParameter("NoSeri", adInteger, adParamInput, 5, txtNoSeri)
cmd.Parameters.append prm2
Set prm3 = cmd.CreateParameter("JenisTarif", adVarChar, adParamInput, 10, "BMP")
cmd.Parameters.append prm3
Set prm4 = cmd.CreateParameter("KodeFasilitas", adVarChar, adParamInput, 5, Left(cboTarifFasilitas, 1))
cmd.Parameters.append prm4
Set prm5 = cmd.CreateParameter("KodeTarif", adVarChar, adParamInput, 5, Null)
cmd.Parameters.append prm5
Set prm6 = cmd.CreateParameter("NilaiBayar", adDecimal, adParamInput, , NilaiBayar)
prm6.Precision = 38
prm6.NumericScale = 2
cmd.Parameters.append prm6
Set prm7 = cmd.CreateParameter("NilaiFasilitas", adDecimal, adParamInput, , NilaiFasilitas)
prm7.Precision = 38
prm7.NumericScale = 2
cmd.Parameters.append prm7
Set prm8 = cmd.CreateParameter("Tarif", adDecimal, adParamInput, , CDbl(txtBMPs))
prm8.Precision = 38
prm8.NumericScale = 2
cmd.Parameters.append prm8
Set prm9 = cmd.CreateParameter("TarifFasilitas", adDecimal, adParamInput, , 100)
prm9.Precision = 38
prm9.NumericScale = 2
cmd.Parameters.append prm9
Set prm10 = cmd.CreateParameter("NoSeriBahanBaku", adInteger, adParamInput, 5, txtNoSeriBahanBaku)
cmd.Parameters.append prm10
Set prm11 = cmd.CreateParameter("KodeAsal", adInteger, adParamInput, 5, 0)
cmd.Parameters.append prm11

cmd.Execute Y

    If Y = 0 Then
        Set cmd = New ADODB.Command
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 0
        cmd.ActiveConnection = Db
        cmd.CommandText = "sp_BC25DetailBeaMasukTambahanBahanBaku_Ins"
            
        
        Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
        cmd.Parameters.append prm1
        Set prm2 = cmd.CreateParameter("NoSeri", adInteger, adParamInput, 5, txtNoSeri)
        cmd.Parameters.append prm2
        Set prm3 = cmd.CreateParameter("JenisTarif", adVarChar, adParamInput, 10, "BMP")
        cmd.Parameters.append prm3
        Set prm4 = cmd.CreateParameter("KodeFasilitas", adVarChar, adParamInput, 5, Left(cboTarifFasilitas, 1))
        cmd.Parameters.append prm4
        Set prm5 = cmd.CreateParameter("KodeTarif", adVarChar, adParamInput, 5, Null)
        cmd.Parameters.append prm5
        Set prm6 = cmd.CreateParameter("NilaiBayar", adDecimal, adParamInput, , NilaiBayar)
        prm6.Precision = 38
        prm6.NumericScale = 2
        cmd.Parameters.append prm6
        Set prm7 = cmd.CreateParameter("NilaiFasilitas", adDecimal, adParamInput, , NilaiFasilitas)
        prm7.Precision = 38
        prm7.NumericScale = 2
        cmd.Parameters.append prm7
        Set prm8 = cmd.CreateParameter("Tarif", adDecimal, adParamInput, , CDbl(txtBMPs))
        prm8.Precision = 38
        prm8.NumericScale = 2
        cmd.Parameters.append prm8
        Set prm9 = cmd.CreateParameter("TarifFasilitas", adDecimal, adParamInput, , 100)
        prm9.Precision = 38
        prm9.NumericScale = 2
        cmd.Parameters.append prm9
        Set prm10 = cmd.CreateParameter("NoSeriBahanBaku", adInteger, adParamInput, 5, txtNoSeriBahanBaku)
        cmd.Parameters.append prm10
        Set prm11 = cmd.CreateParameter("KodeAsal", adInteger, adParamInput, 5, 0)
        cmd.Parameters.append prm11
        
        cmd.Execute Y
    End If
'###################################################################################################
'###################################################################################################


'Dim rssum As New Recordset
'Dim sql As String
'Dim lsNilaiBMT As Double
'
''sql = " Select SUM(NILAI_FASILITAS) AS NILAIBMT From Bea_Cukai_TPB_Barang_Tarif  " & vbCrLf & _
''            " WHERE JENIS_TARIF IN ('BMAD', 'BMI', 'BMTP', 'BMP') " & vbCrLf & _
''            " AND NO_PENGAJUAN = '" & Replace(txtNoPengajuan, "-", "") & "'  " & vbCrLf & _
''            "  "
''
'sql = " Select ISNULL(SUM(NILAI_BAYAR),0) AS NILAIBAYAR,  " & vbCrLf & _
'            "       ISNULL(SUM(NILAI_FASILITAS),0) AS NILAIFASILITAS " & vbCrLf & _
'            " From Bea_Cukai_TPB_Bahan_Baku_Tarif " & vbCrLf & _
'            " WHERE JENIS_TARIF IN ('BMAD', 'BMI', 'BMTP', 'BMP')  " & vbCrLf & _
'            " AND NO_PENGAJUAN = '" & Replace(txtNoPengajuan, "-", "") & "' " & vbCrLf & _
'            "  "
'
'rssum.Open sql, Db, adOpenDynamic, adLockOptimistic
'
'If Not rssum.EOF Then
'    If Left(cboTarifFasilitas, 1) = "0" Then
'        lsNilaiBMT = rssum.fields("NILAIBAYAR")
'    ElseIf Left(cboTarifFasilitas, 1) = "6" Then
'        lsNilaiBMT = rssum.fields("NILAIFASILITAS")
'    End If
'
'End If
'
'Y = 0
'
''INSERT DATA BMT KE DALAM TABEL TPB PUNGUTAN
'Set cmd = New ADODB.Command
'cmd.CommandType = adCmdStoredProc
'cmd.CommandTimeout = 0
'cmd.ActiveConnection = Db
'cmd.CommandText = "sp_BC25DetailPungutan_Upd"
'
'Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
'cmd.Parameters.append prm1
'Set prm2 = cmd.CreateParameter("NoSeri", adInteger, adParamInput, 5, txtNoSeri)
'cmd.Parameters.append prm2
'Set prm3 = cmd.CreateParameter("JenisTarif", adVarChar, adParamInput, 10, "BMT")
'cmd.Parameters.append prm3
'Set prm4 = cmd.CreateParameter("KodeFasilitas", adVarChar, adParamInput, 10, Left(cboTarifFasilitas, 1))
'cmd.Parameters.append prm4
'Set prm5 = cmd.CreateParameter("NilaiPungutan", adDecimal, adParamInput, , lsNilaiBMT)
'prm5.Precision = 38
'prm5.NumericScale = 4
'cmd.Parameters.append prm5
'
'cmd.Execute Y
'
'If Y = 0 Then
'    Set cmd = New ADODB.Command
'    cmd.CommandType = adCmdStoredProc
'    cmd.CommandTimeout = 0
'    cmd.ActiveConnection = Db
'    cmd.CommandText = "sp_BC25DetailPungutan_Ins"
'
'    Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
'    cmd.Parameters.append prm1
'    Set prm2 = cmd.CreateParameter("NoSeri", adInteger, adParamInput, 5, txtNoSeri)
'    cmd.Parameters.append prm2
'    Set prm3 = cmd.CreateParameter("JenisTarif", adVarChar, adParamInput, 10, "BMT")
'    cmd.Parameters.append prm3
'    Set prm4 = cmd.CreateParameter("KodeFasilitas", adVarChar, adParamInput, 10, Left(cboTarifFasilitas, 1))
'    cmd.Parameters.append prm4
'    Set prm5 = cmd.CreateParameter("NilaiPungutan", adDecimal, adParamInput, , lsNilaiBMT)
'    prm5.Precision = 38
'    prm5.NumericScale = 4
'    cmd.Parameters.append prm5
'
'    cmd.Execute
'
'End If

'INSERT DATA BMT KE DALAM TABEL TPB PUNGUTAN


LblerrMsg = DisplayMsg(1101)

End Sub

Private Sub up_SaveData()
Dim RS As ADODB.Recordset
Dim cmd As ADODB.Command
Dim Y As Integer
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim prm3 As ADODB.Parameter
Dim prm4 As ADODB.Parameter
Dim prm5 As ADODB.Parameter
Dim prm6 As ADODB.Parameter
Dim prm7 As ADODB.Parameter
Dim prm8 As ADODB.Parameter
Dim prm9 As ADODB.Parameter
Dim prm10 As ADODB.Parameter
Dim prm11 As ADODB.Parameter
Dim prm12 As ADODB.Parameter

Dim cekFlag As String

'###################################################################################################
'1. BMAD

Set cmd = New ADODB.Command
cmd.CommandType = adCmdStoredProc
cmd.CommandTimeout = 0
cmd.ActiveConnection = Db
cmd.CommandText = "sp_BC25DetailBeaMasukTambahan_Upd"
    
Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
cmd.Parameters.append prm1
Set prm2 = cmd.CreateParameter("NoSeri", adInteger, adParamInput, 5, txtNoSeri)
cmd.Parameters.append prm2
Set prm3 = cmd.CreateParameter("JenisTarif", adVarChar, adParamInput, 10, "BMAD")
cmd.Parameters.append prm3
Set prm4 = cmd.CreateParameter("KodeFasilitas", adVarChar, adParamInput, 5, "2")
cmd.Parameters.append prm4
Set prm5 = cmd.CreateParameter("KodeTarif", adVarChar, adParamInput, 5, Null)
cmd.Parameters.append prm5
Set prm6 = cmd.CreateParameter("NilaiBayar", adDecimal, adParamInput, , 0)
prm6.Precision = 38
prm6.NumericScale = 2
cmd.Parameters.append prm6
Set prm7 = cmd.CreateParameter("NilaiFasilitas", adDecimal, adParamInput, , CDbl(txtBMADsRupiah))
prm7.Precision = 38
prm7.NumericScale = 2
cmd.Parameters.append prm7
Set prm8 = cmd.CreateParameter("Tarif", adDecimal, adParamInput, , CDbl(txtBMADs))
prm8.Precision = 38
prm8.NumericScale = 2
cmd.Parameters.append prm8
Set prm9 = cmd.CreateParameter("TarifFasilitas", adDecimal, adParamInput, , 100)
prm9.Precision = 38
prm9.NumericScale = 2
cmd.Parameters.append prm9
Set prm10 = cmd.CreateParameter("KodeSatuan", adVarChar, adParamInput, 5, Null)
cmd.Parameters.append prm10
Set prm11 = cmd.CreateParameter("JumlahSatuan", adDecimal, adParamInput, , Null)
prm11.Precision = 38
prm11.NumericScale = 4
cmd.Parameters.append prm11

If chkBMADs.Value = Checked Then
    cekFlag = "Y"
Else
    cekFlag = "N"
End If
Set prm12 = cmd.CreateParameter("Flag", adVarChar, adParamInput, 1, cekFlag)
cmd.Parameters.append prm12

cmd.Execute Y

    If Y = 0 Then
        Set cmd = New ADODB.Command
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 0
        cmd.ActiveConnection = Db
        cmd.CommandText = "sp_BC25DetailBeaMasukTambahan_Ins"
                    
        Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
        cmd.Parameters.append prm1
        Set prm2 = cmd.CreateParameter("NoSeri", adInteger, adParamInput, 5, txtNoSeri)
        cmd.Parameters.append prm2
        Set prm3 = cmd.CreateParameter("JenisTarif", adVarChar, adParamInput, 10, "BMAD")
        cmd.Parameters.append prm3
        Set prm4 = cmd.CreateParameter("KodeFasilitas", adVarChar, adParamInput, 5, "2")
        cmd.Parameters.append prm4
        Set prm5 = cmd.CreateParameter("KodeTarif", adVarChar, adParamInput, 5, Null)
        cmd.Parameters.append prm5
        Set prm6 = cmd.CreateParameter("NilaiBayar", adDecimal, adParamInput, , 0)
        prm6.Precision = 38
        prm6.NumericScale = 2
        cmd.Parameters.append prm6
        Set prm7 = cmd.CreateParameter("NilaiFasilitas", adDecimal, adParamInput, , CDbl(txtBMADsRupiah))
        prm7.Precision = 38
        prm7.NumericScale = 2
        cmd.Parameters.append prm7
        Set prm8 = cmd.CreateParameter("Tarif", adDecimal, adParamInput, , CDbl(txtBMADs))
        prm8.Precision = 38
        prm8.NumericScale = 2
        cmd.Parameters.append prm8
        Set prm9 = cmd.CreateParameter("TarifFasilitas", adDecimal, adParamInput, , 100)
        prm9.Precision = 38
        prm9.NumericScale = 2
        cmd.Parameters.append prm9
        Set prm10 = cmd.CreateParameter("KodeSatuan", adVarChar, adParamInput, 5, Null)
        cmd.Parameters.append prm10
        Set prm11 = cmd.CreateParameter("JumlahSatuan", adDecimal, adParamInput, , Null)
        prm11.Precision = 38
        prm11.NumericScale = 4
        cmd.Parameters.append prm11

        If chkBMADs.Value = Checked Then
            cekFlag = "Y"
        Else
            cekFlag = "N"
        End If
        Set prm12 = cmd.CreateParameter("Flag", adVarChar, adParamInput, 1, cekFlag)
        cmd.Parameters.append prm12
        
        cmd.Execute Y
    End If
'###################################################################################################
'###################################################################################################
    
'###################################################################################################
'2. BMI
Y = 0
Set cmd = New ADODB.Command
cmd.CommandType = adCmdStoredProc
cmd.CommandTimeout = 0
cmd.ActiveConnection = Db
cmd.CommandText = "sp_BC25DetailBeaMasukTambahan_Upd"
    
Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
cmd.Parameters.append prm1
Set prm2 = cmd.CreateParameter("NoSeri", adInteger, adParamInput, 5, txtNoSeri)
cmd.Parameters.append prm2
Set prm3 = cmd.CreateParameter("JenisTarif", adVarChar, adParamInput, 10, "BMI")
cmd.Parameters.append prm3
Set prm4 = cmd.CreateParameter("KodeFasilitas", adVarChar, adParamInput, 5, "2")
cmd.Parameters.append prm4
Set prm5 = cmd.CreateParameter("KodeTarif", adVarChar, adParamInput, 5, Null)
cmd.Parameters.append prm5
Set prm6 = cmd.CreateParameter("NilaiBayar", adDecimal, adParamInput, , 0)
prm6.Precision = 38
prm6.NumericScale = 2
cmd.Parameters.append prm6
Set prm7 = cmd.CreateParameter("NilaiFasilitas", adDecimal, adParamInput, , CDbl(txtBMIsRupiah))
prm7.Precision = 38
prm7.NumericScale = 2
cmd.Parameters.append prm7
Set prm8 = cmd.CreateParameter("Tarif", adDecimal, adParamInput, , CDbl(txtBMIs))
prm8.Precision = 38
prm8.NumericScale = 2
cmd.Parameters.append prm8
Set prm9 = cmd.CreateParameter("TarifFasilitas", adDecimal, adParamInput, , 100)
prm9.Precision = 38
prm9.NumericScale = 2
cmd.Parameters.append prm9
Set prm10 = cmd.CreateParameter("KodeSatuan", adVarChar, adParamInput, 5, Null)
cmd.Parameters.append prm10
Set prm11 = cmd.CreateParameter("JumlahSatuan", adDecimal, adParamInput, , Null)
prm11.Precision = 38
prm11.NumericScale = 4
cmd.Parameters.append prm11

If chkBMIS.Value = Checked Then
    cekFlag = "Y"
Else
    cekFlag = "N"
End If
Set prm12 = cmd.CreateParameter("Flag", adVarChar, adParamInput, 1, cekFlag)
cmd.Parameters.append prm12

cmd.Execute Y

    If Y = 0 Then
        Set cmd = New ADODB.Command
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 0
        cmd.ActiveConnection = Db
        cmd.CommandText = "sp_BC25DetailBeaMasukTambahan_Ins"
            
        
        Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
        cmd.Parameters.append prm1
        Set prm2 = cmd.CreateParameter("NoSeri", adInteger, adParamInput, 5, txtNoSeri)
        cmd.Parameters.append prm2
        Set prm3 = cmd.CreateParameter("JenisTarif", adVarChar, adParamInput, 10, "BMI")
        cmd.Parameters.append prm3
        Set prm4 = cmd.CreateParameter("KodeFasilitas", adVarChar, adParamInput, 5, "2")
        cmd.Parameters.append prm4
        Set prm5 = cmd.CreateParameter("KodeTarif", adVarChar, adParamInput, 5, Null)
        cmd.Parameters.append prm5
        Set prm6 = cmd.CreateParameter("NilaiBayar", adDecimal, adParamInput, , 0)
        prm6.Precision = 38
        prm6.NumericScale = 2
        cmd.Parameters.append prm6
        Set prm7 = cmd.CreateParameter("NilaiFasilitas", adDecimal, adParamInput, , CDbl(txtBMIsRupiah))
        prm7.Precision = 38
        prm7.NumericScale = 2
        cmd.Parameters.append prm7
        Set prm8 = cmd.CreateParameter("Tarif", adDecimal, adParamInput, , CDbl(txtBMIs))
        prm8.Precision = 38
        prm8.NumericScale = 2
        cmd.Parameters.append prm8
        Set prm9 = cmd.CreateParameter("TarifFasilitas", adDecimal, adParamInput, , 100)
        prm9.Precision = 38
        prm9.NumericScale = 2
        cmd.Parameters.append prm9
        Set prm10 = cmd.CreateParameter("KodeSatuan", adVarChar, adParamInput, 5, Null)
        cmd.Parameters.append prm10
        Set prm11 = cmd.CreateParameter("JumlahSatuan", adDecimal, adParamInput, , Null)
        prm11.Precision = 38
        prm11.NumericScale = 4
        cmd.Parameters.append prm11
   
        If chkBMIS.Value = Checked Then
            cekFlag = "Y"
        Else
            cekFlag = "N"
        End If
        Set prm12 = cmd.CreateParameter("Flag", adVarChar, adParamInput, 1, cekFlag)
        cmd.Parameters.append prm12
        
        cmd.Execute Y
    End If
'###################################################################################################
'###################################################################################################
        
'###################################################################################################
'3. BMTP
Y = 0
Set cmd = New ADODB.Command
cmd.CommandType = adCmdStoredProc
cmd.CommandTimeout = 0
cmd.ActiveConnection = Db
cmd.CommandText = "sp_BC25DetailBeaMasukTambahan_Upd"
    
Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
cmd.Parameters.append prm1
Set prm2 = cmd.CreateParameter("NoSeri", adInteger, adParamInput, 5, txtNoSeri)
cmd.Parameters.append prm2
Set prm3 = cmd.CreateParameter("JenisTarif", adVarChar, adParamInput, 10, "BMTP")
cmd.Parameters.append prm3
Set prm4 = cmd.CreateParameter("KodeFasilitas", adVarChar, adParamInput, 5, "2")
cmd.Parameters.append prm4
Set prm5 = cmd.CreateParameter("KodeTarif", adVarChar, adParamInput, 5, Null)
cmd.Parameters.append prm5
Set prm6 = cmd.CreateParameter("NilaiBayar", adDecimal, adParamInput, , 0)
prm6.Precision = 38
prm6.NumericScale = 2
cmd.Parameters.append prm6
Set prm7 = cmd.CreateParameter("NilaiFasilitas", adDecimal, adParamInput, , CDbl(txtBMTPsRupiah))
prm7.Precision = 38
prm7.NumericScale = 2
cmd.Parameters.append prm7
Set prm8 = cmd.CreateParameter("Tarif", adDecimal, adParamInput, , CDbl(txtBMTPs))
prm8.Precision = 38
prm8.NumericScale = 2
cmd.Parameters.append prm8
Set prm9 = cmd.CreateParameter("TarifFasilitas", adDecimal, adParamInput, , 100)
prm9.Precision = 38
prm9.NumericScale = 2
cmd.Parameters.append prm9
Set prm10 = cmd.CreateParameter("KodeSatuan", adVarChar, adParamInput, 5, Null)
cmd.Parameters.append prm10
Set prm11 = cmd.CreateParameter("JumlahSatuan", adDecimal, adParamInput, , Null)
prm11.Precision = 38
prm11.NumericScale = 4
cmd.Parameters.append prm11

If chkBMTPs.Value = Checked Then
    cekFlag = "Y"
Else
    cekFlag = "N"
End If
Set prm12 = cmd.CreateParameter("Flag", adVarChar, adParamInput, 1, cekFlag)
cmd.Parameters.append prm12

cmd.Execute Y

    If Y = 0 Then
        Set cmd = New ADODB.Command
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 0
        cmd.ActiveConnection = Db
        cmd.CommandText = "sp_BC25DetailBeaMasukTambahan_Ins"
            
        
        Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
        cmd.Parameters.append prm1
        Set prm2 = cmd.CreateParameter("NoSeri", adInteger, adParamInput, 5, txtNoSeri)
        cmd.Parameters.append prm2
        Set prm3 = cmd.CreateParameter("JenisTarif", adVarChar, adParamInput, 10, "BMTP")
        cmd.Parameters.append prm3
        Set prm4 = cmd.CreateParameter("KodeFasilitas", adVarChar, adParamInput, 5, "2")
        cmd.Parameters.append prm4
        Set prm5 = cmd.CreateParameter("KodeTarif", adVarChar, adParamInput, 5, Null)
        cmd.Parameters.append prm5
        Set prm6 = cmd.CreateParameter("NilaiBayar", adDecimal, adParamInput, , 0)
        prm6.Precision = 38
        prm6.NumericScale = 2
        cmd.Parameters.append prm6
        Set prm7 = cmd.CreateParameter("NilaiFasilitas", adDecimal, adParamInput, , CDbl(txtBMTPsRupiah))
        prm7.Precision = 38
        prm7.NumericScale = 2
        cmd.Parameters.append prm7
        Set prm8 = cmd.CreateParameter("Tarif", adDecimal, adParamInput, , CDbl(txtBMTPs))
        prm8.Precision = 38
        prm8.NumericScale = 2
        cmd.Parameters.append prm8
        Set prm9 = cmd.CreateParameter("TarifFasilitas", adDecimal, adParamInput, , 100)
        prm9.Precision = 38
        prm9.NumericScale = 2
        cmd.Parameters.append prm9
        Set prm10 = cmd.CreateParameter("KodeSatuan", adVarChar, adParamInput, 5, Null)
        cmd.Parameters.append prm10
        Set prm11 = cmd.CreateParameter("JumlahSatuan", adDecimal, adParamInput, , Null)
        prm11.Precision = 38
        prm11.NumericScale = 4
        cmd.Parameters.append prm11

        If chkBMTPs.Value = Checked Then
            cekFlag = "Y"
        Else
            cekFlag = "N"
        End If
        Set prm12 = cmd.CreateParameter("Flag", adVarChar, adParamInput, 1, cekFlag)
        cmd.Parameters.append prm12
        
        cmd.Execute Y
    End If
'###################################################################################################
'###################################################################################################

'###################################################################################################
'4. BMP
Y = 0
Set cmd = New ADODB.Command
cmd.CommandType = adCmdStoredProc
cmd.CommandTimeout = 0
cmd.ActiveConnection = Db
cmd.CommandText = "sp_BC25DetailBeaMasukTambahan_Upd"
    
Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
cmd.Parameters.append prm1
Set prm2 = cmd.CreateParameter("NoSeri", adInteger, adParamInput, 5, txtNoSeri)
cmd.Parameters.append prm2
Set prm3 = cmd.CreateParameter("JenisTarif", adVarChar, adParamInput, 10, "BMP")
cmd.Parameters.append prm3
Set prm4 = cmd.CreateParameter("KodeFasilitas", adVarChar, adParamInput, 5, "2")
cmd.Parameters.append prm4
Set prm5 = cmd.CreateParameter("KodeTarif", adVarChar, adParamInput, 5, Null)
cmd.Parameters.append prm5
Set prm6 = cmd.CreateParameter("NilaiBayar", adDecimal, adParamInput, , 0)
prm6.Precision = 38
prm6.NumericScale = 2
cmd.Parameters.append prm6
Set prm7 = cmd.CreateParameter("NilaiFasilitas", adDecimal, adParamInput, , CDbl(txtBMPsRupiah))
prm7.Precision = 38
prm7.NumericScale = 2
cmd.Parameters.append prm7
Set prm8 = cmd.CreateParameter("Tarif", adDecimal, adParamInput, , CDbl(txtBMPs))
prm8.Precision = 38
prm8.NumericScale = 2
cmd.Parameters.append prm8
Set prm9 = cmd.CreateParameter("TarifFasilitas", adDecimal, adParamInput, , 100)
prm9.Precision = 38
prm9.NumericScale = 2
cmd.Parameters.append prm9
Set prm10 = cmd.CreateParameter("KodeSatuan", adVarChar, adParamInput, 5, Null)
cmd.Parameters.append prm10
Set prm11 = cmd.CreateParameter("JumlahSatuan", adDecimal, adParamInput, , Null)
prm11.Precision = 38
prm11.NumericScale = 4
cmd.Parameters.append prm11

If chkBMPs.Value = Checked Then
    cekFlag = "Y"
Else
    cekFlag = "N"
End If
Set prm12 = cmd.CreateParameter("Flag", adVarChar, adParamInput, 1, cekFlag)
cmd.Parameters.append prm12

cmd.Execute Y

    If Y = 0 Then
        Set cmd = New ADODB.Command
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 0
        cmd.ActiveConnection = Db
        cmd.CommandText = "sp_BC25DetailBeaMasukTambahan_Ins"
            
        
        Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
        cmd.Parameters.append prm1
        Set prm2 = cmd.CreateParameter("NoSeri", adInteger, adParamInput, 5, txtNoSeri)
        cmd.Parameters.append prm2
        Set prm3 = cmd.CreateParameter("JenisTarif", adVarChar, adParamInput, 10, "BMP")
        cmd.Parameters.append prm3
        Set prm4 = cmd.CreateParameter("KodeFasilitas", adVarChar, adParamInput, 5, "2")
        cmd.Parameters.append prm4
        Set prm5 = cmd.CreateParameter("KodeTarif", adVarChar, adParamInput, 5, Null)
        cmd.Parameters.append prm5
        Set prm6 = cmd.CreateParameter("NilaiBayar", adDecimal, adParamInput, , 0)
        prm6.Precision = 38
        prm6.NumericScale = 2
        cmd.Parameters.append prm6
        Set prm7 = cmd.CreateParameter("NilaiFasilitas", adDecimal, adParamInput, , CDbl(txtBMPsRupiah))
        prm7.Precision = 38
        prm7.NumericScale = 2
        cmd.Parameters.append prm7
        Set prm8 = cmd.CreateParameter("Tarif", adDecimal, adParamInput, , CDbl(txtBMPs))
        prm8.Precision = 38
        prm8.NumericScale = 2
        cmd.Parameters.append prm8
        Set prm9 = cmd.CreateParameter("TarifFasilitas", adDecimal, adParamInput, , 100)
        prm9.Precision = 38
        prm9.NumericScale = 2
        cmd.Parameters.append prm9
        Set prm10 = cmd.CreateParameter("KodeSatuan", adVarChar, adParamInput, 5, Null)
        cmd.Parameters.append prm10
        Set prm11 = cmd.CreateParameter("JumlahSatuan", adDecimal, adParamInput, , Null)
        prm11.Precision = 38
        prm11.NumericScale = 4
        cmd.Parameters.append prm11
       
        If chkBMPs.Value = Checked Then
            cekFlag = "Y"
        Else
            cekFlag = "N"
        End If
        Set prm12 = cmd.CreateParameter("Flag", adVarChar, adParamInput, 1, cekFlag)
        cmd.Parameters.append prm12
        
        cmd.Execute Y
    End If
'###################################################################################################
'###################################################################################################


Dim rssum As New Recordset
Dim sql As String
Dim lsNilaiBMT As Double

sql = " Select SUM(NILAI_FASILITAS) AS NILAIBMT From Bea_Cukai_TPB_Barang_Tarif  " & vbCrLf & _
            " WHERE JENIS_TARIF IN ('BMAD', 'BMI', 'BMTP', 'BMP') " & vbCrLf & _
            " AND NO_PENGAJUAN = '" & Replace(txtNoPengajuan, "-", "") & "'  " & vbCrLf & _
            "  "
            
rssum.Open sql, Db, adOpenDynamic, adLockOptimistic

If Not rssum.EOF Then
    lsNilaiBMT = rssum.Fields("NILAIBMT")
End If

Y = 0

'INSERT DATA BMT KE DALAM TABEL TPB PUNGUTAN
Set cmd = New ADODB.Command
cmd.CommandType = adCmdStoredProc
cmd.CommandTimeout = 0
cmd.ActiveConnection = Db
cmd.CommandText = "sp_BC23DetailPungutan_Upd"
    
Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
cmd.Parameters.append prm1
Set prm2 = cmd.CreateParameter("NoSeri", adInteger, adParamInput, 5, txtNoSeri)
cmd.Parameters.append prm2
Set prm3 = cmd.CreateParameter("JenisTarif", adVarChar, adParamInput, 10, "BMT")
cmd.Parameters.append prm3
Set prm4 = cmd.CreateParameter("KodeFasilitas", adVarChar, adParamInput, 10, "2")
cmd.Parameters.append prm4
Set prm5 = cmd.CreateParameter("NilaiPungutan", adDecimal, adParamInput, , lsNilaiBMT)
prm5.Precision = 38
prm5.NumericScale = 4
cmd.Parameters.append prm5

cmd.Execute Y

If Y = 0 Then
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = Db
    cmd.CommandText = "sp_BC23DetailPungutan_Ins"
        
    Set prm1 = cmd.CreateParameter("NoPengajuan", adVarChar, adParamInput, 50, txtNoPengajuan)
    cmd.Parameters.append prm1
    Set prm2 = cmd.CreateParameter("NoSeri", adInteger, adParamInput, 5, txtNoSeri)
    cmd.Parameters.append prm2
    Set prm3 = cmd.CreateParameter("JenisTarif", adVarChar, adParamInput, 10, "BMT")
    cmd.Parameters.append prm3
    Set prm4 = cmd.CreateParameter("KodeFasilitas", adVarChar, adParamInput, 10, "2")
    cmd.Parameters.append prm4
    Set prm5 = cmd.CreateParameter("NilaiPungutan", adDecimal, adParamInput, , lsNilaiBMT)
    prm5.Precision = 38
    prm5.NumericScale = 4
    cmd.Parameters.append prm5
    
    cmd.Execute

End If

'INSERT DATA BMT KE DALAM TABEL TPB PUNGUTAN


LblerrMsg = DisplayMsg(1101)

End Sub

Private Sub cmdBMADs_Click()
If MsgBox("Are you sure want to delete this data?", vbYesNo + vbExclamation, "Delete") = vbYes Then
    If txtNoSeriBahanBaku = "" Then
        up_Delete ("BMAD")
    Else
        up_DeleteBahanBaku ("BMAD")
    End If
    
End If
End Sub

Private Sub cmdBMIs_Click()
If MsgBox("Are you sure want to delete this data?", vbYesNo + vbExclamation, "Delete") = vbYes Then
    If txtNoSeriBahanBaku = "" Then
        up_Delete ("BMI")
    Else
        up_DeleteBahanBaku ("BMI")
    End If
End If
End Sub

Private Sub cmdBMPs_Click()
If MsgBox("Are you sure want to delete this data?", vbYesNo + vbExclamation, "Delete") = vbYes Then
    If txtNoSeriBahanBaku = "" Then
        up_Delete ("BMP")
    Else
        up_DeleteBahanBaku ("BMP")
    End If
End If
End Sub

Private Sub cmdBMTPs_Click()
If MsgBox("Are you sure want to delete this data?", vbYesNo + vbExclamation, "Delete") = vbYes Then
    If txtNoSeriBahanBaku = "" Then
        up_Delete ("BMTP")
    Else
        up_DeleteBahanBaku ("BMTP")
    End If
End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub CmdSubmit_Click()
    If txtNoSeriBahanBaku = "" Then
        up_SaveData
    Else
        up_SaveDataBahanBaku
    End If
    
End Sub

Private Sub Form_Load()
    LblerrMsg.Caption = ""
End Sub

Private Sub txtBMADs_LostFocus()
    If txtBMADs = "" Then txtBMADs = 0
    
    If CDbl(txtBMADs) > 0 Then
        txtBMADsRupiah = Format((CDbl(txtBMADs) / 100) * CDbl(txtCIFRupiah), "#,0.00")
    Else
        txtBMADsRupiah = 0
    End If
    
    txtTotalBM = Format(CDbl(txtBMADsRupiah) + CDbl(txtBMIsRupiah) + CDbl(txtBMTPsRupiah) + CDbl(txtBMPsRupiah), "#,0.00")
    txtBMADs = Format(CDbl(txtBMADs), "#,0.00")
End Sub

Private Sub txtBMIs_LostFocus()
    If txtBMIs = "" Then txtBMIs = 0
    
    If CDbl(txtBMIs) > 0 Then
        txtBMIsRupiah = Format((CDbl(txtBMIs) / 100) * CDbl(txtCIFRupiah), "#,0.00")
    Else
        txtBMIsRupiah = 0
    End If
    
    txtTotalBM = Format(CDbl(txtBMADsRupiah) + CDbl(txtBMIsRupiah) + CDbl(txtBMTPsRupiah) + CDbl(txtBMPsRupiah), "#,0.00")
    txtBMIs = Format(CDbl(txtBMIs), "#,0.00")
End Sub

Private Sub txtBMPs_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtBMPs_LostFocus()
    If txtBMPs = "" Then txtBMPs = 0
    
    If CDbl(txtBMPs) > 0 Then
        txtBMPsRupiah = Format((CDbl(txtBMPs) / 100) * CDbl(txtCIFRupiah), "#,0.00")
    Else
        txtBMPsRupiah = 0
    End If
    
    txtTotalBM = Format(CDbl(txtBMADsRupiah) + CDbl(txtBMIsRupiah) + CDbl(txtBMTPsRupiah) + CDbl(txtBMPsRupiah), "#,0.00")
    
    txtBMPs = Format(CDbl(txtBMPs), "#,0.00")
End Sub

Private Sub txtBMTPs_LostFocus()
    If txtBMTPs = "" Then txtBMTPs = 0
    
    If CDbl(txtBMIs) > 0 Then
        txtBMTPsRupiah = Format((CDbl(txtBMTPs) / 100) * CDbl(txtCIFRupiah), "#,0.00")
    Else
        txtBMTPsRupiah = 0
    End If
    
    txtTotalBM = Format(CDbl(txtBMADsRupiah) + CDbl(txtBMIsRupiah) + CDbl(txtBMTPsRupiah) + CDbl(txtBMPsRupiah), "#,0.00")
    
    txtBMTPs = Format(CDbl(txtBMTPs), "#,0.00")
End Sub

Private Sub txtCIF_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtCIFRupiah_GotFocus()
    txtCIFRupiah = CDbl(txtCIFRupiah)
End Sub

Private Sub txtCIFRupiah_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub




