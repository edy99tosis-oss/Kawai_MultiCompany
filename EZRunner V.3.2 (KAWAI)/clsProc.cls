VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub AlignHeader(nmGrid As VSFlexGrid, Optional brs As Integer)
Dim kolAwal As Integer
'#Add by Herfin 20050511
With nmGrid
    .Cell(flexcpAlignment, 0, kolAwal, 0, .ColS - 1) = flexAlignCenterCenter
    If brs = 1 Then .Cell(flexcpAlignment, 1, kolAwal, 1, .ColS - 1) = flexAlignCenterCenter
End With
End Sub

Function nilConvertUnit(qty As Double, unit As String, baseUnit As String) As Double
    
    nilConvertUnit = qty
    
'    If unit = "pcs" Or unit = "01" Then
'        If baseUnit = "pcs" Or baseUnit = "01" Then nilConvertUnit = qty
'    ElseIf unit = "Cyl" Or unit = "10" Then
'        If baseUnit = "Cyl" Or baseUnit = "10" Then nilConvertUnit = qty
'    ElseIf unit = "Rol" Or unit = "11" Then
'        If baseUnit = "Rol" Or baseUnit = "11" Then nilConvertUnit = qty
'
'    ElseIf unit = "t" Or unit = "04" Or unit = "Kg" Or unit = "03" Or unit = "g" Or unit = "02" Then
'        If unit = "t" Or unit = "04" Then
'            If baseUnit = "t" Or baseUnit = "04" Then
'                nilConvertUnit = qty
'            ElseIf baseUnit = "Kg" Or baseUnit = "03" Then
'                nilConvertUnit = qty * 1000
'            ElseIf baseUnit = "g" Or baseUnit = "02" Then
'                nilConvertUnit = qty * 1000000
'            End If
'        ElseIf unit = "Kg" Or unit = "03" Then
'            If baseUnit = "t" Or baseUnit = "04" Then
'                nilConvertUnit = qty / 1000
'            ElseIf baseUnit = "Kg" Or baseUnit = "03" Then
'                nilConvertUnit = qty
'            ElseIf baseUnit = "g" Or baseUnit = "02" Then
'                nilConvertUnit = qty * 1000
'            End If
'        ElseIf unit = "g" Or unit = "02" Then
'            If baseUnit = "t" Or baseUnit = "04" Then
'                nilConvertUnit = qty / 10000000
'            ElseIf baseUnit = "Kg" Or baseUnit = "03" Then
'                nilConvertUnit = qty / 1000
'            ElseIf baseUnit = "g" Or baseUnit = "02" Then
'                nilConvertUnit = qty
'            End If
'        End If
'
'    ElseIf unit = "m" Or unit = "06" Or unit = "cm" Or unit = "05" Then
'        If unit = "m" Or unit = "06" Then
'            If baseUnit = "m" Or baseUnit = "06" Then
'                nilConvertUnit = qty
'            ElseIf baseUnit = "cm" Or baseUnit = "05" Then
'                nilConvertUnit = qty * 100
'            End If
'        ElseIf unit = "cm" Or unit = "05" Then
'            If baseUnit = "m" Or baseUnit = "06" Then
'                nilConvertUnit = qty / 100
'            ElseIf baseUnit = "cm" Or baseUnit = "05" Then
'                nilConvertUnit = qty
'            End If
'        End If
'
'    ElseIf unit = "Kl" Or unit = "09" Or unit = "Ltr" Or unit = "08" Or unit = "cc" Or unit = "07" Then
'        If unit = "Kl" Or unit = "09" Then
'            If baseUnit = "Kl" Or baseUnit = "09" Then
'                nilConvertUnit = qty
'            ElseIf baseUnit = "Ltr" Or baseUnit = "08" Then
'                nilConvertUnit = qty * 1000
'            ElseIf baseUnit = "cc" Or baseUnit = "07" Then
'                nilConvertUnit = qty * 1000000
'            End If
'        ElseIf unit = "Ltr" Or unit = "08" Then
'            If baseUnit = "Kl" Or baseUnit = "09" Then
'                nilConvertUnit = qty / 1000
'            ElseIf baseUnit = "Ltr" Or baseUnit = "08" Then
'                nilConvertUnit = qty
'            ElseIf baseUnit = "cc" Or baseUnit = "07" Then
'                nilConvertUnit = qty * 1000
'            End If
'        ElseIf unit = "cc" Or unit = "07" Then
'            If baseUnit = "Kl" Or baseUnit = "09" Then
'                nilConvertUnit = qty / 1000000
'            ElseIf baseUnit = "Ltr" Or baseUnit = "08" Then
'                nilConvertUnit = qty / 1000
'            ElseIf baseUnit = "cc" Or baseUnit = "07" Then
'                nilConvertUnit = qty
'            End If
'        End If
'    End If
End Function

Function formatnilai(nilai As Double, Optional curr As String, Optional jmlDec As Integer) As String
    If nilai = 0 Then
        formatnilai = 0
    Else
        If curr = "" Then
            If Round(CDec(nilai)) / CDec(nilai) = 1 Then
                formatnilai = Format(nilai, "#,##0")
            Else
                If jmlDec = 0 Then
                    formatnilai = Format(nilai, "#,##0.#0")
                ElseIf jmlDec = 3 Then
                    formatnilai = Format(nilai, "#,##0.##0")
                End If
            End If
        Else
            If curr = "IDR" Then formatnilai = Format(nilai, "#,##0") Else formatnilai = Format(nilai, "#,##0.#0")
        End If
    End If
End Function
