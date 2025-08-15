Attribute VB_Name = "MdlSubModule"
Public HakU As Byte, i As Long
Public ClsProc As New ClsProc

Public Const ColGraph = "&H80000005,&H00FFC0FF,&H00FFC0C0"
Public StPrint As Boolean

Public Sub isiCbo(nmcbo, tableNm As String, _
    fieldNm1 As String, fieldNm2 As String, colWidth1 As Integer, colwidth2 As Integer, _
    sortByField As String, _
    Optional fieldNm3 As String, Optional colWidth3 As Integer, _
    Optional kondisi As String, Optional notAll As Byte, _
    Optional fieldNm4 As String, Optional colWidth4 As Integer, _
    Optional TxtCol As Integer)

Dim rscbo As New ADODB.Recordset
With nmcbo
    .columnCount = 4
    .TextColumn = IIf(TxtCol = 0, 1, TxtCol)
        
    sql = "Select distinct " & fieldNm1 & " as a," & fieldNm2 & " as b "
    
    If fieldNm3 <> "" Then sql = sql & ", " & fieldNm3 & " As c "
    If fieldNm4 <> "" Then sql = sql & ", " & fieldNm4 & " As d "
    sql = sql & "from " & tableNm
    If kondisi <> "" Then sql = sql & " Where " & kondisi
    
    sql = sql & " order by " & sortByField
    
    Set rscbo = Db.Execute(sql)
    
    If notAll = 0 Then
        .clear: i = 0
    ElseIf notAll = 1 Then
        .AddItem ""
        .List(0, 0) = "All"
        .List(0, 1) = "All"
        i = 1
    Else
        .AddItem ""
        .List(0, 0) = ""
        .List(0, 1) = ""
        i = 1
    End If
    
    Do While Not rscbo.EOF
        DoEvents
        .AddItem ""
        .List(i, 0) = Trim(rscbo!a)
        .List(i, 1) = Trim(rscbo!b)
        If fieldNm3 <> "" Then .List(i, 2) = Trim(rscbo!C)
        If fieldNm4 <> "" Then .List(i, 3) = Trim(rscbo!d)
        i = i + 1
        rscbo.MoveNext
    Loop
    Set rscbo = Nothing
    
    .ListRows = 20
    .ListWidth = colWidth1 + colwidth2 + colWidth3
    .ColumnWidths = colWidth1 & "pt;" & colwidth2 & "pt;" & colWidth3 & "pt;" & colWidth4
    nmcbo = ""
End With
End Sub

Public Function DeleteFile(file) As String
Dim fso

DeleteFile = ""
On Error GoTo HandleErr

    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(file) Then fso.DeleteFile file, True

HandleErr:
    If err.Description = "Permission denied" Then
        DeleteFile = "The file is still opened"
    ElseIf err.Description <> "" Then
        DeleteFile = err.Description
    End If
End Function

Public Function setRange(ByVal maxVal As Double, Optional EffPercent As Boolean) As Double
Dim pangkat As Integer
    
    maxVal = Abs(maxVal)
    If EffPercent = False Then
        pangkat = 0: setRange = 0
        
toSetRange:
        If maxVal >= 0 And maxVal <= 10 Then
            setRange = 1
        ElseIf maxVal > 10 And maxVal <= 100 Then
            setRange = 10
        ElseIf maxVal > 100 And maxVal <= 500 Then
            setRange = 50
        ElseIf maxVal > 500 And maxVal <= 1000 Then
            setRange = 100
        End If
        setRange = setRange * (10 ^ pangkat)
        
        If setRange = 0 Then
            pangkat = pangkat + 1
            maxVal = maxVal / (10)
            GoTo toSetRange
        End If
    Else
        If maxVal > 0 And maxVal <= 1 Then
            setRange = 1
        ElseIf maxVal > 1 And maxVal <= 5 Then
            setRange = 5
        ElseIf maxVal > 5 And maxVal <= 10 Then
            setRange = 10
        ElseIf maxVal > 10 And maxVal <= 100 Then
            setRange = 100
        End If
    End If
End Function

Public Function RoundUp(NiL As Double) As Double
    NiL = Abs(NiL): RoundUp = NiL
    If Round(NiL) / NiL <> 1 Then RoundUp = Fix(NiL) + 1
End Function
