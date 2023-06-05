Attribute VB_Name = "test_score_os"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076

Public Sub ts_score_addHelp()
Application.MacroOptions _
    Macro:="ts_score_os", _
    Description:="one-sample Score test", _
    category:=14, _
    ArgumentDescriptions:=Array( _
        "range with data", _
        "range with the two codes to count", _
        "expected proportion (default 0.5)", _
        "use of continuity correction, either " & Chr(34) & "none" & Chr(34) & "(default), or " & Chr(34) & "yates" & Chr(34), _
        "output to show, either " & Chr(34) & "all (default)" & Chr(34) & ", " & Chr(34) & "pvalue" & Chr(34) & ", " & Chr(34) & "statistic" & Chr(34))
        
End Sub

Function ts_score_os(data As Range, Optional codes As Range, _
                    Optional p0 = 0.5, _
                    Optional cc = "none", _
                    Optional output = "all")
Attribute ts_score_os.VB_Description = "one-sample score test"
Attribute ts_score_os.VB_ProcData.VB_Invoke_Func = " \n14"

'one-sample score test
'an approximation using the normal distribution for a one-sample binomial test

If codes Is Nothing Then

    k = 0
    nt = data.Rows.Count
    
    k1 = data.Cells(1, 1)
    i = 2
    If k1 = "" Then
        Do While k1 = ""
            k1 = data.Cells(i, 1)
            i = i + 1
        Loop
    End If
    
    k2 = data.Cells(i, 1)
    If k2 = "" Or k2 = k1 Then
        i = i + 1
        Do While k2 = "" Or k2 = k1
            k2 = data.Cells(i, 1)
            i = i + 1
        Loop
    End If

Else
    k1 = codes.Cells(1, 1)
    k2 = codes.Cells(2, 1)
End If

n1 = WorksheetFunction.CountIf(data, k1)
n2 = WorksheetFunction.CountIf(data, k2)
n = n1 + n2

minCount = n1
ExpProp = p0
If n2 < n1 Then
    minCount = n2
    ExpProp = 1 - ExpProp
End If

'Normal approximation
If cc = "none" Then
    p = minCount / n
    q = 1 - p
    se = (p0 * (1 - p0) / n) ^ 0.5
    Z = (p - ExpProp) / se
    sig2 = 2 * (1 - WorksheetFunction.NormSDist(Abs(Z)))
    testValue = Z
    testUsed = "Normal approximation"
    testStatistic = "Z"
    ccUsed = "none"
    
ElseIf cc = "yates" Then
'Normal approximation with continuity correction
    p = (minCount + 0.5) / n
    q = 1 - p
    se = (p0 * (1 - p0) / n) ^ 0.5
    Z = (p - ExpProp) / se
    sig2 = 2 * (1 - WorksheetFunction.NormSDist(Abs(Z)))
    testValue = Z
    testUsed = "Normal approximation with Yates continuity correction"
    testStatistic = "Z-adjusted"
    ccUsed = "standard"
End If

If output = "all" Then
    'Results
    Dim res(1 To 2, 1 To 3)
    res(1, 1) = "statistic"
    res(1, 2) = "p-value"
    res(1, 3) = "test"
    res(2, 1) = testValue
    res(2, 2) = sig2
    res(2, 3) = testUsed
    
    ts_score_os = res

Else

    If output = "statistic" Then
        ts_score_os = testValue
    Else
        ts_score_os = sig2
    End If
    
End If



End Function



