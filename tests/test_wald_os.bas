Attribute VB_Name = "test_wald_os"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076


Function ts_wald_os(data As Range, Optional codes As Range, _
                    Optional p0 As Double = 0.5, _
                    Optional cc As String = "none", _
                    Optional output As String = "all")
Attribute ts_wald_os.VB_Description = "one-sample Wald test"
Attribute ts_wald_os.VB_ProcData.VB_Invoke_Func = " \n14"

'one-sample Wald test
'approximation of one-sample binomial using normal distribution

Dim n1, n2, n, minCount As Integer
Dim ExpProp, sig2 As Double

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

'Wald approximation
If cc = "none" Then
    p = minCount / n
    q = 1 - p
    se = (p * q / n) ^ 0.5
    Z = (p - ExpProp) / se
    sig2 = 2 * (1 - WorksheetFunction.NormSDist(Abs(Z)))
    testValue = Z
    testUsed = "Normal approximation"
    testStatistic = "Z"
    ccUsed = "none"

ElseIf cc = "yates" Then
'Wald approximation with continuity correction
    p = (minCount + 0.5) / n
    q = 1 - p
    se = (p * q / n) ^ 0.5
    Z = (p - ExpProp) / se
    sig2 = 2 * (1 - WorksheetFunction.NormSDist(Abs(Z)))
    testValue = Z
    testUsed = "Normal approximation with Yates continuity correction"
    testStatistic = "Z-adjusted"
    ccUsed = "standard"

End If

If output = "all" Then
    'Results
    Dim res(1 To 2, 1 To 4)
    res(1, 1) = "n"
    res(1, 2) = "statistic"
    res(1, 3) = "p-value"
    res(1, 4) = "test"
    res(2, 1) = n
    res(2, 2) = testValue
    res(2, 3) = sig2
    res(2, 4) = testUsed
    
    ts_wald_os = res

Else

    If output = "statistic" Then
        ts_wald_os = testValue
    Else
        ts_wald_os = sig2
    End If

End If


End Function
