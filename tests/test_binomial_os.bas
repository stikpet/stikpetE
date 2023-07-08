Attribute VB_Name = "test_binomial_os"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076

    
Function ts_binomial_os(data As Range, Optional codes As Range, _
                        Optional p0 = 0.5, _
                        Optional TwoSidedMethod = "eqdist", _
                        Optional output = "all")
Attribute ts_binomial_os.VB_Description = "one-sample binomial test"
Attribute ts_binomial_os.VB_ProcData.VB_Invoke_Func = " \n14"

'one-sample exact binomial test
'data list of data
'codes list of the two codes of the two categories to compare
'p0 the expected proportion of the first category
'twoSidedMethod to indicate the method to use for two sided in exact test only, either "double", "smallp", or "eqdist"

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

'one sided test
sig1 = WorksheetFunction.BinomDist(minCount, n, ExpProp, True)

'two sided tests
If TwoSidedMethod = "double" Then
    'double one-sided

    sigR = sig1
    testUsed = "exact binomial, double one-tail"

ElseIf TwoSidedMethod = "eqdist" Then
    'Equal distance
    expCount = n * ExpProp
    Dist = expCount - minCount
    RightCount = expCount + Dist
    sigR = 1 - WorksheetFunction.BinomDist(RightCount - 1, n, ExpProp, True)
    testUsed = "exact binomial, equal distance"
    
Else
    'Method of small p
    binSmall = WorksheetFunction.BinomDist(minCount, n, ExpProp, False)
    sigR = 0
    For i = minCount + 1 To n
        binDist = WorksheetFunction.BinomDist(i, n, ExpProp, False)
        If binDist <= binSmall Then
            sigR = sigR + binDist
        End If
    Next i
    testUsed = "exact binomial, method of small p"
End If

sig2 = sig1 + sigR
    
If output = "all" Then
    'Results
    Dim res(1 To 2, 1 To 2)
    res(1, 1) = "p-value"
    res(1, 2) = "test"
    res(2, 1) = sig2
    res(2, 2) = testUsed
    ts_binomial_os = res

Else
    ts_binomial_os = sig2
End If

End Function

