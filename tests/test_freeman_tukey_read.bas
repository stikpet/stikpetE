Attribute VB_Name = "test_freeman_tukey_read"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076


Function ts_freeman_tukey_read(data As Range, Optional expCount As Range, Optional weights As Range, Optional cc = "none", Optional output = "all")
Attribute ts_freeman_tukey_read.VB_Description = "Freeman-Tukey-Read  Test"
Attribute ts_freeman_tukey_read.VB_ProcData.VB_Invoke_Func = " \n14"

Dim weightsArr As Variant
'Set standard weights if none provided
If weights Is Nothing Then
    
    ReDim weightsArr(1 To 2, 1 To 1)
    weightsArr(1, 1) = 4 / 3
    weightsArr(2, 1) = 8 / 3
    nWeights = 2
Else
    nWeights = weights.Rows.Count
    ReDim weightsArr(1 To nWeights, 1 To 1)
    For i = 1 To nWeights
        weightsArr(i, 1) = weights(i, 1)
    Next i
End If
    

'determine how many categories there are and the total sample size (n).
'cats(i, 1) the label of the category i
'cats(i, 2) the frequency of category i
Dim cats As Variant
nr = data.Rows.Count

'The Frequency table
If expCount Is Nothing Then

    ReDim cats(1 To nr, 1 To 2)
    
    k = 0
    n = 0
    For i = 1 To nr
        If data(i, 1) <> "" Then
            n = n + 1
            newCat = True
            If k <> 0 Then
                For j = 1 To k
                    If cats(j, 1) = data(i, 1) Then
                        cats(j, 2) = cats(j, 2) + 1
                        newCat = False
                    End If
                Next j
            End If
            
            If newCat = True Then
                k = k + 1
                cats(k, 1) = data(i, 1)
                cats(k, 2) = 1
            End If
        End If
    Next i
    
    'sum of expected counts equals regular count
    nE = n

Else
    'frequency table if expected counts are given
    k = expCount.Rows.Count
    ReDim cats(1 To k, 1 To 2)
    n = 0
    
    'sum the expected counts just in case they are different
    nE = 0
    For i = 1 To k
        cats(i, 1) = expCount(i, 1)
        cats(i, 2) = WorksheetFunction.CountIf(data, expCount(i, 1))
        n = n + cats(i, 2)
        nE = nE + expCount(i, 2)
    Next i
End If

'the degrees of freedom
df = k - 1

If output = "df" Then
    ts_freeman_tukey_read = df
Else
    'determine expected count
    Dim expCounts As Variant
    ReDim expCounts(1 To k)
    
    For i = 1 To k
        If expCount Is Nothing Then
            'assume for each category equal
            expCounts(i) = n / k
        Else
            For j = 1 To k
                If expCount(i, 1) = cats(j, 1) Then
                    expCounts(i) = expCount(i, 2) / nE * n
                End If
            Next j
        End If
    Next i
    
    chiVal = 0
    For i = 1 To k
        If cats(i, 2) <> 0 Then
            
            If cc = "yates" Then
                If cats(i, 2) > expCounts(i) Then
                    cats(i, 2) = cats(i, 2) - 0.5
                ElseIf cats(i, 2) < expCounts(i) Then
                    cats(i, 2) = cats(i, 2) + 0.5
                End If
            End If
            
            bSum = 0
            For j = 1 To nWeights
                bSum = bSum + weightsArr(j, 1) * Sqr(cats(i, 2) / expCounts(i)) ^ (j - 1)
            Next j
            
            chiVal = chiVal + (Sqr(cats(i, 2)) - Sqr(expCounts(i))) ^ 2 * bSum
            
        End If
    Next i
    
    If cc = "pearson" Then
        chiVal = chiVal * (n - 1) / n
    ElseIf cc = "williams" Then
        chiVal = chiVal / (1 + (k ^ 2 - 1) / (6 * n * (k - 1)))
    End If
    
    If output = "statistic" Then
        ts_freeman_tukey_read = chiVal
    ElseIf output = "pvalue" Then
        pVal = WorksheetFunction.ChiDist(chiVal, k - 1)
        ts_freeman_tukey_read = pVal
    Else
        'Which test was used
        testUsed = "Freeman-Tukey-Read test of goodness-of-fit"
        If cc = "pearson" Then
            testUsed = testUsed + ", with E. Pearson continuity correction"
        ElseIf cc = "williams" Then
            testUsed = testUsed + ", with Williams continuity correction"
        ElseIf cc = "yates" Then
            testUsed = testUsed + ", with Yates continuity correction"
        End If
        
        'Minimum expected counts
        propBelow5 = 0
        minExp = -1
        For i = 1 To k
            If expCounts(i) < minExp Or minExp < 0 Then
                minExp = expCounts(i)
            End If
            
            If expCounts(i) < 5 Then
                propBelow5 = propBelow5 + 1
            End If
        Next i
        
        propBelow5 = propBelow5 / k
        
        'Results
        Dim res(1 To 2, 1 To 8)
        res(1, 1) = "n"
        res(1, 2) = "k"
        res(1, 3) = "statistic"
        res(1, 4) = "df"
        res(1, 5) = "p-value"
        res(1, 6) = "minExp"
        res(1, 7) = "propBelow5"
        res(1, 8) = "test"
        res(2, 1) = n
        res(2, 2) = k
        res(2, 3) = chiVal
        res(2, 4) = df
        res(2, 5) = pVal
        res(2, 6) = minExp
        res(2, 7) = propBelow5
        res(2, 8) = testUsed
        
        ts_freeman_tukey_read = res
    End If
End If

End Function
