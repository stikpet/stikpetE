Attribute VB_Name = "test_g_gof"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076



Function ts_g_gof(data As Range, Optional expCount As Range, Optional cc = "none", Optional output = "all")
Attribute ts_g_gof.VB_Description = "Performs a G-test test of goodness-of-fit (a.k.a. likelihood ratio)"
Attribute ts_g_gof.VB_ProcData.VB_Invoke_Func = " \n14"
'Performs a G-test test of goodness-of-fit (a.k.a. likelihood ratio)
'Assumes expected frequency is the same for all categories
'Input: single column specific range
'cc the continuity correction to use, either none (default), yates, pearson, or williams
'output the result to show, either pvalue (default), statistic or df

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
    ts_g_gof = df
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
        
            chiVal = chiVal + cats(i, 2) * Log(cats(i, 2) / expCounts(i))
            
        End If
    Next i
    
    chiVal = chiVal * 2
    
    If cc = "pearson" Then
        chiVal = chiVal * (n - 1) / n
    ElseIf cc = "williams" Then
        chiVal = chiVal / (1 + (k ^ 2 - 1) / (6 * n * (k - 1)))
    End If
    
    If output = "statistic" Then
        ts_g_gof = chiVal
    ElseIf output = "pvalue" Then
        pVal = WorksheetFunction.ChiDist(chiVal, k - 1)
        ts_g_gof = pVal
    Else
        'Which test was used
        testUsed = "G (Likelihood Ratio) test of goodness-of-fit"
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
        
        ts_g_gof = res
    End If
End If

End Function
