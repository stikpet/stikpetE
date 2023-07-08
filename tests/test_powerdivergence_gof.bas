Attribute VB_Name = "test_powerdivergence_gof"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076


Function ts_powerdivergence_gof(data As Range, _
        Optional expCount As Range, _
        Optional lambd = 2 / 3, _
        Optional corr, _
        Optional output = "all")
Attribute ts_powerdivergence_gof.VB_Description = "Power Divergence tests"
Attribute ts_powerdivergence_gof.VB_ProcData.VB_Invoke_Func = " \n14"

    If IsMissing(corr) Then
        corr = "none"
    End If

    'Set correction factor to 1 (no correction)
    corFactor = 1
    
    'Test Used
    If lambd = 2 / 3 Or lambd = "cressie-read" Then
        lambd = 2 / 3
        testUsed = "Cressie-Read"
    
    ElseIf lambd = 0 Or lambd = "likelihood-ratio" Then
        lambd = 0
        testUsed = "likelihood ratio"
        
    ElseIf lambd = -1 Or lambd = "mod-log" Then
        lambd = -1
        testUsed = "mod-log likelihood ratio"
        
    ElseIf lambd = 1 Or lambd = "pearson" Then
        lambd = 1
        testUsed = "Pearson chi-square"
    
    ElseIf lambd = -0.5 Or lambd = "freeman-tukey" Then
        lambd = -0.5
        testUsed = "Freeman-Tukey"
        
    ElseIf lambd = -2 Or lambd = "neyman" Then
        lambd = -2
        testUsed = "Neyman"
    Else
        testUsed = "power divergence with lambda = " + Str(lambd)
    End If
    
    
    'THE TESTS
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
        ts_powerdivergence_gof = df
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
                
                If corr = "yates" Then
                    If cats(i, 2) > expCounts(i) Then
                        cats(i, 2) = cats(i, 2) - 0.5
                    ElseIf cats(i, 2) < expCounts(i) Then
                        cats(i, 2) = cats(i, 2) + 0.5
                    End If
                End If
                
                If lambd = 0 Then
                    chiVal = chiVal + cats(i, 2) * Log(cats(i, 2) / expCounts(i))
                ElseIf lambd = -1 Then
                    chiVal = chiVal + expCounts(i) * Log(expCounts(i) / cats(i, 2))
                Else
                    chiVal = chiVal + cats(i, 2) * ((cats(i, 2) / expCounts(i)) ^ lambd - 1)
                End If
                
            End If
        Next i
        
        If lambd = 0 Or lambd = -1 Then
            chiVal = chiVal * 2
        Else
            chiVal = chiVal * 2 / (lambd * (lambd + 1))
        End If
        
        chiVal = chiVal
        
        If corr = "pearson" Then
            chiVal = chiVal * (n - 1) / n
        ElseIf corr = "williams" Then
            chiVal = chiVal / (1 + (k ^ 2 - 1) / (6 * n * (k - 1)))
        End If
        
        If output = "statistic" Then
            ts_powerdivergence_gof = chiVal
        ElseIf output = "pvalue" Then
            pVal = WorksheetFunction.ChiDist(chiVal, k - 1)
            ts_powerdivergence_gof = pVal
        Else
            'Which test was used
            If corr = "pearson" Then
                testUsed = testUsed + ", with E. Pearson continuity correction"
            ElseIf corr = "williams" Then
                testUsed = testUsed + ", with Williams continuity correction"
            ElseIf corr = "yates" Then
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
            
            ts_powerdivergence_gof = res
        End If
    End If

End Function
