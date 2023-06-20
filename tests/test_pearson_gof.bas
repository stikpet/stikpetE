Attribute VB_Name = "test_pearson_gof"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076

Public Sub ts_pearson_gof_addHelp()
Application.MacroOptions _
    Macro:="ts_pearson_gof", _
    Description:="Pearson Chi-Square Goodness-of-Fit Test", _
    category:=14, _
    ArgumentDescriptions:=Array( _
        "range with data", _
        "Optional range with categories and expected counts", _
        "use of continuity correction, either " & Chr(34) & "none" & Chr(34) & "(default), " & Chr(34) & "yates" & Chr(34) & ", " & Chr(34) & "pearson" & Chr(34) & ", " & Chr(34) & "williams", _
        "output to show, either " & Chr(34) & "all (default)" & ", " & Chr(34) & "pvalue" & ", " & Chr(34) & "df" & Chr(34) & Chr(34) & ", " & Chr(34) & "statistic" & Chr(34))
               
End Sub

Function ts_pearson_gof(data As Range, Optional expCount As Range, Optional cc, Optional output = "all")
Attribute ts_pearson_gof.VB_Description = "Pearson Chi-Square Goodness-of-Fit Test"
Attribute ts_pearson_gof.VB_ProcData.VB_Invoke_Func = " \n14"
If IsMissing(cc) Then
    cc = "none"
End If

'determine how many categories there are and the total sample size (n).
'cats(i, 1) the label of the category i
'cats(i, 2) the frequency of category i
Dim cats As Variant
nr = data.Rows.Count

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
    ts_pearson_gof = df

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
    If cc = "none" Or cc = "pearson" Or cc = "williams" Then
        For i = 1 To k
            chiVal = chiVal + (cats(i, 2) - expCounts(i)) ^ 2 / expCounts(i)
        Next i
        
        If cc = "pearson" Then
            chiVal = (n - 1) / n * chiVal
        ElseIf cc = "williams" Then
            chiVal = chiVal / (1 + (k ^ 2 - 1) / (6 * n * (k - 1)))
        End If
       
    ElseIf cc = "yates" Then
        For i = 1 To k
            chiVal = chiVal + (Abs(cats(i, 2) - expCounts(i)) - 0.5) ^ 2 / expCounts(i)
        Next i
        
    End If

    If output = "statistic" Then
        ts_pearson_gof = chiVal
    ElseIf output = "pvalue" Then
        pVal = WorksheetFunction.ChiSq_Dist_RT(chiVal, k - 1)
        ts_pearson_gof = pVal
    Else
        'Which test was used
        testUsed = "Pearson chi-square test of goodness-of-fit"
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
        Dim res(1 To 2, 1 To 6)
        res(1, 1) = "statistic"
        res(1, 2) = "df"
        res(1, 3) = "p-value"
        res(1, 4) = "minExp"
        res(1, 5) = "propBelow5"
        res(1, 6) = "test"
        res(2, 1) = chiVal
        res(2, 2) = df
        res(2, 3) = pVal
        res(2, 3) = minExp
        res(2, 3) = propBelow5
        res(2, 6) = testUsed
        
        ts_pearson_gof = res

    End If
    
End If

End Function
