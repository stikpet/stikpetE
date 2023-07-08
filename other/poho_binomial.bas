Attribute VB_Name = "poho_binomial"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076

Function ph_binomial(data As Range, Optional expCount As Range, _
                        Optional TwoSidedMethod = "eqdist", _
                        Optional posthoc = "bonferroni")
Attribute ph_binomial.VB_Description = "Pairwise Binomial Test for Post-Hoc Analysis"
Attribute ph_binomial.VB_ProcData.VB_Invoke_Func = " \n14"

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
    
    'loop over all possible pairs
    pairNr = 0
    'Results
    nPairs = WorksheetFunction.Combin(k, 2)
    Dim res As Variant
    ReDim res(1 To (nPairs + 1), 1 To 8)
    res(1, 1) = "category 1"
    res(1, 2) = "category 2"
    res(1, 3) = "n1"
    res(1, 4) = "n2"
    res(1, 5) = "obs. prop. cat. 1"
    res(1, 6) = "exp. prop. cat. 1"
    res(1, 7) = "p-value"
    res(1, 8) = "adj. p-value"
            
    For i = 1 To k - 1
        For j = i + 1 To k
            pairNr = pairNr + 1
    
            n1 = cats(i, 2)
            n2 = cats(j, 2)
            n = n1 + n2
            
            minCount = n1
            ExpProp = expCounts(i) / (expCounts(i) + expCounts(j))
            
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
                expC = n * ExpProp
                Dist = expC - minCount
                RightCount = expC + Dist
                sigR = 1 - WorksheetFunction.BinomDist(RightCount - 1, n, ExpProp, True)
                testUsed = "exact binomial, equal distance"
                
            Else
                'Method of small p
                binSmall = WorksheetFunction.BinomDist(minCount, n, ExpProp, False)
                sigR = 0
                For m = minCount + 1 To n
                    binDist = WorksheetFunction.BinomDist(m, n, ExpProp, False)
                    If binDist <= binSmall Then
                        sigR = sigR + binDist
                    End If
                Next m
                testUsed = "exact binomial, method of small p"
            End If
            
            sig2 = sig1 + sigR
            If sig2 > 1 Then
                sig2 = 1
            End If
            
            If posthoc = "bonferroni" Then

                sigAdj = sig2 * nPairs
                If sigAdj > 1 Then
                    sigAdj = 1
                End If
            End If

            res(pairNr + 1, 1) = cats(i, 1)
            res(pairNr + 1, 2) = cats(j, 1)
            res(pairNr + 1, 3) = cats(i, 2)
            res(pairNr + 1, 4) = cats(j, 2)
            res(pairNr + 1, 5) = n1 / n
            res(pairNr + 1, 6) = ExpProp
            res(pairNr + 1, 7) = sig2
            res(pairNr + 1, 8) = sigAdj
            
        Next j
    Next i
    

    ph_binomial = res

End Function

