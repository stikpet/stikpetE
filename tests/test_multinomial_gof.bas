Attribute VB_Name = "test_multinomial_gof"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076

Function ts_multinomial_gof(data As Range, Optional expCount As Range, Optional output = "all")
Attribute ts_multinomial_gof.VB_Description = "Performs a G-test test of goodness-of-fit (a.k.a. likelihood ratio)"
Attribute ts_multinomial_gof.VB_ProcData.VB_Invoke_Func = " \n14"
'Computes the p-value for an exact multinomial goodness-of-fit test

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

'the total number of all possible combinations
nCombs = (n + 1) ^ k

'a variable to store the possible combinations
Dim combs() As Integer
Dim freqs() As Integer
'creating the possible combinations with only 1 category
ReDim combs(1 To nCombs, 1 To k + 1)
ReDim freqs(1 To k)

For i = 1 To k
    freqs(i) = cats(i, 2)
    For j = 1 To nCombs
        combs(j, i) = WorksheetFunction.RoundDown((j - 1) / ((n + 1) ^ (k - i)), 0) Mod (n + 1)
    Next j
Next i

'add the sum of the rows, set to 0 if it is not equal to n
For i = 1 To nCombs
    rowSum = 0
    For j = 1 To k
        rowSum = rowSum + combs(i, j)
    Next j
    
    If rowSum = n Then
        combs(i, k + 1) = rowSum
    Else
        combs(i, k + 1) = 0
    End If
    
Next i

'The probability of the observed frequencies
pObs = 1
For i = 1 To k
    pObs = pObs * (expCounts(i) / n) ^ cats(i, 2)
Next i
pObs = pObs * WorksheetFunction.MultiNomial(freqs)

If output = "pobs" Then
    ts_multinomial_gof = pObs
Else

    'The significance
    pVal = 0
    eqCombs = 0
    For i = 1 To nCombs
        If combs(i, k + 1) = n Then
            eqCombs = eqCombs + 1
        
            denom = 1
            pComb = 1
            For j = 1 To k
                pComb = pComb * (expCounts(j) / n) ^ combs(i, j)
                denom = denom * WorksheetFunction.Fact(combs(i, j))
            Next j
            pComb = pComb * WorksheetFunction.Fact(n) / denom
            
            If pComb <= pObs Then
                pVal = pVal + pComb
            End If
            
        End If
    Next i
    
    If output = "ncomb" Then
        ts_multinomial_gof = eqCombs
    ElseIf output = "pvalue" Then
        ts_multinomial_gof = pVal
    Else
        'Results
        Dim res(1 To 2, 1 To 4)
        res(1, 1) = "p-obs"
        res(1, 2) = "n comb."
        res(1, 3) = "p-value"
        res(1, 4) = "test"
        res(2, 1) = pObs
        res(2, 2) = eqCombs
        res(2, 3) = pVal
        res(2, 4) = "one-sample multinomial exact goodness-of-fit test"
        
        ts_multinomial_gof = res
    
    End If

End If

End Function

