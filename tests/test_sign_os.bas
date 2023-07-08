Attribute VB_Name = "test_sign_os"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076

Function ts_sign_os(data As Range, _
                    Optional levels As Range, _
                    Optional mu = "none", _
                    Optional output = "all")
Attribute ts_sign_os.VB_Description = "perform a one-sample sign test"
Attribute ts_sign_os.VB_ProcData.VB_Invoke_Func = " \n14"
                    
    'get data as numeric values
    If levels Is Nothing Then
        dataN = he_range_to_num_array(data)
    Else
        dataN = he_replace(data, levels)
    End If
    
    'sort the numeric values
    dataN = he_sort(dataN)
    
    'data -> scores as vector
    'mu -> hypothesized median, default is use of midrange
    If mu = "none" Then
        mu = (WorksheetFunction.Min(dataN) + WorksheetFunction.Max(dataN)) / 2
    End If
    
    n = UBound(dataN, 1) - LBound(dataN, 1) + 1
    
    'count cases below hypothesized median and number of cases unequal to it.
    nbelow = 0
    For i = 0 To n - 1
        If dataN(i) < mu Then
            nbelow = nbelow + 1
        End If
    Next i
    
    'determine expected number of cases in each group
    nExp = n * 0.5
    
    'determine the probability
    If nbelow < nExp Then
        oneTail = WorksheetFunction.BinomDist(nbelow, n, 0.5, True)
    Else
        oneTail = 1 - WorksheetFunction.BinomDist(nbelow - 1, n, 0.5, True)
    End If
    
    pVal = 2 * oneTail
    
    testUsed = "one-sample sign test"

    'Results
    If output = "pvalue" Then
        ts_sign_os = pVal
    ElseIf output = "mu" Then
        ts_sign_os = mu
    Else
        Dim res(1 To 2, 1 To 3)
        res(1, 1) = "mu"
        res(1, 2) = "p-value"
        res(1, 3) = "test"
        res(2, 1) = mu
        res(2, 2) = pVal
        res(2, 3) = testUsed
        
        ts_sign_os = res
    End If

End Function
