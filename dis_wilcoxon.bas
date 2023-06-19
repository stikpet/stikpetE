Attribute VB_Name = "dis_wilcoxon"
Function di_wpmf(k, n)
'Wilcoxon Signed Ranks pmf
    di_wpmf = srf(k, n) / (2 ^ n)
End Function


Function di_wcdf(k, n)
'Wilcoxon Signed Ranks cdf

Dim i As Integer
Dim p As Double


For i = 0 To k
    p = p + di_wpmf(i, n)
Next i

di_wcdf = p

End Function
