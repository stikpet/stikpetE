Attribute VB_Name = "eff_size_cramer_v_gof"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076

Function es_cramer_v_gof(chi2, n, k, Optional bergsma = False)
Attribute es_cramer_v_gof.VB_Description = "Cramer's V for Goodness-of-Fit"
Attribute es_cramer_v_gof.VB_ProcData.VB_Invoke_Func = " \n14"
'this function calculates Cramer's V for a Goodness-of-Fit test
'input is the chi-square value, the total sample size and the number of categories.
'optional the use of a Bergsma correction, and the output to show (either "value", or "qual"
   
    df = k - 1
    
    If bergsma Then
        kAvg = k - (k - 1) ^ 2 / (n - 1)
        phi2 = chi2 / n
        phi2Avg = WorksheetFunction.Max(0, phi2 - (k - 1) / (n - 1))
        v = Sqr(phi2Avg / (kAvg - 1))
    Else
        v = Sqr(chi2 / (n * df))
    End If
   
es_cramer_v_gof = v
    
End Function

