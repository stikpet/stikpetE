Attribute VB_Name = "eff_size_cramer_v_gof"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076

Public Sub es_cramer_v_gof_addHelp()
Application.MacroOptions _
    Macro:="es_cramer_v_gof", _
    Description:="Cramer's V for Goodness-of-Fit", _
    category:=14, _
    ArgumentDescriptions:=Array( _
        "the chi-square test statistic", _
        "the sample size", _
        "the number of categories", _
        "optional boolean to indicate the use of the Bergsma correction (default is False)")

End Sub

Function es_cramer_v_gof(chi2, n, k, Optional bergsma = False)
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

