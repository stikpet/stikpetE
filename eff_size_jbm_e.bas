Attribute VB_Name = "eff_size_jbm_e"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076

Public Sub es_jbm_e_addHelp()
Application.MacroOptions _
    Macro:="es_jbm_e", _
    Description:="Johnston-Berry-Mielke E", _
    category:=14, _
    ArgumentDescriptions:=Array( _
        "the chi-square test statistic", _
        "the sample size", _
        "the minimum expected count", _
        "test used, either " & Chr(34) & "chi" & Chr(34) & " (default) or " & Chr(34) & "g" & Chr(34))
        
End Sub

Function es_jbm_e(chi2, n, minExp, Optional test = "chi")
'calculates Johnston-Berry-Mielke E
'chiVal -> chi-square value
'minExp -> minimum expected count, if all expected counts equal simply use n/k
'n -> total sample size
'ver -> version of chi-square test, either pearson (default) or g

If test = "chi" Then
    E = chi2 * minExp / (n * (n - minExp))
Else
    E = -1 / WorksheetFunction.Ln(minExp / n) * chi2 / (2 * n)
End If

es_jbm_e = E

End Function
