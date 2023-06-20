Attribute VB_Name = "eff_size_cohen_w"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076

Public Sub es_cohen_w_addHelp()
Application.MacroOptions _
    Macro:="es_cohen_w", _
    Description:="Cohen's w", _
    category:=14, _
    ArgumentDescriptions:=Array( _
        "the chi-square test statistic", _
        "the sample size")
End Sub

Function es_cohen_w(chi2, n)
'this function calculates Cohens w
'input is the chi-square value and the total sample size
   
    w = Sqr(chi2 / n)
   
es_cohen_w = w

End Function


