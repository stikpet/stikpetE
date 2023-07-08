Attribute VB_Name = "eff_size_cohen_w"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076


Function es_cohen_w(chi2, n)
Attribute es_cohen_w.VB_Description = "Cohen's w"
Attribute es_cohen_w.VB_ProcData.VB_Invoke_Func = " \n14"
'this function calculates Cohens w
'input is the chi-square value and the total sample size
   
    w = Sqr(chi2 / n)
   
es_cohen_w = w

End Function


