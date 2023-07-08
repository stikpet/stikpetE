Attribute VB_Name = "cor_rosenthal"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076

Function r_rosenthal(zVal, n)
Attribute r_rosenthal.VB_Description = "Calculate a Rosenthal Correlation Coefficient"
Attribute r_rosenthal.VB_ProcData.VB_Invoke_Func = " \n14"
'function to calculate a Rosenthal Correlation Coefficient
'zVal : z-value of test
'n : total sample size

    r = zVal / Sqr(n)

    r_rosenthal = r

End Function

