Attribute VB_Name = "cor_rosenthal"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076

Public Sub r_rosenthal_addHelp()
Application.MacroOptions _
    Macro:="r_rosenthal", _
    Description:="Calculate a Rosenthal Correlation Coefficient", _
    category:=14, _
    ArgumentDescriptions:=Array( _
        "the z-value of test", _
        "the sample size")
        
End Sub
Function r_rosenthal(zVal, n)
'function to calculate a Rosenthal Correlation Coefficient
'zVal : z-value of test
'n : total sample size

    r = zVal / Sqr(n)

    r_rosenthal = r

End Function

