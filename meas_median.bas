Attribute VB_Name = "meas_median"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076

Public Sub me_median_addHelp()
Application.MacroOptions _
    Macro:="me_median", _
    Description:="Median", _
    category:=14, _
    ArgumentDescriptions:=Array( _
        "vertical specific range with data", _
        "optional vertical range with labels in order if data is non-numeric.", _
        "optional optional which to return if median falls between two values, either between (default), low or high")
               
End Sub
Function me_median(data As Range, Optional levels As Range, Optional tieBreaker = "between")
    
    'get data as numeric values
    If levels Is Nothing Then
        dataN = he_range_to_num_array(data)
    Else
        dataN = he_replace(data, levels)
    End If
    
    'sort the numeric values
    dataN = he_sort(dataN)
    
    n = UBound(dataN, 1) - LBound(dataN, 1) + 1
    'Note the -1 is for the adjustment to an array ranging from 0 to n-1
    medIndex = (n + 1) / 2 - 1
  
    If medIndex = Round(medIndex, 0) Then
        medN = dataN(medIndex)
    Else
        medLow = dataN(medIndex - 0.5)
        medHigh = dataN(medIndex + 0.5)
        
        If tieBreaker = "low" Then
            medN = medLow
        ElseIf tieBreaker = "high" Then
            medN = medHigh
        Else
            medN = (medLow + medHigh) / 2
        End If
    End If
    
    If levels Is Nothing Then
        med = medN
    Else
        If medN = Round(medN, 0) Then
            med = levels(medN + 1)
        Else
            med = "between " + medLow + " and " + medHigh
        End If
    End If
    
    
    me_median = med

End Function
