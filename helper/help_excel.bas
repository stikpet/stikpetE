Attribute VB_Name = "help_excel"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076

'this makes these not appear in Excel, but still accesible in VBA
Option Private Module

Function he_range_to_num_array(data As Range)
'function to convert a range to an array with numeric value from 0 to n-1
    nr = data.Rows.Count
    nData = WorksheetFunction.Count(data)
    
    
    Dim dataArr() As Double
    ReDim dataArr(0 To nData - 1)
    Dim i, k As Integer
    k = 0
    i = 1
    Do While k < nData
        If Not IsEmpty(data(i)) And WorksheetFunction.IsNumber(data(i)) Then
            dataArr(k) = data(i)
            k = k + 1
        End If
        
        i = i + 1
    Loop
    
    he_range_to_num_array = dataArr

End Function
Function he_sort(data)
'function to sort a numeric array

nr = UBound(data, 1) - LBound(data, 1) + 1

Dim dataSorted() As Double
ReDim dataSorted(0 To nr - 1)

'sort scores
dataSorted = data
changes = 1
Do While changes <> 0
    changes = 0
    For i = 1 To nr - 1
        If dataSorted(i - 1) > dataSorted(i) Then
            ff1 = dataSorted(i)
            dataSorted(i) = dataSorted(i - 1)
            dataSorted(i - 1) = ff1
            changes = 1
        End If
    Next i
Loop

he_sort = dataSorted

End Function

Function he_replace(data As Range, levels As Range)

    n = 0
    For i = 1 To levels.Rows.Count
        n = n + WorksheetFunction.CountIf(data, levels(i))
    Next i
    
    Dim dataRec() As Double
    ReDim dataRec(0 To n - 1)
    nr = 0
    For i = 1 To data.Rows.Count
        For j = 1 To levels.Rows.Count
            If data(i) = levels(j) Then
                dataRec(nr) = j
                nr = nr + 1
            End If
        Next j
    Next i
            
    
    he_replace = dataRec


End Function

Function he_srf(k, n)
'signed ranks frequencies

If k < 0 Then
    he_srf = 0
ElseIf k > WorksheetFunction.Combin(n + 1, 2) Then
    he_srf = 0

ElseIf n = 1 And (k = 0 Or k = 1) Then
    he_srf = 1
Else
    he_srf = he_srf(k - n, n - 1) + he_srf(k, n - 1)
End If


