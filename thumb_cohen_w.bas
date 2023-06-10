Attribute VB_Name = "thumb_cohen_w"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076

Public Sub th_cohen_w_addHelp()
Application.MacroOptions _
    Macro:="th_cohen_w", _
    Description:="Cohen's w rule-of-thumb", _
    category:=14, _
    ArgumentDescriptions:=Array( _
        "Cohen w value", _
        "which rule of thumb, currently only cohen", _
        "output to show, either ref, qual, or both (default)")
End Sub
Function th_cohen_w(w, Optional qual = "cohen", Optional output = "both")

    'Qualification
    
    If qual = "cohen" Then
        'Use Cohen (1988, p. 227).
        ref = "Cohen (1988, p. 227)"
        If Abs(w) < 0.1 Then
            clas = "negligible"
        ElseIf Abs(w) < 0.3 Then
            clas = "small"
        ElseIf Abs(w) < 0.5 Then
            clas = "medium"
        Else
            clas = "large"
        End If
    End If
    
    'the output
    If output = "qual" Then
        th_cohen_w = clas
    ElseIf output = "ref" Then
        th_cohen_w = ref
    Else
        'Results
        Dim res(1 To 2, 1 To 2)
        res(1, 1) = "classification"
        res(1, 2) = "reference"
        res(2, 1) = clas
        res(2, 2) = ref
        
        th_cohen_w = res
    End If
End Function
