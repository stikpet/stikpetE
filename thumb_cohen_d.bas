Attribute VB_Name = "thumb_cohen_d"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076

Public Sub th_cohen_d_addHelp()
Application.MacroOptions _
    Macro:="th_cohen_d", _
    Description:="Rules of Thumb for Cohen d", _
    category:=14, _
    ArgumentDescriptions:=Array( _
        "the Cohen d value", _
        "optional optional the rule of thumb to be used. Either: cohen, lovakov, rosenthal, or sawilowsky (default)")
End Sub

Function th_cohen_d(d, Optional qual = "sawilowsky")
    'Cohen (1988, p. 40).
    If qual = "cohen" Then
        If Abs(d) < 0.2 Then
        qual = "negligible"
        ElseIf Abs(d) < 0.5 Then
        qual = "small"
        ElseIf Abs(d) < 0.8 Then
        qual = "medium"
        Else
        qual = "large"
        End If
    End If
    
    'Lovakov and Agadullina (2021, p. 501)
    If qual = "lovakov" Then
        If Abs(d) < 0.15 Then
        qual = "negligible"
        ElseIf Abs(d) < 0.35 Then
        qual = "small"
        ElseIf Abs(d) < 0.65 Then
        qual = "medium"
        Else
        qual = "large"
        End If
    End If
    
    'Rosenthal (1996, p. 45).
    If qual = "rosenthal" Then
        If Abs(d) < 0.2 Then
        qual = "negligible"
        ElseIf Abs(d) < 0.5 Then
        qual = "small"
        ElseIf Abs(d) < 0.8 Then
        qual = "medium"
        ElseIf Abs(d) < 1.3 Then
        qual = "large"
        Else
        qual = "very large"
        End If
    End If
    
    'Sawilowsky (2009, p. 599)
    If qual = "sawilowsky" Then
        If Abs(d) < 0.1 Then
        qual = "negligible"
        ElseIf Abs(d) < 0.2 Then
        qual = "very small"
        ElseIf Abs(d) < 0.5 Then
        qual = "small"
        ElseIf Abs(d) < 0.8 Then
        qual = "medium"
        ElseIf Abs(d) < 1.2 Then
        qual = "large"
        ElseIf Abs(d) < 2 Then
        qual = "very large"
        Else
        qual = "huge"
        End If
    End If
    
    th_cohen_d = qual
    
End Function
   
    
