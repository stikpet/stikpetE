Attribute VB_Name = "thumb_cohen_h"
Public Sub th_cohen_h_addHelp()
Application.MacroOptions _
    Macro:="th_cohen_d", _
    Description:="Rules of Thumb for Cohen h", _
    category:=14, _
    ArgumentDescriptions:=Array( _
        "the Cohen h value", _
        "optional optional the rule of thumb to be used. Currently only cohen available (also then default)", _
        "output to show, either all (default) for array result, qual for only the classification, or ref for the reference")
End Sub

Function th_cohen_h(h, Optional qual = "cohen", Optional output = "all")

    'Cohen (1988, pp. 184-185)
    If (qual = "cohen") Then
  
        ref = "Cohen (1988, p. 198)"
        
        If (Abs(h) < 0.2) Then
          qual = "negligible"
        ElseIf (Abs(h) < 0.5) Then
          qual = "small"
        ElseIf (Abs(h) < 0.8) Then
            qual = "medium"
        Else
          qual = "large"
        End If
    End If
    
    'Output
    If output = "all" Then
        Dim res(1 To 2, 1 To 2)
        res(1, 1) = "classification"
        res(1, 2) = "source"
        
        res(2, 1) = qual
        res(2, 2) = ref
        th_cohen_h = res
        
    ElseIf output = "ref" Then
        th_cohen_h = ref
    Else
        th_cohen_h = qual
    End If
    

End Function

  
