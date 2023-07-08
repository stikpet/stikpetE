Attribute VB_Name = "thumb_cohen_h"
Function th_cohen_h(h, Optional qual = "cohen", Optional output = "all")
Attribute th_cohen_h.VB_Description = "Rules of Thumb for Cohen h"
Attribute th_cohen_h.VB_ProcData.VB_Invoke_Func = " \n14"

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

  
