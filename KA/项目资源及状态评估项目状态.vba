Sub 更新项目状态()
'
' 更新项目状态 宏
'
    Dim a, limit
    a = 0
    limit1 = 0.3
    limit2 = 0.15
    alarm1 = 0.1
    alarm2 = 0.8
    
For i = 3 To 100
    If Range("A" & i).Value = "" And Range("J" & i).Value = "" And Range("K" & i).Value = "" Then Exit For
    If Range("J" & i).Value = "" Or Range("K" & i).Value = "" Then GoTo Lab1
    a = (Range("K" & i).Value - Date) / (Range("K" & i).Value - Range("J" & i).Value)
    If a < limit2 And Range("E" & i).Value <= alarm2 Then
        With Range("A" & i).Interior
        .Color = RGB(220, 20, 60)
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .TintAndShade = 0.2
        End With
    ElseIf (a > limit2 And a < limit1 And Range("F" & i).Value < -alarm1) Or Range("F" & i).Value < -(1 - alarm2) Then
        With Range("A" & i).Interior
        .Color = RGB(255, 175, 0)
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .TintAndShade = 0
        End With
    Else
        With Range("A" & i).Interior
        .Color = RGB(144, 238, 144)
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .TintAndShade = 0
        End With
    
Lab1:
    End If
Next i
    
    
'    Range("A3").Select
'    With Selection.Interior
'        .Pattern = xlSolid
'        .PatternColorIndex = xlAutomatic
'        .Color = 254
'        .TintAndShade = 0
'        .PatternTintAndShade = 0
'    End With

'
End Sub

Sub 恢复项目状态()
'
' 恢复项目状态 宏
'
    
For i = 3 To 100
    If Range("A" & i).Value = "" And Range("J" & i).Value = "" And Range("K" & i).Value = "" Then Exit For
        With Range("A" & i).Interior
        .Color = RGB(144, 238, 144)
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .TintAndShade = 0
        End With
Next i
    
    
'    Range("A3").Select
'    With Selection.Interior
'        .Pattern = xlSolid
'        .PatternColorIndex = xlAutomatic
'        .Color = 254
'        .TintAndShade = 0
'        .PatternTintAndShade = 0
'    End With

'
End Sub


