Sub 生成表格()
'
' 生成表格 宏
'
n = 0
x = 0
t = 500
For i = 3 To 100
    If Range("A" & i).Value = "" And Range("A" & i + 1).Value = "" And Range("A" & i + 2).Value = "" Then Exit For
    If Range("A" & i).Value = "" Then GoTo Lab0
    Range("AD2:AE2,AD" & i & ":AE" & i).Select
    Range("AD" & i).Activate
    With ActiveSheet.Shapes.AddChart
        .Left = (x + 1) * 35 + x * 200
        .Top = t
        .Width = 200
        .Height = 220
        .Select
    End With
        
    With ActiveChart
        .ChartType = xlPie
        .SetSourceData Source:=Range("AD2:AE2,AD" & i & ":AE" & i)
        .ClearToMatchStyle
        .ChartStyle = 1
'        .ClearToMatchStyle
        .ApplyLayout (2)
        .SeriesCollection(1).Select
        .SeriesCollection(1).DataLabels.Select
    End With
    Selection.NumberFormat = "0.00%"
    With ActiveChart
        .ChartTitle.Text = Range("A" & i)
    End With
    
    n = n + 1
    x = x + 1
    If x Mod 4 = 0 And x <> 0 Then
        t = 500 + 250 * (n / 4)
        x = 0
    End If
Lab0:
Next i
    Range("I30").Select
End Sub

Sub 清除图表()
'
' 清除图表 宏
    If ActiveSheet.ChartObjects.Count = 0 Then
        MsgBox "Sheet中没有图表"
        Exit Sub
    End If
    If ActiveSheet.ChartObjects.Count <> 0 Then
        ActiveSheet.ChartObjects.Delete
    End If
'ActiveSheet.ChartObjects.Delete
End Sub
