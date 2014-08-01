Sub 刷新项目人员()
'
' 刷新项目人员 宏
'

For i = 3 To 200 Step 7
    If Sheets("项目基本状态").Range("A" & i).Value = "" Then Exit For
    For j = 5 To 30
        If Sheets("人力使用").Range("A" & j).Value = "" Then Exit For
        For k = 70 To 90 Step 3
                                    '从“人力使用”取出有效时间在相应项目中的人员
            If Sheets("人力使用").Range(Chr(k) & j).Value = "" Then GoTo Lab1
            If InStr("@" & Sheets("人力使用").Range(Chr(k) & j).Value, Sheets("项目基本状态").Range("A" & i).Value) > 0 And Sheets("人力使用").Range(Chr(k + 2) & j).Value >= Date Then
                Sheets("项目基本状态").Range("G" & i + 2).Value = Sheets("项目基本状态").Range("G" & i + 2).Value + "、" + Sheets("人力使用").Range("A" & j).Value
            End If
        
        Next k
        
        For k = 65 To 88 Step 3
                                    '功能同上，取AA以后的列
            If Sheets("人力使用").Range("A" & Chr(k) & j).Value = "" Then GoTo Lab1
            If InStr("@" & Sheets("人力使用").Range("A" & Chr(k) & j).Value, Sheets("项目基本状态").Range("A" & i).Value) > 0 And Sheets("人力使用").Range("A" & Chr(k + 2) & j).Value >= Date Then
                Sheets("项目基本状态").Range("G" & i + 2).Value = Sheets("项目基本状态").Range("G" & i + 2).Value + "、" + Sheets("人力使用").Range("A" & j).Value
            End If
        
        Next k
Lab1:
    Next j
Next i


For i = 3 To 200 Step 7
                                    '从“项目资源和状态”表中取出项目人数及实际进度(and合同额)
    If Sheets("项目基本状态").Range("A" & i).Value = "" Then Exit For
    For j = 3 To 30
            If Sheets("项目资源及状态").Range("A" & j).Value = "" Then Exit For
            If InStr("@" & Sheets("项目资源及状态").Range("A" & j).Value, Sheets("项目基本状态").Range("A" & i).Value) > 0 Then
                Sheets("项目基本状态").Range("E" & i).Value = Sheets("项目资源及状态").Range("E" & j).Value
                Sheets("项目基本状态").Range("D" & i).Value = Sheets("项目资源及状态").Range("C" & j).Value
            End If
    Next j
    
    For k = 65 To 90
        If Sheets("项目占用人员明细").Range(Chr(k) & 2).Value = "" Then Exit For
        If InStr("@" & Sheets("项目占用人员明细").Range(Chr(k) & 2).Value, Sheets("项目基本状态").Range("A" & i).Value) > 0 Then
            Sheets("项目基本状态").Range("J" & i).Value = Sheets("项目占用人员明细").Range(Chr(k) & 3).Value
        End If
    Next k
    
Next i


For i = 3 To 200 Step 7
                                    '去掉字头的“、”
    If Sheets("项目基本状态").Range("A" & i).Value = "" Then Exit For
    If Left(Sheets("项目基本状态").Range("G" & i + 2).Value, 1) = "、" Then
        Sheets("项目基本状态").Range("G" & i + 2).Value = Right(Sheets("项目基本状态").Range("G" & i + 2).Value, Len(Sheets("项目基本状态").Range("G" & i + 2).Value) - 1)
    End If
Next i

'If Sheets("人力使用").Range("H" & 7).Value > Date Then
'    Sheets("人力使用").Range("H" & 6).Value = 0
'Else
'    Sheets("人力使用").Range("H" & 6).Value = 1
'
'End If

End Sub

Sub 清理人员()
'
' 清理人员 宏
'
'Workbooks("x项目汇总需求信息.xlsm").Sheets("项目基本状态").Range("G" & 4) = Workbooks("项目汇总需求信息0.2.xlsm").Sheets("人力使用").Range("C" & 7)
'调用不同工作簿的数据 例子


For i = 3 To 200 Step 7
    If Sheets("项目基本状态").Range("A" & i).Value = "" Then Exit For
    
    Sheets("项目基本状态").Range("G" & i + 2).Value = ""
Next i


End Sub
