Sub 更新项目占用人员明细()

    For i = 5 To 30
        If Sheets("人力使用").Range("A" & i).Value = "" Or Sheets("项目占用人员明细").Range("A" & i).Value = "" Then Exit Sub
        For j = 66 To 90
            If Sheets("项目占用人员明细").Range(Chr(j) & 2).Value = "" Then Exit For
            For k = 70 To 90 Step 3
                If Sheets("人力使用").Range(Chr(k) & i).Value = "" Then Exit For
                If InStr("@" & Sheets("人力使用").Range(Chr(k) & i).Value, Sheets("项目占用人员明细").Range(Chr(j) & 2).Value) > 0 Then
                    If Sheets("人力使用").Range(Chr(k + 2) & i).Value = "" Or Sheets("人力使用").Range(Chr(k + 2) & i).Value = "待定" Or Sheets("人力使用").Range(Chr(k + 2) & i).Value = "长期维护" Then
                        Sheets("项目占用人员明细").Range(Chr(j) & i).Value = "占用"
                    Else
                        Sheets("项目占用人员明细").Range(Chr(j) & i).Value = Sheets("人力使用").Range(Chr(k + 2) & i).Value
                    End If
                End If
            Next k
        Next j
    Next i
    
    
'    Selection.NumberFormatLocal = "m""月""d""日"";@"
End Sub
