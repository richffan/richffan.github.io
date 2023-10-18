# c数字



## c数字

<!--more-->

```VBA
Sub 千分位符(control As IRibbonControl) '数字-千分位符
    On Error Resume Next
    Dim i As Range, Acell As Cell, CR As Range
    On Error Resume Next
    Application.ScreenUpdating = False
    If Selection.Type = 2 Then '文档选定
        For Each i In Selection.Words
            If IsNumeric(i) Then
                If i Like "####*" = True Then
                    If i.Next Like "." = True And i.Next(wdWord, 2) Like "#*" = True Then
                        i.SetRange Start:=i.Start, End:=i.Next(wdWord, 2).End
                        NC = Format(i, "#,##0.00;-#,##0.00; ")
                        i.Text = NC
                    Else
                        NC = Format(i, "#,##0.00;-#,##0.00; ")
                        i.Text = NC
                    End If
                End If
            End If
        Next i
    ElseIf Selection.Type = 4 Or Selection.Type = 5 Then '竖形表格（5为横形表格）
        For Each Acell In Selection.Cells
            Set CR = ActiveDocument.Range(Acell.Range.Start, Acell.Range.End - 1)
            '            MsgBox CR
            If CR Like "-####*" Or "-####.#*" = True Then
                Yn = Format(CR, "#,##0.00;-#,##0.00; ")
                CR.Text = Yn
            Else
                If CR Like "####*" Or "####.#*" = True Then
                    Yn = Format(CR, "#,##0.00;-#,##0.00; ")
                    CR.Text = Yn
                End If
            End If
        Next Acell
    Else
        MsgBox "您只能选定文本或者表格之一!", vbOK + vbInformation
    End If
    Application.ScreenUpdating = True
    Application.Activate
End Sub
Sub 除以一万(control As IRibbonControl)
    Application.ScreenUpdating = False
    If Selection.Type = 2 Then
        If IsNumeric(Selection.Text) Then
            i = Selection.Text
            P = i / 10000
            q = Format(Round(P, 2), "#,##0.00;-#,##0.00; ")
            Selection.Text = q & "万"
        End If
    ElseIf Selection.Type = 5 Or Selection.Type = 4 Then
        For Each Acell In Selection.Cells
            Set CR = ActiveDocument.Range(Acell.Range.Start, Acell.Range.End - 1)
            'MsgBox CR
            If IsNumeric(CR.Text) Then
                i = CR.Text
                P = i / 10000
                q = Format(Round(P, 2), "#,##0.00;-#,##0.00; ")
                CR.Text = q & "万"
            End If
        Next
    End If
    Application.ScreenUpdating = True
End Sub
Sub 乘百(control As IRibbonControl)
    Application.ScreenUpdating = False
    If Selection.Type = 2 Then
        If IsNumeric(Selection.Text) Then
            i = Selection.Text
            P = i * 100
            q = Format(Round(P, 2), "#,##0.00;-#,##0.00; ")
            Selection.Text = q & "%"
        End If
    ElseIf Selection.Type = 5 Or Selection.Type = 4 Then
        For Each Acell In Selection.Cells
            Set CR = ActiveDocument.Range(Acell.Range.Start, Acell.Range.End - 1)
            'MsgBox CR
            If IsNumeric(CR.Text) Then
                i = CR.Text
                P = i * 100
                q = Format(Round(P, 2), "#,##0.00;-#,##0.00; ")
                CR.Text = q & "%"
            End If
        Next
    End If
        Application.ScreenUpdating = True
End Sub
```



---

> 作者: [richfan](https://richfan.site/)  
> URL: http://richfan.site/c%E6%95%B0%E5%AD%97/  

