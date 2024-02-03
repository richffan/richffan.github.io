# c数字



## c数字

&lt;!--more--&gt;

```vb
Sub 千分位符(control As IRibbonControl) &#39;数字-千分位符
    On Error Resume Next
    Dim i As Range, Acell As Cell, CR As Range
    On Error Resume Next
    Application.ScreenUpdating = False
    If Selection.Type = 2 Then &#39;文档选定
        For Each i In Selection.Words
            If IsNumeric(i) Then
                If i Like &#34;####*&#34; = True Then
                    If i.Next Like &#34;.&#34; = True And i.Next(wdWord, 2) Like &#34;#*&#34; = True Then
                        i.SetRange Start:=i.Start, End:=i.Next(wdWord, 2).End
                        NC = Format(i, &#34;#,##0.00;-#,##0.00; &#34;)
                        i.Text = NC
                    Else
                        NC = Format(i, &#34;#,##0.00;-#,##0.00; &#34;)
                        i.Text = NC
                    End If
                End If
            End If
        Next i
    ElseIf Selection.Type = 4 Or Selection.Type = 5 Then &#39;竖形表格（5为横形表格）
        For Each Acell In Selection.Cells
            Set CR = ActiveDocument.Range(Acell.Range.Start, Acell.Range.End - 1)
            &#39;            MsgBox CR
            If CR Like &#34;-####*&#34; Or &#34;-####.#*&#34; = True Then
                Yn = Format(CR, &#34;#,##0.00;-#,##0.00; &#34;)
                CR.Text = Yn
            Else
                If CR Like &#34;####*&#34; Or &#34;####.#*&#34; = True Then
                    Yn = Format(CR, &#34;#,##0.00;-#,##0.00; &#34;)
                    CR.Text = Yn
                End If
            End If
        Next Acell
    Else
        MsgBox &#34;您只能选定文本或者表格之一!&#34;, vbOK &#43; vbInformation
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
            q = Format(Round(P, 2), &#34;#,##0.00;-#,##0.00; &#34;)
            Selection.Text = q &amp; &#34;万&#34;
        End If
    ElseIf Selection.Type = 5 Or Selection.Type = 4 Then
        For Each Acell In Selection.Cells
            Set CR = ActiveDocument.Range(Acell.Range.Start, Acell.Range.End - 1)
            &#39;MsgBox CR
            If IsNumeric(CR.Text) Then
                i = CR.Text
                P = i / 10000
                q = Format(Round(P, 2), &#34;#,##0.00;-#,##0.00; &#34;)
                CR.Text = q &amp; &#34;万&#34;
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
            q = Format(Round(P, 2), &#34;#,##0.00;-#,##0.00; &#34;)
            Selection.Text = q &amp; &#34;%&#34;
        End If
    ElseIf Selection.Type = 5 Or Selection.Type = 4 Then
        For Each Acell In Selection.Cells
            Set CR = ActiveDocument.Range(Acell.Range.Start, Acell.Range.End - 1)
            &#39;MsgBox CR
            If IsNumeric(CR.Text) Then
                i = CR.Text
                P = i * 100
                q = Format(Round(P, 2), &#34;#,##0.00;-#,##0.00; &#34;)
                CR.Text = q &amp; &#34;%&#34;
            End If
        Next
    End If
        Application.ScreenUpdating = True
End Sub
```



---

> 作者:   
> URL: http://richfan.site/%E7%A8%8B%E6%8A%80/vba/%E5%BA%95%E7%A8%BF%E5%B0%8F%E5%B8%AE%E6%89%8B%E4%BB%A3%E7%A0%81/c%E6%95%B0%E5%AD%97/  

