# h行列校验


## h行列校验


&lt;!--more--&gt;

```vb
Sub 选行校验(control As IRibbonControl) &#39;行列校验-选行校验
    i = 0
    x = 0
    n = 0
    If Selection.Type = wdSelectionColumn Then
        For Each Acell In Selection.Cells &#39;求所选列合计数
            If Acell.ColumnIndex &gt; n Then
                n = Acell.ColumnIndex
            End If
            Set CR1 = ActiveDocument.Range(Acell.Range.Start, Acell.Range.End - 1)
            If IsNumeric(CR1.Text) Then
                i = CR1.Text &#43; i
            End If
        Next
        For Each Acell In Selection.Cells &#39;求所选列最后一行合计数
            Set CR2 = ActiveDocument.Range(Acell.Range.Start, Acell.Range.End - 1)
            If Acell.ColumnIndex = n Then
                If IsNumeric(CR2.Text) Then
                    x = CR2.Text &#43; x
                End If
            End If
        Next
        P = i - x
        q = P - x
        y = &#34;列校验:&#34; &amp; Format(P, &#34;Standard&#34;) &amp; &#34;-&#34; &amp; Format(x, &#34;Standard&#34;) &amp; &#34;=&#34; &amp; Format(q, &#34;Standard&#34;)
    Else
        y = &#34;只支持选中表格范围!&#34;
    End If
    Application.StatusBar = y
End Sub
Sub 选列校验(control As IRibbonControl) &#39;行列校验-选列校验
    i = 0
    x = 0
    n = 0
    If Selection.Type = wdSelectionColumn Then
        For Each Acell In Selection.Cells &#39;求所选列合计数
            If Acell.RowIndex &gt; n Then
                n = Acell.RowIndex
            End If
            Set CR1 = ActiveDocument.Range(Acell.Range.Start, Acell.Range.End - 1)
            If IsNumeric(CR1.Text) Then
                i = CR1.Text &#43; i
            End If
        Next
        For Each Acell In Selection.Cells &#39;求所选列最后一行合计数
            Set CR2 = ActiveDocument.Range(Acell.Range.Start, Acell.Range.End - 1)
            If Acell.RowIndex = n Then
                If IsNumeric(CR2.Text) Then
                    x = CR2.Text &#43; x
                End If
            End If
        Next
        P = i - x
        q = P - x
        y = &#34;列校验:&#34; &amp; Format(P, &#34;Standard&#34;) &amp; &#34;-&#34; &amp; Format(x, &#34;Standard&#34;) &amp; &#34;=&#34; &amp; Format(q, &#34;Standard&#34;)
    Else
        y = &#34;只支持选中表格范围!&#34;
    End If
    Application.StatusBar = y
End Sub
Sub 区域求和(control As IRibbonControl) &#39;行列校验-区域求和
    i = 0
    If Selection.Type = 5 Then
        For Each Acell In Selection.Cells
            Set CR = ActiveDocument.Range(Acell.Range.Start, Acell.Range.End - 1)
            If IsNumeric(CR.Text) Then
                i = CR.Text &#43; i
            End If
        Next
        i = &#34;合计:&#34; &amp; Format(i, &#34;#,##0.00;-#,##0.00; &#34;)
    ElseIf Selection.Type = 4 Then
        For Each Acell In Selection.Cells
            Set CR = ActiveDocument.Range(Acell.Range.Start, Acell.Range.End - 1)
            If IsNumeric(CR.Text) Then
                i = CR.Text &#43; i
            End If
        Next
        i = &#34;合计:&#34; &amp; Format(i, &#34;#,##0.00;-#,##0.00; &#34;)
    Else
        i = &#34;只支持表格内行或列求和!&#34;
    End If
    Application.StatusBar = i
End Sub
```



---

> 作者:   
> URL: http://richfan.site/%E7%A8%8B%E6%8A%80/vba/%E5%BA%95%E7%A8%BF%E5%B0%8F%E5%B8%AE%E6%89%8B%E4%BB%A3%E7%A0%81/h%E8%A1%8C%E5%88%97%E6%A0%A1%E9%AA%8C/  

