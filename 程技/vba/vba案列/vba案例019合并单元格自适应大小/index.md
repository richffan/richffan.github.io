# 【VBA案例019】合并单元格自适应大小

大家好！

如果你是文职类工作，可能会遇到下面这种情况：

经常面对各种各样的表格，并且很多都是制式的，里边又充满个各种各样的格式，其中就有今天的主角儿：合并单元格。

而你的工作看似也不复杂，就是把合并单元格中显示不全的内容，通过调整单元格的大小来显示出来。

这种痛苦，只有手动调整过的人能懂。

所以，通过今天的案例讲解，将解决你的烦恼，文末视频对这个过程做了详细的讲解演示，希望对你有帮助。

![](https://img.richfan.site/program/vba/vba案列/【VBA案例019】合并单元格自适应大小.gif)

以下是VBA代码。详细解析请看文末的视频。

调整合并单元格行高：

```vb
Sub 调整合并单元格行高()
    Dim cel As Range
    Dim rng As Range
    Dim n, r, c
    Dim mergeWidth, newHeight, celWidth

    Set rng = Range(&#34;B4&#34;)
    For Each cel In rng
        If cel.MergeCells Then
            With cel.MergeArea
                mergeWidth = 0
                For Each c In .Columns &#39;合并区域中的每一个单元格
                    mergeWidth = mergeWidth &#43; c.ColumnWidth &#39;新合并区域列宽=每一列列宽宽的和
                Next

                .MergeCells = False &#39;取消合并单元格

                With .Cells(1)
                    .WrapText = True &#39;自动换行
                    celWidth = .ColumnWidth &#39;记录取消合并后列宽，目的是调整回去
                    .ColumnWidth = mergeWidth &#39;调整第一个单元格宽度
                    .EntireRow.AutoFit &#39;自适应大小
                    newHeight = .RowHeight &#39;记录此时的行高，即合并后新的行高
                    .ColumnWidth = celWidth &#39;调整回原始列宽
                End With

                .MergeCells = True &#39;合并单元格

                n = .Rows.Count
                For Each r In .Rows
                    r.RowHeight = newHeight / n * 1.1 &#39;调整每一行的行高，并*1.1微调
                Next
                .HorizontalAlignment = xlCenter &#39;左右居中
                .VerticalAlignment = xlCenter &#39;&#39;上下居中
            End With
        End If
    Next
End Sub
```

调整合并单元格列宽：

```vb
Sub 调整合并单元格列宽()
    Dim cel As Range
    Dim rng As Range
    Dim n, r, c
    Dim mergeHeight, newWidth, celHeight

    Set rng = Range(&#34;B4&#34;)
    For Each cel In rng
        If cel.MergeCells Then
            With cel.MergeArea

                mergeHeight = 0
                For Each r In .Rows &#39;合并区域中的每一个行
                    mergeHeight = mergeHeight &#43; r.RowHeight &#39;新合并区域行高=每一行行高的和
                Next

                .MergeCells = False &#39;取消合并单元格

                With .Cells(1)
                    .WrapText = True &#39;自动换行
                    celHeight = .RowHeight &#39;记录取消合并后行高，目的是调整回去
                    .RowHeight = mergeHeight &#39;调整第一个单元格高度
                    .EntireColumn.AutoFit &#39;列宽自适应大小
                    newWidth = .ColumnWidth &#39;记录此时的列宽，即合并后新的列宽
                    .RowHeight = celHeight &#39;调整回原始列宽
                End With

                .MergeCells = True &#39;合并单元格

                n = .Columns.Count
                For Each c In .Columns
                    c.ColumnWidth = newWidth / n * 1.1 &#39;调整每一列的列宽，并*1.1微调
                Next
                .HorizontalAlignment = xlCenter &#39;左右居中
                .VerticalAlignment = xlCenter &#39;上下居中
            End With
        End If
    Next
End Sub
```

[原始链接](https://mp.weixin.qq.com/s?__biz=MzIyOTc3NzQ2NA==&amp;mid=2247485287&amp;idx=1&amp;sn=022ca1d3312a748fd7591b522a993764&amp;chksm=e8bcce30dfcb4726d09083aa07afaaf286c376e4dba0955139f28a9a8d5afa28e401231498b2&amp;scene=178&amp;cur_album_id=3115603487041503237#rd)

---

> 作者:   
> URL: http://richfan.site/%E7%A8%8B%E6%8A%80/vba/vba%E6%A1%88%E5%88%97/vba%E6%A1%88%E4%BE%8B019%E5%90%88%E5%B9%B6%E5%8D%95%E5%85%83%E6%A0%BC%E8%87%AA%E9%80%82%E5%BA%94%E5%A4%A7%E5%B0%8F/  

