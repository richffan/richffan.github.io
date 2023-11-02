# 【VBA案例014】拆分工作表（上）

大家好！如何按照表中的某一列，拆分成独立的Sheet? 如下：

![](https://img.richfan.site/program/vba/vba案列/【VBA案例014】拆分工作表（上）.gif)

这是一个特别常见常用的问题，本期分享本人用的最多的两个方法中的第一个。

因为确实不太容易理解，所以分为两部分。

这个方法非常的实用，在其他地方也可以发挥很大的作用，所以墙裂推荐大家掌握！

以下是VBA代码。详细解析请看文末的视频。

```vb
Sub 数组装进字典()
    Dim i, j, k
    Dim ar, tmp()
    Dim d As Object, kw$
    Set d = CreateObject("Scripting.Dictionary")
    'd.CompareMode = vbTextCompare '不区分大小写

    ar = Range("a1:e" & [a65536].End(3).Row)
    Dim irow
    For i = 2 To UBound(ar)
        kw = ar(i, 4)
        If Not d.exists(kw) Then
            ReDim tmp(1 To 5000, 1 To UBound(ar, 2) + 1)
            For j = 1 To UBound(ar, 2)
                tmp(1, j) = ar(1, j)
                tmp(2, j) = ar(i, j)
            Next
            tmp(1, UBound(ar, 2) + 1) = 2
            d(kw) = tmp
        Else
            tmp = d(kw)
            irow = tmp(1, UBound(ar, 2) + 1) + 1
            For j = 1 To UBound(ar, 2)
                tmp(irow, j) = ar(i, j)
            Next
            tmp(1, UBound(ar, 2) + 1) = irow
            d(kw) = tmp
        End If
    Next i

    Dim dk
    For Each dk In d.keys
        With ThisWorkbook.Worksheets.Add(after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            .Name = dk
            tmp = d(dk)
            .[a1].Resize(tmp(1, UBound(ar, 2) + 1), UBound(ar, 2)) = tmp
        End With
    Next

End Sub
```

[原始链接](https://mp.weixin.qq.com/s?__biz=MzIyOTc3NzQ2NA==&mid=2247485240&idx=1&sn=4fb6d29d247c9738f7c8c2ad5041c54c&chksm=e8bcce6fdfcb47790d046fc09014a1e28640e51b51b4f5492888e5fd1672c27fa4628698119b&scene=178&cur_album_id=3115603487041503237#rd)

---

> 作者: [richfan](https://richfan.site/)  
> URL: http://richfan.site/%E7%BC%96%E7%A8%8B/vba/vba%E6%A1%88%E5%88%97/vba%E6%A1%88%E4%BE%8B014%E6%8B%86%E5%88%86%E5%B7%A5%E4%BD%9C%E8%A1%A8%E4%B8%8A/  

