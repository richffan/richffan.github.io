# 【VBA案例009】合并单元格汇总


大家好！今天分享一下遇到合并单元格的处理思路。

工作中，为了方便查看数据，很多朋友习惯性的把单元格合并起来，这样使得表格看起来清晰美观了不少。但这样做的同时，会给以后的计算增加很多困难。

比如下边：要计算产品的总库存。

![](https://img.richfan.site/program/vba/vba案列/【VBA案例009】合并单元格汇总.gif)

当我们要对合并单元格进行汇总统计的时候，总会遇到很多问题。其实只要你明白合并单元格的底层逻辑之后，问题就会变的简单起来。

合并单元格：其实只保留了区域左上角的一个值，其他位置都是空的。

明白原理之后，我们再来看下VBA代码，详细解析请看文末视频。

```vb
Sub 汇总()

    Dim i, j, k
    Dim ar

    ar = Range("a1:c" & [c65536].End(3).Row)

    Dim d As Object, kw$
    Set d = CreateObject("Scripting.Dictionary")
    'd.CompareMode = vbTextCompare '不区分大小写

    For i = 2 To UBound(ar)
        If ar(i, 1) = "" Then ar(i, 1) = ar(i - 1, 1)
        d(ar(i, 1)) = d(ar(i, 1)) + ar(i, 3)
    Next
    [F2].Resize(d.Count) = Application.WorksheetFunction.Transpose(d.keys)
    [g2].Resize(d.Count) = Application.WorksheetFunction.Transpose(d.items)

End Sub
```

[原始链接](https://mp.weixin.qq.com/s?__biz=MzIyOTc3NzQ2NA==&mid=2247485175&idx=1&sn=21a86379e469df4052c9e26bc5615e4b&chksm=e8bccfa0dfcb46b67aa9719e78f9d5ef4e2316ead39351e037282d360f6fc752126496c0e063&scene=178&cur_album_id=3115603487041503237#rd)

---

> 作者: [richfan](https://richfan.site/)  
> URL: http://richfan.site/program/vba/vba%E6%A1%88%E4%BE%8B/vba%E6%A1%88%E4%BE%8B009%E5%90%88%E5%B9%B6%E5%8D%95%E5%85%83%E6%A0%BC%E6%B1%87%E6%80%BB/  

