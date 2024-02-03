# 【VBA案例017】合并单元格

大家好！

合并单元格是经常遇到的操作，在WPS中，提供了非常好用的快捷按钮。遗憾的是Excel里并没有这个一键合并单元格的功能。

今天分享用VBA合并单元格的两个最常用的方法，如果你是WPS用户，虽然不需要代码，但是编程的思路，还是有参考价值的。

![](https://img.richfan.site/program/vba/vba案列/【VBA案例017】合并单元格.gif)

以下是VBA代码。详细解析请看文末的视频。

## 方法一：

```vb
Sub union并集函数()

    Dim i
    Dim rng As Range

    Set rng = [a2]

    Application.DisplayAlerts = False
    For i = 2 To 20
        If Range(&#34;a&#34; &amp; i &#43; 1) = Range(&#34;a&#34; &amp; i) Then
            Set rng = Union(rng, Range(&#34;a&#34; &amp; i &#43; 1))
        Else
            rng.Merge
            Set rng = Range(&#34;a&#34; &amp; i &#43; 1)
        End If
    Next i
    Application.DisplayAlerts = True

End Sub
```

## 方法二：

```vb
Sub 循环数组()

    Dim i, j, ar
    Dim start_row, end_row

    ar = [a1:a20]
    start_row = 2

    Application.DisplayAlerts = False
    For i = 2 To UBound(ar)
        For j = i &#43; 1 To UBound(ar)
            If ar(j, 1) = ar(i, 1) Then
                end_row = j
            Else
                end_row = j - 1
                Exit For
            End If
        Next j
        Range(&#34;a&#34; &amp; start_row &amp; &#34;:a&#34; &amp; end_row).Merge
        start_row = j
        i = start_row - 1
    Next i
    Application.DisplayAlerts = True

End Sub
```

[原始链接](https://mp.weixin.qq.com/s?__biz=MzIyOTc3NzQ2NA==&amp;mid=2247485264&amp;idx=1&amp;sn=6307b2628df616ca9d0e4246b1e32d7a&amp;chksm=e8bcce07dfcb4711cce31a831ada9b655400d1e573b0d8e57c70fcfafbaec6efde3d9bb9a2f8&amp;scene=178&amp;cur_album_id=3115603487041503237#rd)


---

> 作者:   
> URL: http://richfan.site/%E7%A8%8B%E6%8A%80/vba/vba%E6%A1%88%E5%88%97/vba%E6%A1%88%E4%BE%8B017%E5%90%88%E5%B9%B6%E5%8D%95%E5%85%83%E6%A0%BC/  

