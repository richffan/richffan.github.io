# 【VBA案例006】数据去重

大家好！你是否在为数据去重感到烦恼？

今天，我们将分享两种高效的 VBA 方法，帮助你轻松应对数据去重难题。

举个例子。

我们要对以下数据的产品和型号进行数据去重，保留唯一值。

![](https://img.richfan.site/program/vba/vba案列/【VBA案例006】数据去重.png)

以下是VBA代码，你也可以直接观看下方的视频解析：

## 方法一：

```vb
Sub 方法一()
    Range("f1:h18").RemoveDuplicates Array(1, 2), xlYes
End Sub
```

## 方法二：

```vb
Sub 方法二第一次出现的值()
    Dim i, j, k
    Dim ar, br()

    ar = Range("a1:c18")
    ReDim br(1 To UBound(ar), 1 To UBound(ar, 2))

    Dim d As Object, kw$
    Set d = CreateObject("Scripting.Dictionary")
    'd.CompareMode = vbTextCompare '不区分大小写

    For i = 1 To UBound(ar)
        kw = ar(i, 1) & ar(i, 2)
        If Not d.exists(kw) Then
            k = k + 1
            For j = 1 To UBound(br, 2)
                br(k, j) = ar(i, j)
            Next j
            d(kw) = ""
        End If
    Next i

    [f1].Resize(k, UBound(br, 2)) = br

End Sub
```

## 方法二：扩展

```vb
Sub 方法二最后值()
    Dim i, j, k
    Dim ar, br()

    ar = Range("a1:c18")
    ReDim br(1 To UBound(ar), 1 To UBound(ar, 2))

    Dim d As Object, kw$
    Set d = CreateObject("Scripting.Dictionary")
    'd.CompareMode = vbTextCompare '不区分大小写

    For i = 1 To UBound(ar)
        kw = ar(i, 1) & ar(i, 2)
        d(kw) = i
    Next i

    Dim dickey
    For Each dickey In d.keys
        k = k + 1
        For j = 1 To UBound(br, 2)
            br(k, j) = ar(d(dickey), j)
        Next j
    Next

    [f1].Resize(k, UBound(br, 2)) = br

End Sub
```

[原始链接](https://mp.weixin.qq.com/s?__biz=MzIyOTc3NzQ2NA==&mid=2247485123&idx=1&sn=74578fe8bd5288519db8f509207b0caf&chksm=e8bccf94dfcb468231cb18e7d79f65bd75b01a4636fd0523cb5f36f565447027255b915e52e9&scene=178&cur_album_id=3115603487041503237#rd)

---

> 作者: [richfan](https://richfan.site/)  
> URL: http://richfan.site/%E7%BC%96%E7%A8%8B/vba/vba%E6%A1%88%E5%88%97/vba%E6%A1%88%E4%BE%8B006%E6%95%B0%E6%8D%AE%E5%8E%BB%E9%87%8D/  

