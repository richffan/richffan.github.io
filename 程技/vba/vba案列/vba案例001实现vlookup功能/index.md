# 【VBA案例001】实现VLOOKUP功能

## VBA实现VLOOKUP函数功能

| 数据               |    |   | VBA  |    |
|------------------|----|---|------|----|
| 姓名               | 年龄 |   | 姓名   | 年龄 |
| 潘全桂              | 24 |   | 荆琛泽  |    |
| 霍栋保              | 35 |   | 吉栋松  |    |
| 荆琛泽              | 24 |   | 百里刚晓 |    |
| 越伦信              | 25 |   | 农康雪  |    |
| 吉栋松              | 34 |   | 越伦信  |    |
| 桂真顺              | 27 |   | 霍栋保  |    |
| 百里刚晓             | 19 |   | 潘全桂  |    |
| 农康雪              | 33 |   | 桂真顺  |    |


## 直接附上VBA代码：

```vb
&#39;Option Explicit

Sub 单元格循环()
    Dim cel As Range
    Dim cel2 As Range
    Dim t As Double
    t = Timer
    [e6:e13] = &#34;&#34;
    For Each cel In Range(&#34;a6:a13&#34;)
        For Each cel2 In Range(&#34;d6:d13&#34;)
            If cel.Value = cel2.Value Then
                cel2.Offset(0, 1).Value = cel.Offset(0, 1).Value
                Exit For
            End If
        Next
    Next
    Debug.Print Format(Timer - t, &#34;0.00000000000000s&#34;)
End Sub

Sub 数组循环()
    Dim ar, br
    Dim t As Double
    t = Timer
    [e6:e13] = &#34;&#34;
    ar = [a6:b13] &#39;range(&#34;a6:b13&#34;)
    br = [d6:e13]

    Dim i, j
    For i = 1 To UBound(ar)
        For j = 1 To UBound(ar)
            If ar(i, 1) = br(j, 1) Then
                br(j, 2) = ar(i, 2)
                Exit For
            End If
        Next j
    Next i

    [d6:e13] = br
    Debug.Print Format(Timer - t, &#34;0.00000000000000s&#34;)
End Sub

Sub 字典循环()
    Dim d As Object, kw$
    Set d = CreateObject(&#34;Scripting.Dictionary&#34;)
    &#39;d.CompareMode = vbTextCompare &#39;不区分大小写
    Dim ar, br
    Dim t As Double
    t = Timer
    [e6:e13] = &#34;&#34;
    ar = [a6:b13] &#39;range(&#34;a6:b13&#34;)
    br = [d6:e13]

    Dim i, j
    For i = 1 To UBound(ar)
        d(ar(i, 1)) = ar(i, 2) &#39;KEY ITEM
    Next i

    For j = 1 To UBound(br)
        br(j, 2) = d(br(j, 1))
    Next j

    [d6:e13] = br
    Debug.Print Format(Timer - t, &#34;0.00000000000000s&#34;)
End Sub
```

[原始链接](https://mp.weixin.qq.com/s?__biz=MzIyOTc3NzQ2NA==&amp;mid=2247484991&amp;idx=1&amp;sn=a927b9a6a9a79b2c568fe0e9dbbbc307&amp;chksm=e8bccf68dfcb467e9b7852aa0733798a77f134e89be55e732fee7c57afced35a919993e62809&amp;scene=178&amp;cur_album_id=3115603487041503237#rd)

---

> 作者:   
> URL: http://richfan.site/%E7%A8%8B%E6%8A%80/vba/vba%E6%A1%88%E5%88%97/vba%E6%A1%88%E4%BE%8B001%E5%AE%9E%E7%8E%B0vlookup%E5%8A%9F%E8%83%BD/  

