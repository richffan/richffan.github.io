# 【VBA案例015】拆分工作表（下）

大家好！书接上文，继续聊一聊拆分工作表的第二个方法。

众所周知，字典中的值不仅可以是数字、字符串，还可以是数组和对象！

上一个方法是将数组装到了字典里，这第二个方法想必大家已经猜到了，就是把对象装进字典里。

![](https://img.richfan.site/program/vba/vba案列/【VBA案例015】拆分工作表（下）.gif)

首先简单介绍一下Union函数的使用方法。Union：返回两个或多个区域的合并区域。

语法：
   Union(Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30) AS Range
参数：
   Arg1 必需 Range类型。
   Arg2 必需 Range类型。
   Arg3– Arg30 可选 Variant类型。
返回：
   Range类型。
另外一点需要强调的是，VBA中给对象变量赋值使用Set，并且Set不能省略。

以下是VBA代码。详细解析请看文末的视频。

```vb
Sub range装进字典()
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
            Set d(kw) = Union(Range("a1:e1"), Range("a" & i & ":e" & i))
        Else
            Set d(kw) = Union(d(kw), Range("a" & i & ":e" & i))
        End If
    Next i

    Dim dk
    For Each dk In d.keys
        With ThisWorkbook.Worksheets.Add(after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            .Name = dk
            d(dk).Copy .[a1]
        End With
    Next

End Sub
```

[原始链接](https://mp.weixin.qq.com/s?__biz=MzIyOTc3NzQ2NA==&mid=2247485249&idx=1&sn=a5358d8fe595364149e6f41e84cdec3b&chksm=e8bcce16dfcb47004200062e717b1f72ba0326d94589d37bdfad489431f3e72491eb119824a8&scene=178&cur_album_id=3115603487041503237#rd)

---

> 作者: [richfan](https://richfan.site/)  
> URL: http://richfan.site/program/vba/vba%E6%A1%88%E4%BE%8B/vba%E6%A1%88%E4%BE%8B015%E6%8B%86%E5%88%86%E5%B7%A5%E4%BD%9C%E8%A1%A8%E4%B8%8B/  

