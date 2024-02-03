# 【VBA案例007】多条件汇总


大家好！今天回答一位网友的问题。

就是用VBA进行多条件汇总，来实现数据透视表的效果，并且要对结果进行排序。

先来看例子。

假如我们有一份产品信息表，需要对它的所有产品和型号进行汇总。左侧是原始数据，右侧是处理结果。

![](https://img.richfan.site/program/vba/vba案列/【VBA案例007】多条件汇总.png)

我们来通过三个不同的方法，来解决这个问题。其中方法一：最容易理解，适合对字典刚入门的情况。方法二：具有有一定的难度，需要对字典有更加深刻的了解。方法三：作为拓展内容。

以下是VBA代码，你也可以直接观看下方的视频解析：

## 方法一：

```vb
Sub 方法一()
    Dim i, j, k
    Dim ar

    Dim d1 As Object, d2 As Object, d3 As Object
    Set d1 = CreateObject(&#34;Scripting.Dictionary&#34;)
    Set d2 = CreateObject(&#34;Scripting.Dictionary&#34;)
    Set d3 = CreateObject(&#34;Scripting.Dictionary&#34;)

    ar = Range(&#34;a1:c&#34; &amp; [a65536].End(3).Row)

    For i = 2 To UBound(ar)
        d1(ar(i, 1)) = &#34;&#34;
        d2(ar(i, 2)) = &#34;&#34;
        d3(ar(i, 1) &amp; ar(i, 2)) = d3(ar(i, 1) &amp; ar(i, 2)) &#43; ar(i, 3)
    Next i

    [f2].Resize(d1.Count) = Application.WorksheetFunction.Transpose(d1.keys)
    [g1].Resize(1, d2.Count) = d2.keys

    For i = 1 To d1.Count
        For j = 1 To d2.Count
            Cells(i &#43; 1, j &#43; 6) = d3(Cells(i &#43; 1, 6).Value &amp; Cells(1, j &#43; 6).Value)
        Next j
    Next i

    &#39;range.Sort
    Range(&#34;f1&#34;).Resize(d1.Count &#43; 1, d2.Count &#43; 1).Sort [f1], xlAscending, , , , , , xlYes, , , xlTopToBottom
    Range(&#34;g1&#34;).Resize(d1.Count &#43; 1, d2.Count).Sort [g1], xlAscending, , , , , , , , , xlLeftToRight
End Sub
```

## 方法二：

```vb
Sub 方法二()
    Dim i, j, k
    Dim ar, br()
    Dim d As Object, kw$
    Set d = CreateObject(&#34;Scripting.Dictionary&#34;)
    &#39;d.CompareMode = vbTextCompare &#39;不区分大小写
    ar = Range(&#34;a1:c&#34; &amp; [a65536].End(3).Row)

    ReDim br(1 To UBound(ar), 1 To 1000)
    Dim rowNum, colNum
    rowNum = 1: colNum = 1
    For i = 2 To UBound(ar)
        &#39;\\型号
        If Not d.exists(ar(i, 2)) Then
            colNum = colNum &#43; 1
            br(1, colNum) = ar(i, 2)
            d(ar(i, 2)) = colNum
        End If

        &#39;\\名称
        If Not d.exists(ar(i, 1)) Then
            rowNum = rowNum &#43; 1
            br(rowNum, 1) = ar(i, 1)
            d(ar(i, 1)) = rowNum
            br(rowNum, d(ar(i, 2))) = ar(i, 3)
        Else
            br(d(ar(i, 1)), d(ar(i, 2))) = br(d(ar(i, 1)), d(ar(i, 2))) &#43; ar(i, 3)
        End If

    Next i

    [f1].Resize(rowNum, colNum) = br
    Range(&#34;f1&#34;).Resize(rowNum, colNum).Sort [f1], xlAscending, , , , , , xlYes, , , xlTopToBottom
    Range(&#34;g1&#34;).Resize(rowNum, colNum - 1).Sort [g1], xlAscending, , , , , , , , , xlLeftToRight
End Sub
```

## 方法三：

```vb
Sub SQL查询()
    &#39;定义变量
    Dim cnn, rst, SQL$
    Dim i, j, k
    Set cnn = CreateObject(&#34;adodb.connection&#34;) &#39;创建数据库连接
    Set rst = CreateObject(&#34;adodb.recordset&#34;) &#39;创建一个数据集保存数据

    &#39;设置数据库连接
    If Val(Application.Version) &lt; 12 Then
        cnn.Open &#34;Provider=Microsoft.Jet.Oledb.4.0;Extended Properties=&#39;Excel 8.0;HDR=yes&#39;;Data Source=&#34; &amp; ThisWorkbook.FullName
    Else
        cnn.Open &#34;Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties=&#39;Excel 12.0;HDR=yes&#39;;Data Source=&#34; &amp; ThisWorkbook.FullName
    End If

    &#39;设置SQL语句
    SQL = &#34;TRANSFORM SUM(数量) SELECT 名称 from [Sheet1$a1:c18] GROUP BY 名称 PIVOT 型号&#34; &#39;

    &#39;SQL结果处理
    Set rst = cnn.Execute(SQL)

    Range(&#34;f2&#34;).CopyFromRecordset rst
    For i = 1 To rst.Fields.Count
        Cells(1, i &#43; 5) = rst.Fields(i - 1).Name
    Next

    rst.Close
    cnn.Close &#39;关闭数据库连接
    Set rst = Nothing
    Set cnn = Nothing &#39;将cnn从内存中删除
End Sub
```

[原始链接](https://mp.weixin.qq.com/s?__biz=MzIyOTc3NzQ2NA==&amp;mid=2247485136&amp;idx=1&amp;sn=fcc6dd5ad8e974e4b31bea47b811c8b6&amp;chksm=e8bccf87dfcb46912f89badebc8df22ecbec9bc0390633120dbb78e2e28392a310aa9d5401fe&amp;scene=178&amp;cur_album_id=3115603487041503237#rd)

---

> 作者:   
> URL: http://richfan.site/%E7%A8%8B%E6%8A%80/vba/vba%E6%A1%88%E5%88%97/vba%E6%A1%88%E4%BE%8B007%E5%A4%9A%E6%9D%A1%E4%BB%B6%E6%B1%87%E6%80%BB/  

