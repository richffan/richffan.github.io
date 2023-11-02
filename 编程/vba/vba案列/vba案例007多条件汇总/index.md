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
    Set d1 = CreateObject("Scripting.Dictionary")
    Set d2 = CreateObject("Scripting.Dictionary")
    Set d3 = CreateObject("Scripting.Dictionary")

    ar = Range("a1:c" & [a65536].End(3).Row)

    For i = 2 To UBound(ar)
        d1(ar(i, 1)) = ""
        d2(ar(i, 2)) = ""
        d3(ar(i, 1) & ar(i, 2)) = d3(ar(i, 1) & ar(i, 2)) + ar(i, 3)
    Next i

    [f2].Resize(d1.Count) = Application.WorksheetFunction.Transpose(d1.keys)
    [g1].Resize(1, d2.Count) = d2.keys

    For i = 1 To d1.Count
        For j = 1 To d2.Count
            Cells(i + 1, j + 6) = d3(Cells(i + 1, 6).Value & Cells(1, j + 6).Value)
        Next j
    Next i

    'range.Sort
    Range("f1").Resize(d1.Count + 1, d2.Count + 1).Sort [f1], xlAscending, , , , , , xlYes, , , xlTopToBottom
    Range("g1").Resize(d1.Count + 1, d2.Count).Sort [g1], xlAscending, , , , , , , , , xlLeftToRight
End Sub
```

## 方法二：

```vb
Sub 方法二()
    Dim i, j, k
    Dim ar, br()
    Dim d As Object, kw$
    Set d = CreateObject("Scripting.Dictionary")
    'd.CompareMode = vbTextCompare '不区分大小写
    ar = Range("a1:c" & [a65536].End(3).Row)

    ReDim br(1 To UBound(ar), 1 To 1000)
    Dim rowNum, colNum
    rowNum = 1: colNum = 1
    For i = 2 To UBound(ar)
        '\\型号
        If Not d.exists(ar(i, 2)) Then
            colNum = colNum + 1
            br(1, colNum) = ar(i, 2)
            d(ar(i, 2)) = colNum
        End If

        '\\名称
        If Not d.exists(ar(i, 1)) Then
            rowNum = rowNum + 1
            br(rowNum, 1) = ar(i, 1)
            d(ar(i, 1)) = rowNum
            br(rowNum, d(ar(i, 2))) = ar(i, 3)
        Else
            br(d(ar(i, 1)), d(ar(i, 2))) = br(d(ar(i, 1)), d(ar(i, 2))) + ar(i, 3)
        End If

    Next i

    [f1].Resize(rowNum, colNum) = br
    Range("f1").Resize(rowNum, colNum).Sort [f1], xlAscending, , , , , , xlYes, , , xlTopToBottom
    Range("g1").Resize(rowNum, colNum - 1).Sort [g1], xlAscending, , , , , , , , , xlLeftToRight
End Sub
```

## 方法三：

```vb
Sub SQL查询()
    '定义变量
    Dim cnn, rst, SQL$
    Dim i, j, k
    Set cnn = CreateObject("adodb.connection") '创建数据库连接
    Set rst = CreateObject("adodb.recordset") '创建一个数据集保存数据

    '设置数据库连接
    If Val(Application.Version) < 12 Then
        cnn.Open "Provider=Microsoft.Jet.Oledb.4.0;Extended Properties='Excel 8.0;HDR=yes';Data Source=" & ThisWorkbook.FullName
    Else
        cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties='Excel 12.0;HDR=yes';Data Source=" & ThisWorkbook.FullName
    End If

    '设置SQL语句
    SQL = "TRANSFORM SUM(数量) SELECT 名称 from [Sheet1$a1:c18] GROUP BY 名称 PIVOT 型号" '

    'SQL结果处理
    Set rst = cnn.Execute(SQL)

    Range("f2").CopyFromRecordset rst
    For i = 1 To rst.Fields.Count
        Cells(1, i + 5) = rst.Fields(i - 1).Name
    Next

    rst.Close
    cnn.Close '关闭数据库连接
    Set rst = Nothing
    Set cnn = Nothing '将cnn从内存中删除
End Sub
```

[原始链接](https://mp.weixin.qq.com/s?__biz=MzIyOTc3NzQ2NA==&mid=2247485136&idx=1&sn=fcc6dd5ad8e974e4b31bea47b811c8b6&chksm=e8bccf87dfcb46912f89badebc8df22ecbec9bc0390633120dbb78e2e28392a310aa9d5401fe&scene=178&cur_album_id=3115603487041503237#rd)

---

> 作者: [richfan](https://richfan.site/)  
> URL: http://richfan.site/%E7%BC%96%E7%A8%8B/vba/vba%E6%A1%88%E5%88%97/vba%E6%A1%88%E4%BE%8B007%E5%A4%9A%E6%9D%A1%E4%BB%B6%E6%B1%87%E6%80%BB/  

