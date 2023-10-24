# 【VBA案例002】一对多查询

一对多查询的方法有很多，这里附上VBA代码，详细过程请看文章最后的视频。

![](https://img.richfan.site/program/vba/vba案列/【VBA案例002】一对多查询_1.png)
![](https://img.richfan.site/program/vba/vba案列/【VBA案例002】一对多查询_2.png)

## 方法一：

```vb
Sub 单元格循环()
    Dim i, j, k, irow
    Dim cel As Range
    Dim t As Double
    t = Timer
    Sheets("查询").Range("a6:d65536").ClearContents
    Dim str As String
    str = Sheets("查询").Range("b3")

    k = 5
    With Sheets("数据源")
        For Each cel In .Range("a2:d" & .[a65536].End(3).Row) 'xlup
            If cel.Value = str Then
                k = k + 1
                For j = 1 To 4
                    Sheets("查询").Cells(k, j) = cel.Offset(0, j - 1)
                Next
            End If
        Next
    End With
    MsgBox Format(Timer - t, "0.000s")
End Sub
```

## 方法二：

```vb
Sub 数组循环()
    Dim i, j, k, irow
    Dim cel As Range
    Dim t As Double
    t = Timer
    Sheets("查询").Range("a6:d65536").ClearContents
    Dim str As String
    str = Sheets("查询").Range("b3")

    Dim ar, br() 'ar 数据源 ,br 结果
    With Sheets("数据源")
        irow = .[a65536].End(3).Row
        ar = .Range("a2:d" & irow)
    End With

    ReDim br(1 To UBound(ar), 1 To UBound(ar, 2))
    For i = 1 To UBound(ar)
        If ar(i, 1) = str Then
            k = k + 1
            For j = 1 To UBound(br, 2)
                br(k, j) = ar(i, j)
            Next j
        End If
    Next i

    Sheets("查询").Range("a6").Resize(k, UBound(br, 2)) = br
    MsgBox Format(Timer - t, "0.000s")
End Sub
```

## 方法三：

```vb
Sub 字典查询()
    Dim i, j, k, irow
    Dim cel As Range
    Dim t As Double
    t = Timer
    Sheets("查询").Range("a6:d65536").ClearContents
    Dim str As String
    str = Sheets("查询").Range("b3")

    Dim ar, br() 'ar 数据源 ,br 结果
    With Sheets("数据源")
        irow = .[a65536].End(3).Row
        ar = .Range("a2:d" & irow)
    End With

    Dim d As Object, kw$
    Set d = CreateObject("Scripting.Dictionary")
    'd.CompareMode = vbTextCompare '不区分大小写

    For i = 1 To UBound(ar)
        If Not d.exists(ar(i, 1)) Then
            d(ar(i, 1)) = i
        Else
            d(ar(i, 1)) = d(ar(i, 1)) & "," & i
        End If
    Next i

    Dim tmpAr
    tmpAr = Split(d(str), ",")

    ReDim br(1 To UBound(tmpAr) + 1, 1 To UBound(ar, 2))
    For i = 0 To UBound(tmpAr)
        For j = 1 To UBound(ar, 2)
            br(i + 1, j) = ar(tmpAr(i), j)
        Next j
    Next i

    Sheets("查询").Range("a6").Resize(UBound(br), UBound(br, 2)) = br
    MsgBox Format(Timer - t, "0.000s")
End Sub
```

## 方法四：

```vb
Sub SQL查询()
    '定义变量
    Dim cnn, rst, SQL$
    Dim i, j, k
    Set cnn = CreateObject("adodb.connection") '创建数据库连接
    Set rst = CreateObject("adodb.recordset") '创建一个数据集保存数据
    Dim t As Double
    t = Timer
    '设置数据库连接
    If Val(Application.Version) < 12 Then
        cnn.Open "Provider=Microsoft.Jet.Oledb.4.0;Extended Properties='Excel 8.0;HDR=yes';Data Source=" & ThisWorkbook.FullName
    Else
        cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties='Excel 12.0;HDR=yes';Data Source=" & ThisWorkbook.FullName
    End If

    '设置SQL语句
    SQL = "SELECT * FROM [数据源$a1:d100] WHERE 班级='" & Sheets("查询").[B3] & "'"

    'SQL结果处理
    Set rst = cnn.Execute(SQL)

    Sheets("查询").Range("a6:d65536").ClearContents '清理保存数据的区域
    Sheets("查询").Range("a6").CopyFromRecordset rst '结果输出（不带表头）
    MsgBox Format(Timer - t, "0.000s")
    rst.Close
    cnn.Close '关闭数据库连接
    Set rst = Nothing
    Set cnn = Nothing '将cnn从内存中删除
End Sub
```

[原始链接](https://mp.weixin.qq.com/s?__biz=MzIyOTc3NzQ2NA==&mid=2247485016&idx=1&sn=03ab36893d1e13a10e605f572ca30038&chksm=e8bccf0fdfcb46192da110df69748b7fa65fb2cea568055237e818074b0e076fb0da12aff279&scene=178&cur_album_id=3115603487041503237#rd)

---

> 作者: [richfan](https://richfan.site/)  
> URL: http://richfan.site/posts/program/vba/vba%E6%A1%88%E4%BE%8B/vba%E6%A1%88%E4%BE%8B002%E4%B8%80%E5%AF%B9%E5%A4%9A%E6%9F%A5%E8%AF%A2/  

