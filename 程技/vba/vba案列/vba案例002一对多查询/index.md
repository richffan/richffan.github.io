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
    Sheets(&#34;查询&#34;).Range(&#34;a6:d65536&#34;).ClearContents
    Dim str As String
    str = Sheets(&#34;查询&#34;).Range(&#34;b3&#34;)

    k = 5
    With Sheets(&#34;数据源&#34;)
        For Each cel In .Range(&#34;a2:d&#34; &amp; .[a65536].End(3).Row) &#39;xlup
            If cel.Value = str Then
                k = k &#43; 1
                For j = 1 To 4
                    Sheets(&#34;查询&#34;).Cells(k, j) = cel.Offset(0, j - 1)
                Next
            End If
        Next
    End With
    MsgBox Format(Timer - t, &#34;0.000s&#34;)
End Sub
```

## 方法二：

```vb
Sub 数组循环()
    Dim i, j, k, irow
    Dim cel As Range
    Dim t As Double
    t = Timer
    Sheets(&#34;查询&#34;).Range(&#34;a6:d65536&#34;).ClearContents
    Dim str As String
    str = Sheets(&#34;查询&#34;).Range(&#34;b3&#34;)

    Dim ar, br() &#39;ar 数据源 ,br 结果
    With Sheets(&#34;数据源&#34;)
        irow = .[a65536].End(3).Row
        ar = .Range(&#34;a2:d&#34; &amp; irow)
    End With

    ReDim br(1 To UBound(ar), 1 To UBound(ar, 2))
    For i = 1 To UBound(ar)
        If ar(i, 1) = str Then
            k = k &#43; 1
            For j = 1 To UBound(br, 2)
                br(k, j) = ar(i, j)
            Next j
        End If
    Next i

    Sheets(&#34;查询&#34;).Range(&#34;a6&#34;).Resize(k, UBound(br, 2)) = br
    MsgBox Format(Timer - t, &#34;0.000s&#34;)
End Sub
```

## 方法三：

```vb
Sub 字典查询()
    Dim i, j, k, irow
    Dim cel As Range
    Dim t As Double
    t = Timer
    Sheets(&#34;查询&#34;).Range(&#34;a6:d65536&#34;).ClearContents
    Dim str As String
    str = Sheets(&#34;查询&#34;).Range(&#34;b3&#34;)

    Dim ar, br() &#39;ar 数据源 ,br 结果
    With Sheets(&#34;数据源&#34;)
        irow = .[a65536].End(3).Row
        ar = .Range(&#34;a2:d&#34; &amp; irow)
    End With

    Dim d As Object, kw$
    Set d = CreateObject(&#34;Scripting.Dictionary&#34;)
    &#39;d.CompareMode = vbTextCompare &#39;不区分大小写

    For i = 1 To UBound(ar)
        If Not d.exists(ar(i, 1)) Then
            d(ar(i, 1)) = i
        Else
            d(ar(i, 1)) = d(ar(i, 1)) &amp; &#34;,&#34; &amp; i
        End If
    Next i

    Dim tmpAr
    tmpAr = Split(d(str), &#34;,&#34;)

    ReDim br(1 To UBound(tmpAr) &#43; 1, 1 To UBound(ar, 2))
    For i = 0 To UBound(tmpAr)
        For j = 1 To UBound(ar, 2)
            br(i &#43; 1, j) = ar(tmpAr(i), j)
        Next j
    Next i

    Sheets(&#34;查询&#34;).Range(&#34;a6&#34;).Resize(UBound(br), UBound(br, 2)) = br
    MsgBox Format(Timer - t, &#34;0.000s&#34;)
End Sub
```

## 方法四：

```vb
Sub SQL查询()
    &#39;定义变量
    Dim cnn, rst, SQL$
    Dim i, j, k
    Set cnn = CreateObject(&#34;adodb.connection&#34;) &#39;创建数据库连接
    Set rst = CreateObject(&#34;adodb.recordset&#34;) &#39;创建一个数据集保存数据
    Dim t As Double
    t = Timer
    &#39;设置数据库连接
    If Val(Application.Version) &lt; 12 Then
        cnn.Open &#34;Provider=Microsoft.Jet.Oledb.4.0;Extended Properties=&#39;Excel 8.0;HDR=yes&#39;;Data Source=&#34; &amp; ThisWorkbook.FullName
    Else
        cnn.Open &#34;Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties=&#39;Excel 12.0;HDR=yes&#39;;Data Source=&#34; &amp; ThisWorkbook.FullName
    End If

    &#39;设置SQL语句
    SQL = &#34;SELECT * FROM [数据源$a1:d100] WHERE 班级=&#39;&#34; &amp; Sheets(&#34;查询&#34;).[B3] &amp; &#34;&#39;&#34;

    &#39;SQL结果处理
    Set rst = cnn.Execute(SQL)

    Sheets(&#34;查询&#34;).Range(&#34;a6:d65536&#34;).ClearContents &#39;清理保存数据的区域
    Sheets(&#34;查询&#34;).Range(&#34;a6&#34;).CopyFromRecordset rst &#39;结果输出（不带表头）
    MsgBox Format(Timer - t, &#34;0.000s&#34;)
    rst.Close
    cnn.Close &#39;关闭数据库连接
    Set rst = Nothing
    Set cnn = Nothing &#39;将cnn从内存中删除
End Sub
```

[原始链接](https://mp.weixin.qq.com/s?__biz=MzIyOTc3NzQ2NA==&amp;mid=2247485016&amp;idx=1&amp;sn=03ab36893d1e13a10e605f572ca30038&amp;chksm=e8bccf0fdfcb46192da110df69748b7fa65fb2cea568055237e818074b0e076fb0da12aff279&amp;scene=178&amp;cur_album_id=3115603487041503237#rd)

---

> 作者:   
> URL: http://richfan.site/%E7%A8%8B%E6%8A%80/vba/vba%E6%A1%88%E5%88%97/vba%E6%A1%88%E4%BE%8B002%E4%B8%80%E5%AF%B9%E5%A4%9A%E6%9F%A5%E8%AF%A2/  

