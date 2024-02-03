# 【VBA案例003】模糊查询

大家好，模糊查询，在平时工作中会经常遇到。

本期呢，会将模糊查询的两个最常用的方法分享给大家。

1、instr函数
2、like运算符

在开始之前，我建议先把上一期的【VBA案例002】一对多查询看一遍。因为本次内容是基于上次的文件衍生出来的。并且代码也大同小异。

先附上VBA代码，详细内容，请看文章最后的视频哦！~

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
            &#39;If cel.Value = str Then
            &#39;If InStr(cel.Value, str) &gt; 0 Then
            If cel.Value Like &#34;*&#34; &amp; str &amp; &#34;*&#34; Then
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
        &#39;If ar(i, 1) = str Then
        &#39;If InStr(ar(i, 1), str) &gt; 0 Then
        If ar(i, 1) Like &#34;*&#34; &amp; str &amp; &#34;*&#34; Then
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

### 方法三：

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
        &#39;If InStr(ar(i, 1), str) &gt; 0 Then
        If ar(i, 1) Like &#34;*&#34; &amp; str &amp; &#34;*&#34; Then
            If Not d.exists(str) Then
                d(str) = i
            Else
                d(str) = d(str) &amp; &#34;,&#34; &amp; i
            End If
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
    &#39;SQL = &#34;SELECT * FROM [数据源$a1:d100] WHERE 班级=&#39;&#34; &amp; Sheets(&#34;查询&#34;).[B3] &amp; &#34;&#39;&#34;
    &#39;SQL = &#34;SELECT * FROM [数据源$a1:d100] WHERE instr(班级,&#39;&#34; &amp; Sheets(&#34;查询&#34;).[B3] &amp; &#34;&#39;)&gt;0&#34;
    SQL = &#34;SELECT * FROM [数据源$a1:d100] WHERE 班级 like &#39;&#34; &amp; Sheets(&#34;查询&#34;).[B3] &amp; &#34;%&#39;&#34;
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

[原始链接](https://mp.weixin.qq.com/s?__biz=MzIyOTc3NzQ2NA==&amp;mid=2247485027&amp;idx=1&amp;sn=e54ce7d71355280dcb3e863977e5bb76&amp;chksm=e8bccf34dfcb46228672ac50507c236d2f5e72b8e7867df75e5ce6a41cc2e2031aa8673563f8&amp;scene=178&amp;cur_album_id=3115603487041503237#rd)

---

> 作者:   
> URL: http://richfan.site/%E7%A8%8B%E6%8A%80/vba/vba%E6%A1%88%E5%88%97/vba%E6%A1%88%E4%BE%8B003%E6%A8%A1%E7%B3%8A%E6%9F%A5%E8%AF%A2/  

