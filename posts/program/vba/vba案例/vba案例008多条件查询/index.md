# 【VBA案例008】多条件查询

大家好！今天分享的案例是多条件查询。

这个查询在进销存或者库存管理中特别常用，如果你准备或者正在做一个自己的管理查询工具，这个方法一定要会。先来看一下数据。

比如说，现在有一份产品信息表。

![](https://img.richfan.site/program/vba/vba案列/【VBA案例008】多条件查询_1.jpg)

我们要做的是在查询页面，输入参数后，查询出所有满足条件的内容。其中参数可以不填，不填就表示要查询所有内容。

![](https://img.richfan.site/program/vba/vba案列/【VBA案例008】多条件查询_2.gif)

本期将使用2个方法来实现效果，以下是VBA代码，详细视频解析在文末。

## 方法一：

```vb
Sub 方法一()

    Dim i, j, k
    Dim ar, br()

    With Sheet2
        ar = .Range("a1:f" & .[a65536].End(3).Row)
    End With

    ReDim br(1 To UBound(ar), 1 To UBound(ar, 2))

    Dim customer, product, startDate As Date, endDate As Date
    With Sheet1
        customer = IIf(.[b2] = "", "", "," & .[b2] & ",")
        product = IIf(.[d2] = "", "", "," & .[d2] & ",")
        startDate = IIf(.[f2] = "", #1/1/1900#, .[f2])
        endDate = IIf(.[h2] = "", #1/1/2400#, .[h2])
    End With

    For i = 2 To UBound(ar)
        If ar(i, 1) >= startDate And ar(i, 1) <= endDate Then
            If InStr("," & ar(i, 2) & ",", customer) > 0 Then
                If InStr("," & ar(i, 3) & ",", product) > 0 Then
                    k = k + 1
                    For j = 1 To UBound(br, 2)
                        br(k, j) = ar(i, j)
                    Next j
                End If
            End If
        End If
    Next i

    Sheet1.[a5:f65536].ClearContents
    If k > 0 Then
        Sheet1.[a5].Resize(k, UBound(br, 2)) = br
    End If

End Sub
```

## 方法二：

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

    Dim customer, product, startDate As Date, endDate As Date
    With Sheet1
        customer = IIf(.[b2] = "", "", "," & .[b2] & ",")
        product = IIf(.[d2] = "", "", "," & .[d2] & ",")
        startDate = IIf(.[f2] = "", #1/1/1900#, .[f2])
        endDate = IIf(.[h2] = "", #1/1/2400#, .[h2])
    End With

    '设置SQL语句
    SQL = "select * from [产品信息$A1:F22] where 日期>=#" & startDate & "# and 日期<=#" & endDate & "# and instr(','&客户&',','" & customer & "')>0 and instr(','&产品&',','" & product & "')>0" '

    'SQL结果处理
    Set rst = cnn.Execute(SQL)

    Sheet1.[a5:f65536].ClearContents
    Sheet1.Range("a5").CopyFromRecordset rst
'    For i = 1 To rst.Fields.Count
'        Cells(1, i + 5) = rst.Fields(i - 1).Name
'    Next

    rst.Close
    cnn.Close '关闭数据库连接
    Set rst = Nothing
    Set cnn = Nothing '将cnn从内存中删除
End Sub
```

[原始链接](https://mp.weixin.qq.com/s?__biz=MzIyOTc3NzQ2NA==&mid=2247485165&idx=1&sn=db0537ef59d97d88e01ce1fffdb64ede&chksm=e8bccfbadfcb46acfd719e6ecdd36faf6e914abf46752426d49e7de49ac627a31036befdb81b&scene=178&cur_album_id=3115603487041503237#rd)

---

> 作者: [richfan](https://richfan.site/)  
> URL: http://richfan.site/posts/program/vba/vba%E6%A1%88%E4%BE%8B/vba%E6%A1%88%E4%BE%8B008%E5%A4%9A%E6%9D%A1%E4%BB%B6%E6%9F%A5%E8%AF%A2/  

