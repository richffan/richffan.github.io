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
    Sheets("查询").Range("a6:d65536").ClearContents
    Dim str As String
    str = Sheets("查询").Range("b3")

    k = 5
    With Sheets("数据源")
        For Each cel In .Range("a2:d" & .[a65536].End(3).Row) 'xlup
            'If cel.Value = str Then
            'If InStr(cel.Value, str) > 0 Then
            If cel.Value Like "*" & str & "*" Then
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
        'If ar(i, 1) = str Then
        'If InStr(ar(i, 1), str) > 0 Then
        If ar(i, 1) Like "*" & str & "*" Then
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

### 方法三：

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
        'If InStr(ar(i, 1), str) > 0 Then
        If ar(i, 1) Like "*" & str & "*" Then
            If Not d.exists(str) Then
                d(str) = i
            Else
                d(str) = d(str) & "," & i
            End If
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
    'SQL = "SELECT * FROM [数据源$a1:d100] WHERE 班级='" & Sheets("查询").[B3] & "'"
    'SQL = "SELECT * FROM [数据源$a1:d100] WHERE instr(班级,'" & Sheets("查询").[B3] & "')>0"
    SQL = "SELECT * FROM [数据源$a1:d100] WHERE 班级 like '" & Sheets("查询").[B3] & "%'"
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

[原始链接](https://mp.weixin.qq.com/s?__biz=MzIyOTc3NzQ2NA==&mid=2247485027&idx=1&sn=e54ce7d71355280dcb3e863977e5bb76&chksm=e8bccf34dfcb46228672ac50507c236d2f5e72b8e7867df75e5ce6a41cc2e2031aa8673563f8&scene=178&cur_album_id=3115603487041503237#rd)

---

> 作者: [richfan](https://richfan.site/)  
> URL: http://richfan.site/program/vba/vba%E6%A1%88%E4%BE%8B/vba%E6%A1%88%E4%BE%8B003%E6%A8%A1%E7%B3%8A%E6%9F%A5%E8%AF%A2/  

