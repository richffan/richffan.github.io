# 【VBA案例004】自动填充表单


大家好！今天我们分享一个有关VBA自动填充表单的案例，帮助大家简化表单填写的过程，提高工作效率。

我们经常需要填写大量重复的表单，例如报销单、数据输入表格等，并且每个表单都有各自的字段需要填写。手动逐个填写这些表单既耗时又容易出错。

使用VBA自动填充表单可以极大地简化这一过程，让我们能够更专注于其他重要的工作。

譬如说，现在有两个表格，一个是人员清单，另一个的表单。

![](https://img.richfan.site/program/vba/vba案列/【VBA案例004】自动填充表单_1.png)

现在要做的是，将人员清单中的每一行信息，填到表单里，并导出为一个单独的工作簿。

![](https://img.richfan.site/program/vba/vba案列/【VBA案例004】自动填充表单_2.png)

下面是VBA的代码分享，你也可以观看下方的视频解析：

## 方法一：

```vb
Sub 表格填充()

    Dim i, j, k

    Dim xinXiSht As Worksheet
    Dim qingDanSht As Worksheet

    Set xinXiSht = ThisWorkbook.Worksheets("信息卡")
    Set qingDanSht = ThisWorkbook.Worksheets("人员清单")

    With qingDanSht
        For i = 2 To .Range("a" & .Rows.Count).End(xlUp).Row
            For j = 1 To 5
                xinXiSht.Range("c" & j + 4) = .Cells(i, j)
            Next j

            xinXiSht.Copy
            ActiveWorkbook.SaveAs ThisWorkbook.Path & "\" & .Range("a" & i) & ".xlsx"
            ActiveWorkbook.Close True
        Next i
    End With

End Sub
```

## 方法二：

```vb
Sub 表格填充优化()

    Dim i, j, k

    Dim xinXiSht As Worksheet
    Dim qingDanSht As Worksheet

    Set xinXiSht = ThisWorkbook.Worksheets("信息卡")
    Set qingDanSht = ThisWorkbook.Worksheets("人员清单")

    Application.ScreenUpdating = False

    Dim ar, br()
    With qingDanSht
        ar = .Range("a2:e" & .Range("a" & .Rows.Count).End(xlUp).Row)
        For i = 1 To UBound(ar)
            ReDim br(1 To 5, 1 To 1)
            For j = 1 To 5
                br(j, 1) = ar(i, j)
            Next j
            xinXiSht.[c5].Resize(UBound(br)) = br
            xinXiSht.Copy
            ActiveWorkbook.SaveAs ThisWorkbook.Path & "\" & ar(i, 1) & ".xlsx"
            ActiveWorkbook.Close True
            xinXiSht.[c5:c9].ClearContents
        Next i
    End With

    Application.ScreenUpdating = True
End Sub
```

[原始链接](https://mp.weixin.qq.com/s?__biz=MzIyOTc3NzQ2NA==&mid=2247485040&idx=1&sn=1f89ff1b2c14a4b22607fcc1ba725ce9&chksm=e8bccf27dfcb463130a13e9e02c0e0534f314ffbfc12cf3697a006cb2b215ebc8db0e1f0c3d5&scene=178&cur_album_id=3115603487041503237#rd)

---

> 作者: [richfan](https://richfan.site/)  
> URL: http://richfan.site/%E7%A8%8B%E6%8A%80/vba/vba%E6%A1%88%E5%88%97/vba%E6%A1%88%E4%BE%8B004%E8%87%AA%E5%8A%A8%E5%A1%AB%E5%85%85%E8%A1%A8%E5%8D%95/  

