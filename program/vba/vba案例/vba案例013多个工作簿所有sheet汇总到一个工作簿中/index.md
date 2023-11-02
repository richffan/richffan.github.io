# 【VBA案例013】多个工作簿所有Sheet汇总到一个工作簿中

大家好！昨天视频放错了，今天重新发一下。图片

这次分享的是合并系列的最后一个案例：汇总工作簿。

打个比方：把多个工作簿中的每个Sheet，汇总到一个工作簿里，汇总完之后是所有的Sheet都在同一个工作簿里。

![](https://img.richfan.site/program/vba/vba案列/【VBA案例013】多个工作簿所有Sheet汇总到一个工作簿中.gif)

这次使用了Dir函数，以下是VBA代码。详细解析请看文末的视频。

'Dir函数：返回的是指定路径下【文件】或者【文件夹】的名称。如果不存在，就返回 "" 字符串
'Dir 函数一般搭配 do while 循环遍历文件，结束的条件就是Dir返回空值。

'举例：
'fileName1 = Dir("E:\test\*.xlsx")  查找E盘test文件夹里边的xlsx工作簿，将第一个工作簿名称返回
'fileName2 = Dir                    查找下一个，并且不需要写参数
'fileName3 = Dir                    查找下一个，并且不需要写参数
'fileName4 = Dir                    查找下一个，并且不需要写参数
'......                           ......
'当返回 "" 的时候说明查找完了。如果直接 fileName_N = Dir 就会报错，需要重新指定参数

```vb
Sub 汇总工作簿()

    Dim filePath, fileName
    Dim sht As Worksheet

    filePath = ThisWorkbook.Path & "\文件夹\"

    fileName = Dir(filePath & "*.xlsx")

    Application.ScreenUpdating = False
    Do While fileName <> ""
        With Workbooks.Open(filePath & fileName)
            For Each sht In .Worksheets
                sht.Copy after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
                ActiveSheet.Name = sht.Name & "-" & Replace(.Name, ".xlsx", "")
            Next
            .Close False
        End With
        fileName = Dir
    Loop
    Application.ScreenUpdating = True

End Sub
```

[原始链接](https://mp.weixin.qq.com/s?__biz=MzIyOTc3NzQ2NA==&mid=2247485233&idx=1&sn=a27892d151b9a8332f01f20be6e0f7f9&chksm=e8bcce66dfcb4770c956e41db7a9d08d9fc1e153cd3cf7681b0130d12be97b7908d27d53a0e8&scene=178&cur_album_id=3115603487041503237#rd)

---

> 作者: [richfan](https://richfan.site/)  
> URL: http://richfan.site/program/vba/vba%E6%A1%88%E4%BE%8B/vba%E6%A1%88%E4%BE%8B013%E5%A4%9A%E4%B8%AA%E5%B7%A5%E4%BD%9C%E7%B0%BF%E6%89%80%E6%9C%89sheet%E6%B1%87%E6%80%BB%E5%88%B0%E4%B8%80%E4%B8%AA%E5%B7%A5%E4%BD%9C%E7%B0%BF%E4%B8%AD/  

