# 【VBA案例020】整合工作簿


大家好！今天回答一位粉丝朋友的提问。

问题是：将多个工作簿中的所有工作表合并汇总，要求名称相同的工作表内容要合并在一起，名称不同的要单独作为一个工作表。

为此，我模拟了一份数据，结构如下图：

![](https://img.richfan.site/program/vba/vba案列/【VBA案例020】整合工作簿.png)

这个问题，其实是我之前分享的【案例011合并工作表】和【案例013汇总工作簿】的融合版。方法非常的相似。其实对于工作簿和工作表的合并与拆分的操作，之前的案例基本都分享完了。只要融会贯通，举一反三，相信这种问题将迎刃而解。

效果就不演示了，以下是VBA代码。详细解析请看文末的视频。

```vb
Option Explicit

Sub 汇总合并工作簿()
    Dim shtName
    Dim sht As Worksheet

    For Each sht In ThisWorkbook.Worksheets
        shtName = shtName & "," & sht.Name
    Next

    Dim filePath$, fileName As String

    filePath = ThisWorkbook.Path & "\文件夹\"
    fileName = Dir(filePath & "*.xlsx")

    Dim row_count, thisRow_count
    Application.ScreenUpdating = False
    Do While fileName <> ""
        With Workbooks.Open(filePath & fileName)
            For Each sht In .Worksheets
                If InStr("," & shtName & ",", "," & sht.Name & ",") > 0 Then
                    row_count = sht.[a65536].End(3).Row
                    thisRow_count = ThisWorkbook.Worksheets(sht.Name).[a65536].End(3).Row
                    sht.Range("a2:e" & row_count).Copy ThisWorkbook.Worksheets(sht.Name).Range("a" & thisRow_count + 1)
                Else
                    sht.Copy after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
                    shtName = shtName & "," & sht.Name
                End If
            Next
            .Close False
        End With
        fileName = Dir
    Loop
    Application.ScreenUpdating = True

End Sub
```

[原始链接](https://mp.weixin.qq.com/s?__biz=MzIyOTc3NzQ2NA==&mid=2247485296&idx=1&sn=368ced654f9b46912baa0fba537656af&chksm=e8bcce27dfcb4731f5192f230ed000202ad3c401136f1d56d70c53a901bf99ab6d377d416eaf&scene=178&cur_album_id=3115603487041503237#rd)

---

> 作者: [richfan](https://richfan.site/)  
> URL: http://richfan.site/%E7%A8%8B%E6%8A%80/vba/vba%E6%A1%88%E5%88%97/vba%E6%A1%88%E4%BE%8B020%E6%95%B4%E5%90%88%E5%B7%A5%E4%BD%9C%E7%B0%BF/  

