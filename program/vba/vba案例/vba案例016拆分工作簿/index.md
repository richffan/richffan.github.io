# 【VBA案例016】拆分工作簿

大家好！

通过本次和之前几次的分享。对工作表、工作簿的拆分、合并都做了最基础的介绍以及案例演示。

而面对实际工作中经常遇到的工作表、工作簿的合并拆分问题，相信大家通过对这几个案例的交汇融合，基本上都可以找到对应的解决方案。

今天分享的是拆分系列的最有一项内容：拆分工作簿。就是把一个工作簿中的所有Sheet拆分成一个一个的独立的工作簿。

![](https://img.richfan.site/program/vba/vba案列/【VBA案例016】拆分工作簿_1.gif)

通过今天的案例，你将掌握工作表对象的Move和Copy方法。听起来很陌生，其实是你经常在用的一个很普通的功能：

![](https://img.richfan.site/program/vba/vba案列/【VBA案例016】拆分工作簿_2.png)

语法：
   Sheet.Copy([Before],[After])'参数可选，非必须
   Sheet.Move([Before],[After])'参数可选，非必须
注意：
   如果不填参数，将会把Sheet移动或复制到一个新的工作簿中
以下是VBA代码。详细解析请看文末的视频。

```vb
Option Explicit

Sub 拆分工作簿()

    Dim sht As Worksheet
    Dim filePath

    filePath = ThisWorkbook.Path & "\"

    Application.ScreenUpdating = False
    For Each sht In ThisWorkbook.Worksheets
        If sht.Name <> "首页" Then
            'sht.Move
            sht.Copy
            ActiveWorkbook.SaveAs filePath & sht.Name & ".xlsx"
            ActiveWorkbook.Close
        End If
    Next
    Application.ScreenUpdating = True

End Sub
```

[原始链接](https://mp.weixin.qq.com/s?__biz=MzIyOTc3NzQ2NA==&mid=2247485257&idx=1&sn=1ab601287ccfb83c25899c5cbc734433&chksm=e8bcce1edfcb4708aefdfa5f75ae9cf5f1a6a37466cf0dbeb75dc10bf286c2ba57a3e10ddf38&scene=178&cur_album_id=3115603487041503237#rd)

---

> 作者: [richfan](https://richfan.site/)  
> URL: http://richfan.site/program/vba/vba%E6%A1%88%E4%BE%8B/vba%E6%A1%88%E4%BE%8B016%E6%8B%86%E5%88%86%E5%B7%A5%E4%BD%9C%E7%B0%BF/  

