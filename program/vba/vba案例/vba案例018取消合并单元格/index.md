# 【VBA案例018】取消合并单元格

大家好！

上一篇内容分享了，用VBA合并单元格的两个方法。本期就来看一下如何取消合并的单元格。

![](https://img.richfan.site/program/vba/vba案列/【VBA案例018】取消合并单元格.gif)

同样的，WPS内置了一键取消合并单元格并填充内容的功能，而Excel用户只能流下羡慕的泪水。

这次介绍两个方法，来实现和WPS一模一样的功能。

第一个方法：使用录制宏的操作，把平时手动取消合并的操作记录下来，并对代码加以修改，来达到目的。第二个方法：使用的循环单元格判断是否合并的方法，其中有几个常用的属性是我们必须要了解掌握的。

两个方法都比较简单。大家可以根据实际情况，酌情选择适用的方法。

以下是VBA代码。详细解析请看文末的视频。

## 方法一：

```vb
Sub 宏2()

    Range("a1:a20").UnMerge
    Range("a1:a20").SpecialCells(xlCellTypeBlanks).FormulaR1C1 = "=R[-1]C"
    Range("a1:a20").Value = Range("a1:a20").Value

End Sub
```

## 方法二：

```vb
Sub 取消合并单元格()

    Dim cel As Range
    Dim mRng As Range
    Dim mRngVal

    For Each cel In Range("a1:a20")
        If cel.MergeCells Then
            Set mRng = cel.MergeArea
            mRngVal = mRng(1, 1)
            mRng.UnMerge
            mRng.Value = mRngVal
        End If
    Next

End Sub
```

[原始链接](https://mp.weixin.qq.com/s?__biz=MzIyOTc3NzQ2NA==&mid=2247485272&idx=1&sn=898b64ffdb77aff339d71f7bda45906f&chksm=e8bcce0fdfcb47192a223e90ac62cd0b104af662e8f66c26dfdc87801f61712a960f6777470c&scene=178&cur_album_id=3115603487041503237#rd)

---

> 作者: [richfan](https://richfan.site/)  
> URL: http://richfan.site/program/vba/vba%E6%A1%88%E4%BE%8B/vba%E6%A1%88%E4%BE%8B018%E5%8F%96%E6%B6%88%E5%90%88%E5%B9%B6%E5%8D%95%E5%85%83%E6%A0%BC/  

