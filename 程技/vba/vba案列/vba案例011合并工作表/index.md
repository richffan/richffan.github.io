# 【VBA案例011】合并工作表

大家好！终于放假了，知了祝大家双节快乐！~

今天分享的案例非常的经典，它让我首次见识了VBA的魅力。也是从那会儿开始，命运的齿轮开始转动，走上了学习VBA的道路。

它就是：合并工作表。

问题很简单：就是把一个工作簿里的所有sheet合并到一个总表里边。

如下图：将 4个班级的信息合并。

![](https://img.richfan.site/program/vba/vba案列/【VBA案例011】合并工作表.gif)

以下是VBA代码（长按复制，可向右托动），详细解析请看文末视频。

## 方法一：复制粘贴

```vb
Sub 复制粘贴()

    Dim sht As Worksheet
    Dim zongSht As Worksheet

    Set zongSht = Worksheets("总表")

    Dim rowsCount, zongRows
    For Each sht In Worksheets
        If sht.Name <> "总表" Then
            With sht
                rowsCount = .[a65536].End(3).Row
                zongRows = zongSht.[a65536].End(3).Row + 1
                .Range("a2:d" & rowsCount).Copy zongSht.Range("a" & zongRows)
            End With
        End If
    Next

End Sub
```

## 方法二：数组

```vb
Sub 数组()

    Dim sht As Worksheet
    Dim zongSht As Worksheet

    Set zongSht = Worksheets("总表")

    Dim rowsCount, zongRows
    Dim tmpAr
    For Each sht In Worksheets
        If sht.Name <> "总表" Then
            With sht
                rowsCount = .[a65536].End(3).Row
                zongRows = zongSht.[a65536].End(3).Row + 1
                tmpAr = .Range("a2:d" & rowsCount)
                zongSht.Range("a" & zongRows).Resize(UBound(tmpAr), UBound(tmpAr, 2)) = tmpAr
            End With
        End If
    Next

End Sub
```

[原始链接](https://mp.weixin.qq.com/s?__biz=MzIyOTc3NzQ2NA==&mid=2247485203&idx=1&sn=c64987b006f9b38f49d0878e102b01bc&chksm=e8bcce44dfcb4752782767f464205bccc84b743f39e8a6f8988ae79067c61e85130dbc91aa55&scene=178&cur_album_id=3115603487041503237#rd)

---

> 作者: [richfan](https://richfan.site/)  
> URL: http://richfan.site/%E7%A8%8B%E6%8A%80/vba/vba%E6%A1%88%E5%88%97/vba%E6%A1%88%E4%BE%8B011%E5%90%88%E5%B9%B6%E5%B7%A5%E4%BD%9C%E8%A1%A8/  

