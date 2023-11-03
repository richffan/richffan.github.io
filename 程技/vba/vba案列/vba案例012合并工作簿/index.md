# 【VBA案例012】合并工作簿

大家好！这次分享的是非常经典的案例：合并工作簿。

相信大家已经很熟悉这个问题了，就是把多个工作簿里的工作表合并到同一个sheet里。

![](https://img.richfan.site/program/vba/vba案列/【VBA案例012】合并工作簿_1.png)

![](https://img.richfan.site/program/vba/vba案列/【VBA案例012】合并工作簿_2.gif)

这次同样分享两个方法，以下是VBA代码。详细解析请看文末的视频。

自定义函数：

```vb
Private Function filelist(folderspec, Optional pstr = "*.txt")
    On Error GoTo errline
    Dim fs, f, f1, fc, i, farr
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(folderspec)
    Set fc = f.Files
    ReDim farr(1 To fc.Count)
    For Each f1 In fc
        If f1.Name Like pstr And Not f1.Name Like "*~$*" Then
            i = i + 1
            farr(i) = f1.Name
        End If
    Next
    ReDim Preserve farr(1 To i)
    filelist = farr
errline:
End Function
```

## 方法一：复制粘贴

```vb
Sub 复制粘贴()

    Dim i, j, k
    Dim fileAr

    fileAr = filelist(ThisWorkbook.Path & "\文件夹\", "*.xlsx")

    Application.ScreenUpdating = False
    For i = 1 To UBound(fileAr)
        With Workbooks.Open(ThisWorkbook.Path & "\文件夹\" & fileAr(i))
            With .Sheets(1)
                .Range("a2:e" & .[a65536].End(3).Row).Copy Sheet1.Range("a" & Sheet1.[a65536].End(3).Row + 1)
            End With
            .Close False
        End With
    Next i
    Application.ScreenUpdating = True

End Sub
```

## 方法二：数组

```vb
Sub 数组()

    Dim i, j, k
    Dim fileAr

    fileAr = filelist(ThisWorkbook.Path & "\文件夹\", "*.xlsx")

    Dim tmp
    Application.ScreenUpdating = False
    For i = 1 To UBound(fileAr)
        With Workbooks.Open(ThisWorkbook.Path & "\文件夹\" & fileAr(i))
            With .Sheets(1)
                tmp = .Range("a2:e" & .[a65536].End(3).Row)
                Sheet1.Range("a" & Sheet1.[a65536].End(3).Row + 1).Resize(UBound(tmp), UBound(tmp, 2)) = tmp
            End With
            .Close False
        End With
    Next i
    Application.ScreenUpdating = True

End Sub
```

[原始链接](https://mp.weixin.qq.com/s?__biz=MzIyOTc3NzQ2NA==&mid=2247485223&idx=1&sn=40015e3160260dc67f98668769a745dc&chksm=e8bcce70dfcb47667654536773ea062c2f7eabad314ade51fa33e34e899a4b17770de329e6ac&scene=178&cur_album_id=3115603487041503237#rd)

---

> 作者: [richfan](https://richfan.site/)  
> URL: http://richfan.site/%E7%A8%8B%E6%8A%80/vba/vba%E6%A1%88%E5%88%97/vba%E6%A1%88%E4%BE%8B012%E5%90%88%E5%B9%B6%E5%B7%A5%E4%BD%9C%E7%B0%BF/  

