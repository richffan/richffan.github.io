# 【VBA案例005】自动汇总表单

大家好！书接上文。

有时候，我们需要处理多个工作簿，每个工作簿中包含一些特定的信息。为了将这些信息汇总到一个表中，我们可能需要手动打开每个工作簿，然后复制粘贴所需的数据。但这样做既费时又容易出错。

所以，使用VBA依然可以解决这个问题。极大地简化这一过程，让我们能够更专注于其他重要的工作。

举个例子，现在我们有100个工作簿。

![](https://img.richfan.site/program/vba/vba案列/【VBA案例005】自动汇总表单_1.png)

需要从里边提取员工信息汇总到一个表里边。

![](https://img.richfan.site/program/vba/vba案列/【VBA案例005】自动汇总表单_1.png)

猜猜看，用VBA处理这些需要多久？

下面是VBA的代码分享，你也可以观看下方的视频解析：

公共函数：

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

## 方法一：

```vb
Sub 方法一()
    Dim i, j, k
    Dim filePath As String, fileAr

    filePath = ThisWorkbook.Path & "\表单\"
    fileAr = filelist(filePath, "*.xlsx")

    Dim t As Double
    Dim wb As Workbook
    Application.ScreenUpdating = False

    t = Timer

    k = 1
    For i = 1 To UBound(fileAr)
        Set wb = Workbooks.Open(filePath & fileAr(i))
        With wb.Worksheets(1)
            k = k + 1
            For j = 1 To 5
                ThisWorkbook.Worksheets("Sheet1").Cells(k, j) = .Range("c" & j + 4)
            Next
        End With
        wb.Close False
    Next i
    Application.ScreenUpdating = True

    MsgBox Format(Timer - t, "0.000s")
End Sub
```

## 方法二：

```vb
Sub 方法二数组()
    Dim i, j, k
    Dim filePath As String, fileAr
    Dim br(1 To 5000, 1 To 5)
    filePath = ThisWorkbook.Path & "\表单\"
    fileAr = filelist(filePath, "*.xlsx")

    Dim t As Double
    Dim wb As Workbook
    Application.ScreenUpdating = False

    t = Timer

    k = 0
    Dim tmp
    For i = 1 To UBound(fileAr)
        Set wb = Workbooks.Open(filePath & fileAr(i))

        With wb.Worksheets(1)
            tmp = .Range("c5:c9")
            k = k + 1
            For j = 1 To 5
                br(i, j) = tmp(j, 1)
            Next
        End With
        wb.Close False
    Next i
    ThisWorkbook.Worksheets(1).[a2].Resize(k, UBound(br, 2)) = br
    Application.ScreenUpdating = True

    MsgBox Format(Timer - t, "0.000s")
End Sub
```

## 方法三：

```vb
Sub 方法三()
    Dim i, j, k
    Dim filePath As String, fileAr
    Dim br(1 To 5000, 1 To 5)
    filePath = ThisWorkbook.Path & "\表单\"
    fileAr = filelist(filePath, "*.xlsx")

    Dim t As Double
    Dim wb As Workbook
    Application.ScreenUpdating = False

    t = Timer

    k = 0
    Dim tmp
    For i = 1 To UBound(fileAr)
        k = k + 1
        For j = 1 To 5
            br(i, j) = Application.ExecuteExcel4Macro("'" & filePath & "[" & fileAr(i) & "]信息卡'!" & Range("c" & j + 4).Address(, , xlR1C1))
        Next
    Next i
    ThisWorkbook.Worksheets(1).[a2].Resize(k, UBound(br, 2)) = br
    Application.ScreenUpdating = True

    MsgBox Format(Timer - t, "0.000s")
End Sub
```

[原始链接](https://mp.weixin.qq.com/s?__biz=MzIyOTc3NzQ2NA==&mid=2247485052&idx=1&sn=1459ec753f39116a5682b5b3f558531e&chksm=e8bccf2bdfcb463db93c57b9d87e0bd281a9025543a839ce7762334c2d83bd48cf44a7ec54b1&scene=178&cur_album_id=3115603487041503237#rd)

---

> 作者: [richfan](https://richfan.site/)  
> URL: http://richfan.site/posts/program/vba/vba%E6%A1%88%E4%BE%8B/vba%E6%A1%88%E4%BE%8B005%E8%87%AA%E5%8A%A8%E6%B1%87%E6%80%BB%E8%A1%A8%E5%8D%95/  

