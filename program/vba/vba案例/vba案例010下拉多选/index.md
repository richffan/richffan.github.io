# 【VBA案例010】下拉多选


大家好！今天将为大家介绍在Excel中如何实现下拉多选功能，让数据输入更加灵活高效。

下拉多选功能不仅提高了数据输入的灵活性，还减少了输入错误的可能性，为我们的数据处理工作带来了更高的效率。

也许你见过使用表单控件的方式实现下拉多选，但是通过数据验证和VBA，我们同样能够轻松创建具有下拉多选功能的工作表。

![](https://img.richfan.site/program/vba/vba案列/【VBA案例010】下拉多选.gif)

以下是VBA代码，详细解析请看文末视频。

```vb
Private Sub Worksheet_Change(ByVal Target As Range)

    Dim rngDV As Range
    Dim oldVal As String
    Dim newVal As String
    Dim fenGeFu As String

    If Target.CountLarge > 1 Then Exit Sub
    fenGeFu = "," '规定用逗号分隔

    On Error Resume Next
    Set rngDV = Cells.SpecialCells(xlCellTypeAllValidation)

    If rngDV Is Nothing Then Exit Sub
    If Intersect(Target, rngDV) Is Nothing Then Exit Sub

    Application.EnableEvents = False
    newVal = Target.Value
    Application.Undo
    oldVal = Target.Value
    Target.Value = newVal

    If oldVal <> "" Then
        If newVal <> "" Then
            If InStr(fenGeFu & oldVal & fenGeFu, fenGeFu & newVal & fenGeFu) > 0 Then
                If InStrRev(oldVal, newVal) + Len(newVal) - 1 = Len(oldVal) Then
                    Target.Value = Replace(oldVal, fenGeFu & newVal, "")
                Else
                    Target.Value = Replace(oldVal, newVal & fenGeFu, "")
                End If
            Else
                Target.Value = oldVal & fenGeFu & newVal
            End If
        End If
    End If

    Application.EnableEvents = True
End Sub
```

[原始链接](https://mp.weixin.qq.com/s?__biz=MzIyOTc3NzQ2NA==&mid=2247485185&idx=1&sn=b5dc7e560d9ac9504265d18a7d949286&chksm=e8bcce56dfcb47405814a13e1a7ed65403775118f663853c2c5cd920fd4ed242ad7fccf925ff&scene=178&cur_album_id=3115603487041503237#rd)

---

> 作者: [richfan](https://richfan.site/)  
> URL: http://richfan.site/program/vba/vba%E6%A1%88%E4%BE%8B/vba%E6%A1%88%E4%BE%8B010%E4%B8%8B%E6%8B%89%E5%A4%9A%E9%80%89/  

