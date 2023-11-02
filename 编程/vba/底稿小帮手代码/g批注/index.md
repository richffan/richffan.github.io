# g批注



## g批注

<!--more-->

```vb
Sub 添加批注(control As IRibbonControl) '批注-添加批注
    '添加批注
    Application.ScreenUpdating = False '关闭屏幕更新
    Selection.Collapse Direction:=wdCollapseEnd
    ActiveDocument.Comments.Add _
        Range:=Selection.Range, Text:=""
    Application.ScreenUpdating = True '恢复屏幕更新
End Sub
Sub 删除批注(control As IRibbonControl) '批注-删除批注
    '删除批注
    On Error GoTo err_msgbox
    Selection.Comments(1).Delete
    Exit Sub
err_msgbox:
    MsgBox ("你需要先选中一个批注")
End Sub
Sub 修订上色(control As IRibbonControl)
    Dim oDoc As Document
    Set oDoc = Word.ActiveDocument
    Dim oRevision As Revision
    For Each oRevision In oDoc.Revisions
        With oRevision
            t = .Range.Text
            .Range.HighlightColorIndex = wdYellow
        End With
    Next
End Sub
```



---

> 作者: [richfan](https://richfan.site/)  
> URL: http://richfan.site/%E7%BC%96%E7%A8%8B/vba/%E5%BA%95%E7%A8%BF%E5%B0%8F%E5%B8%AE%E6%89%8B%E4%BB%A3%E7%A0%81/g%E6%89%B9%E6%B3%A8/  

