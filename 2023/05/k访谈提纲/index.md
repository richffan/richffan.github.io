# k访谈提纲


## k访谈提纲


<!--more-->

```VBA
Sub 插入访谈(x)
    Path = ActiveDocument.AttachedTemplate.FullName
    Application.Templates(Path).BuildingBlockEntries(x).Insert Where:=Selection.Range, RichText:=True
    With ActiveDocument.PageSetup
        .Orientation = wdOrientPortrait
        .TopMargin = CentimetersToPoints(2.54)
        .BottomMargin = CentimetersToPoints(2.54)
        .LeftMargin = CentimetersToPoints(3.17)
        .RightMargin = CentimetersToPoints(2.7)
        .HeaderDistance = CentimetersToPoints(1)
        .FooterDistance = CentimetersToPoints(1.3)
    End With
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    Selection.WholeStory
    Selection.Delete
    Application.Templates(Path).BuildingBlockEntries("访谈页眉").Insert Where:=Selection.Range, RichText:=True
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
    Selection.WholeStory
    Selection.Range.ListFormat.RemoveNumbers NumberType:=wdNumberParagraph
    ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary).PageNumbers.Add PageNumberAlignment:=wdAlignPageNumberCenter, FirstPage:=True
    With ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary) '进入页脚编辑状态
        .Range.Font.Size = 9
        .Range.Font.Name = "宋体"
        .Range.Collapse Direction:=wdCollapseEnd
    End With
End Sub
Sub 财务负责人(control As IRibbonControl)
    Call 插入访谈("财务负责人")
End Sub
Sub 采购负责人(control As IRibbonControl)
    Call 插入访谈("采购负责人")
End Sub
Sub 人力负责人(control As IRibbonControl)
    Call 插入访谈("人力负责人")
End Sub
Sub 生产负责人(control As IRibbonControl)
    Call 插入访谈("生产负责人")
End Sub
Sub 投融资负责人(control As IRibbonControl)
    Call 插入访谈("投融资负责人")
End Sub
Sub 销售负责人(control As IRibbonControl)
    Call 插入访谈("销售负责人")
End Sub
Sub 研发负责人(control As IRibbonControl)
    Call 插入访谈("研发负责人")
End Sub
Sub 总经理(control As IRibbonControl)
    Call 插入访谈("总经理")
End Sub
Sub 独立董事(control As IRibbonControl)
    Call 插入访谈("独立董事")
End Sub
Sub 内审负责人(control As IRibbonControl)
    Call 插入访谈("内审负责人")
End Sub
```



---

> 作者: [richfan](https://richfan.site/)  
> URL: http://richfan.site/2023/05/k%E8%AE%BF%E8%B0%88%E6%8F%90%E7%BA%B2/  

