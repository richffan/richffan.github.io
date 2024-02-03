# k访谈提纲


## k访谈提纲


&lt;!--more--&gt;

```vb
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
    Application.Templates(Path).BuildingBlockEntries(&#34;访谈页眉&#34;).Insert Where:=Selection.Range, RichText:=True
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
    Selection.WholeStory
    Selection.Range.ListFormat.RemoveNumbers NumberType:=wdNumberParagraph
    ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary).PageNumbers.Add PageNumberAlignment:=wdAlignPageNumberCenter, FirstPage:=True
    With ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary) &#39;进入页脚编辑状态
        .Range.Font.Size = 9
        .Range.Font.Name = &#34;宋体&#34;
        .Range.Collapse Direction:=wdCollapseEnd
    End With
End Sub
Sub 财务负责人(control As IRibbonControl)
    Call 插入访谈(&#34;财务负责人&#34;)
End Sub
Sub 采购负责人(control As IRibbonControl)
    Call 插入访谈(&#34;采购负责人&#34;)
End Sub
Sub 人力负责人(control As IRibbonControl)
    Call 插入访谈(&#34;人力负责人&#34;)
End Sub
Sub 生产负责人(control As IRibbonControl)
    Call 插入访谈(&#34;生产负责人&#34;)
End Sub
Sub 投融资负责人(control As IRibbonControl)
    Call 插入访谈(&#34;投融资负责人&#34;)
End Sub
Sub 销售负责人(control As IRibbonControl)
    Call 插入访谈(&#34;销售负责人&#34;)
End Sub
Sub 研发负责人(control As IRibbonControl)
    Call 插入访谈(&#34;研发负责人&#34;)
End Sub
Sub 总经理(control As IRibbonControl)
    Call 插入访谈(&#34;总经理&#34;)
End Sub
Sub 独立董事(control As IRibbonControl)
    Call 插入访谈(&#34;独立董事&#34;)
End Sub
Sub 内审负责人(control As IRibbonControl)
    Call 插入访谈(&#34;内审负责人&#34;)
End Sub
```



---

> 作者:   
> URL: http://richfan.site/%E7%A8%8B%E6%8A%80/vba/%E5%BA%95%E7%A8%BF%E5%B0%8F%E5%B8%AE%E6%89%8B%E4%BB%A3%E7%A0%81/k%E8%AE%BF%E8%B0%88%E6%8F%90%E7%BA%B2/  

