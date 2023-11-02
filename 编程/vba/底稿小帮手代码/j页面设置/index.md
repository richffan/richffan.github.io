# j页面设置


## j页面设置

<!--more-->

```vb
Sub 附注页面(control As IRibbonControl) '页面-页面设置
    With ActiveDocument.PageSetup
        .Orientation = wdOrientPortrait
        .TopMargin = CentimetersToPoints(2.54)
        .BottomMargin = CentimetersToPoints(2.54)
        .LeftMargin = CentimetersToPoints(3.17)
        .RightMargin = CentimetersToPoints(2.7)
        .HeaderDistance = CentimetersToPoints(1.5)
        .FooterDistance = CentimetersToPoints(1.75)
    End With
End Sub
Sub 封面页面(control As IRibbonControl) '页面-页面设置
    With ActiveDocument.PageSetup
        .Orientation = wdOrientPortrait
        .TopMargin = CentimetersToPoints(6.06)
        .BottomMargin = CentimetersToPoints(2.54)
        .LeftMargin = CentimetersToPoints(2.24)
        .RightMargin = CentimetersToPoints(2.27)
        .HeaderDistance = CentimetersToPoints(1.5)
        .FooterDistance = CentimetersToPoints(1.75)
    End With
End Sub
Sub 正文页面(control As IRibbonControl) '页面-页面设置
    With ActiveDocument.PageSetup
        .Orientation = wdOrientPortrait
        .TopMargin = CentimetersToPoints(5)
        .BottomMargin = CentimetersToPoints(2)
        .LeftMargin = CentimetersToPoints(3)
        .RightMargin = CentimetersToPoints(3)
        .HeaderDistance = CentimetersToPoints(1.5)
        .FooterDistance = CentimetersToPoints(1.75)
    End With
End Sub
Sub 插入横页(control As IRibbonControl) '页面-插入横页
    Selection.InsertBreak Type:=wdSectionBreakNextPage
    Selection.InsertBreak Type:=wdSectionBreakNextPage
    Selection.GoTo What:=wdGoToPage, Which:=wdGoToPrevious, Count:=1, Name:=""
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = " "
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = True
        .CorrectHangulEndings = True
        .HanjaPhoneticHangul = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    If Selection.PageSetup.Orientation = wdOrientPortrait Then
        Selection.PageSetup.Orientation = wdOrientLandscape
    Else
        Selection.PageSetup.Orientation = wdOrientPortrait
    End If
End Sub
Sub 插入页码(control As IRibbonControl) '页面-插入页码
    Application.ScreenUpdating = False '关闭屏幕更新
    ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary).PageNumbers.Add PageNumberAlignment:=wdAlignPageNumberCenter, FirstPage:=True
    With ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary) '进入页脚编辑状态
        .Range.Font.Size = 9
        .Range.Font.Name = "宋体"
        .Range.Collapse Direction:=wdCollapseEnd
    End With
    With ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).Range '进入页脚编辑状态
        .Delete '删除页眉中的内容
        .ParagraphFormat.Borders(wdBorderBottom).LineStyle = wdLineStyleNone '取消页眉段落下边框线
    End With
    Application.ScreenUpdating = True '恢复屏幕更新
End Sub
```



---

> 作者: [richfan](https://richfan.site/)  
> URL: http://richfan.site/%E7%BC%96%E7%A8%8B/vba/%E5%BA%95%E7%A8%BF%E5%B0%8F%E5%B8%AE%E6%89%8B%E4%BB%A3%E7%A0%81/j%E9%A1%B5%E9%9D%A2%E8%AE%BE%E7%BD%AE/  

