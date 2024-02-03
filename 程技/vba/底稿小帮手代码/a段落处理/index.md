# a段落处理


## a段落处理

&lt;!--more--&gt;

```vb
Sub 段落处理(control As IRibbonControl) &#39;段落-段落处理-
    Application.ScreenUpdating = False
    For Each pg In Selection.Paragraphs
        If pg.Range.Information(wdWithInTable) = False Then
            &#39;-------------------------------------------------------------------------
            If pg.OutlineLevel = wdOutlineLevel1 Then &#39;如果是1级标题用此格式
                With pg.Range.ParagraphFormat
                    .LeftIndent = CentimetersToPoints(0)   &#39;左缩进
                    .RightIndent = CentimetersToPoints(0)  &#39;右缩进
                    .SpaceBefore = 2.5        &#39;段前 2.5磅等于0.5行,如果你的单位为行,则这里乘5
                    .SpaceBeforeAuto = False
                    .SpaceAfter = 2.5         &#39;段后
                    .SpaceAfterAuto = False
                    .LineSpacingRule = wdLineSpace1pt5   &#39;行距,单倍 wdLineSpaceSingle   1.5倍 wdLineSpace1pt5     两倍 wdLineSpaceDouble
                    .Alignment = wdAlignParagraphCenter  &#39;对齐方式  左对齐:wdAlignParagraphLeft  右对齐:wdAlignParagraphRight  居中:wdAlignParagraphCenter
                    .FirstLineIndent = CentimetersToPoints(0)  &#39;首行缩进
                    .CharacterUnitLeftIndent = 0
                    .CharacterUnitRightIndent = 0
                    .CharacterUnitFirstLineIndent = 0

                End With
                With pg.Range
                    .Font.NameFarEast = &#34;黑体&#34;   &#39;中文
                    .Font.NameAscii = &#34;Times New Roman&#34;   &#39;西文
                    .Font.NameOther = &#34;Times New Roman&#34;   &#39;西文
                    .Font.Size = 20          &#39;字号三号
                    .Font.Bold = True        &#39;加粗      不加粗填false
                End With
                &#39;-------------------------------------------------------------------------
            ElseIf pg.OutlineLevel = wdOutlineLevel2 Then &#39;如果是2级标题用此格式
                With pg.Range.ParagraphFormat
                    .LeftIndent = CentimetersToPoints(0)
                    .RightIndent = CentimetersToPoints(0)
                    .SpaceBefore = 2.5
                    .SpaceBeforeAuto = False
                    .SpaceAfter = 2.5
                    .SpaceAfterAuto = False
                    .LineSpacingRule = wdLineSpace1pt5
                    .Alignment = wdAlignParagraphJustify &#39;两端对齐
                    .FirstLineIndent = CentimetersToPoints(0)
                    .CharacterUnitLeftIndent = 0
                    .CharacterUnitRightIndent = 0
                    .CharacterUnitFirstLineIndent = 0

                End With
                With pg.Range
                    .Font.NameFarEast = &#34;黑体&#34;
                    .Font.NameAscii = &#34;Times New Roman&#34;
                    .Font.NameOther = &#34;Times New Roman&#34;
                    .Font.Size = 18          &#39;字号四号
                    .Font.Bold = True

                End With
                &#39;-------------------------------------------------------------------------
            ElseIf pg.OutlineLevel = wdOutlineLevel3 Then  &#39;如果是3级标题用此格式
                With pg.Range.ParagraphFormat
                    .LeftIndent = CentimetersToPoints(0)
                    .RightIndent = CentimetersToPoints(0)
                    .SpaceBefore = 2.5
                    .SpaceBeforeAuto = False
                    .SpaceAfter = 2.5
                    .SpaceAfterAuto = False
                    .LineSpacingRule = wdLineSpace1pt5
                    .Alignment = wdAlignParagraphLeft
                    .FirstLineIndent = CentimetersToPoints(0)
                    .CharacterUnitLeftIndent = 0
                    .CharacterUnitRightIndent = 0
                    .CharacterUnitFirstLineIndent = 0

                End With
                With pg.Range
                    .Font.NameFarEast = &#34;黑体&#34;
                    .Font.NameAscii = &#34;Times New Roman&#34;
                    .Font.NameOther = &#34;Times New Roman&#34;
                    .Font.Size = 16          &#39;字号小四
                    .Font.Bold = True

                End With
                &#39;-------------------------------------------------------------------------
            ElseIf pg.OutlineLevel = wdOutlineLevel4 Then  &#39;如果是4级标题用此格式
                With pg.Range.ParagraphFormat
                    .LeftIndent = CentimetersToPoints(0)
                    .RightIndent = CentimetersToPoints(0)
                    .SpaceBefore = 2.5
                    .SpaceBeforeAuto = False
                    .SpaceAfter = 2.5
                    .SpaceAfterAuto = False
                    .LineSpacingRule = wdLineSpace1pt5
                    .Alignment = wdAlignParagraphJustify
                    .FirstLineIndent = CentimetersToPoints(0)
                    .CharacterUnitLeftIndent = 0
                    .CharacterUnitRightIndent = 0
                    .CharacterUnitFirstLineIndent = 0

                End With
                With pg.Range
                    .Font.NameFarEast = &#34;仿宋&#34;
                    .Font.NameAscii = &#34;Times New Roman&#34;
                    .Font.NameOther = &#34;Times New Roman&#34;
                    .Font.Size = 16          &#39;字号小四
                    .Font.Bold = True
                End With
                &#39;------------------------------------------------------------------------- 下面是5级以及正文的样式设置
            Else
                With pg.Range.ParagraphFormat
                    .LeftIndent = CentimetersToPoints(0)
                    .RightIndent = CentimetersToPoints(0)
                    .SpaceBefore = 2.5
                    .SpaceBeforeAuto = False
                    .SpaceAfter = 2.5
                    .SpaceAfterAuto = False
                    .LineSpacingRule = wdLineSpace1pt5
                    .Alignment = wdAlignParagraphJustify
                    .FirstLineIndent = CentimetersToPoints(0.35)  &#39;首行缩进2
                    .CharacterUnitLeftIndent = 0
                    .CharacterUnitRightIndent = 0
                    .CharacterUnitFirstLineIndent = 2  &#39;首行缩进2

                End With
                With pg.Range
                    .Font.NameFarEast = &#34;仿宋&#34;
                    .Font.NameAscii = &#34;Times New Roman&#34;
                    .Font.NameOther = &#34;Times New Roman&#34;
                    .Font.Size = 14          &#39;字号四号
                    .Font.Bold = False
                End With
            End If
        End If
    Next
    Application.ScreenUpdating = True
End Sub
Sub 表前单位格式(control As IRibbonControl) &#39;
    For Each pg In Selection.Paragraphs
        pg.IndentFirstLineCharWidth -10000
        pg.IndentFirstLineCharWidth 2
        pg.Range.Font.Bold = False
        With pg.Range.ParagraphFormat
            .LeftIndent = CentimetersToPoints(0)
            .RightIndent = CentimetersToPoints(0)
            .SpaceBefore = 0
            .SpaceBeforeAuto = False
            .SpaceAfter = 0
            .SpaceAfterAuto = False
            .LineSpacingRule = wdLineSpaceSingle
            .Alignment = wdAlignParagraphJustify
            .FirstLineIndent = CentimetersToPoints(0)
            .CharacterUnitLeftIndent = 0
            .CharacterUnitRightIndent = 0
            .CharacterUnitFirstLineIndent = 0
            .LineUnitBefore = 0
            .LineUnitAfter = 0
        End With
        With pg.Range
            .Font.NameFarEast = &#34;宋体&#34;
            .Font.NameAscii = &#34;Times New Roman&#34;
            .Font.NameOther = &#34;Times New Roman&#34;
            .Font.Size = 10.5         &#39;字号五号
            .Font.Bold = False
            .ParagraphFormat.Alignment = wdAlignParagraphRight
        End With
        With pg.Range.ParagraphFormat
            .SpaceBefore = 0
            .SpaceBeforeAuto = False
            .SpaceAfter = 0
            .SpaceAfterAuto = False
        End With
    Next
    Application.ScreenUpdating = True
End Sub
Sub 表后注释格式(control As IRibbonControl) &#39;
    For Each pg In Selection.Paragraphs
        pg.IndentFirstLineCharWidth -10000
        pg.IndentFirstLineCharWidth 2
        pg.Range.Font.Bold = False
        With pg.Range.ParagraphFormat
            .LeftIndent = CentimetersToPoints(0)
            .RightIndent = CentimetersToPoints(0)
            .SpaceBefore = 0
            .SpaceBeforeAuto = False
            .SpaceAfter = 0
            .SpaceAfterAuto = False
            .LineSpacingRule = wdLineSpaceSingle
            .Alignment = wdAlignParagraphJustify
            .FirstLineIndent = CentimetersToPoints(0)
            .CharacterUnitLeftIndent = 0
            .CharacterUnitRightIndent = 0
            .CharacterUnitFirstLineIndent = 0
            .LineUnitBefore = 0
            .LineUnitAfter = 0
        End With
        With pg.Range
            .Font.NameFarEast = &#34;宋体&#34;
            .Font.NameAscii = &#34;Times New Roman&#34;
            .Font.NameOther = &#34;Times New Roman&#34;
            .Font.Size = 10        &#39;字号10
            .Font.Bold = False
            .ParagraphFormat.Alignment = wdAlignParagraphJustify
        End With
        With pg.Range.ParagraphFormat
            .SpaceBefore = 0
            .SpaceBeforeAuto = False
            .SpaceAfter = 0
            .SpaceAfterAuto = False
        End With
    Next
    Application.ScreenUpdating = True
End Sub
Sub 自动编号(control As IRibbonControl)  &#39;针对样式（一）
    &#39;先选择一片范围再运行代码,会将例如&#34;(一)&#34;此样式的编号换为自动编号,&#34;()&#34;为中文全角符号
    &#39;注意只有段落开头为&#34;(一)&#34;样式的编号会替换,段中的编号则不会
    Dim r As Range, P As Range, tpf, NF, NS, LI, FI
    &#39;================================================== 配置区
    tpf = &#34;（[一二三四五六七八九十]@）&#34;  &#39;通配符
    NF = &#34;（%1）&#34;   &#39;编号格式,%1为编号本身,不能动,只需要编辑%1旁边的格式,比如&#39;(一)&#39;为&#39;（%1）&#39; 或者 &#39;1、&#39;为 &#39;%1、&#39; 或者 &#39;第一章&#39;为&#39;第%1章&#39;
    NS = wdListNumberStyleSimpChinNum3  &#39;编号的样式:wdListNumberStyleArabic阿拉伯数字    wdListNumberStyleSimpChinNum3中文数字
    LI = CentimetersToPoints(0)    &#39;左缩进
    FI = CentimetersToPoints(0.74)   &#39;首行缩进
    &#39;================================================== 配置区
    Application.ScreenUpdating = False
    If Selection.Type = wdSelectionIP Then
        MsgBox &#34;请选择范围！&#34;
        Exit Sub
    Else
        Set r = Selection.Range
        Set P = Selection.Range
    End If
    With ListGalleries(wdNumberGallery).ListTemplates(1).ListLevels(1)  &#39;设置编号格式
        .NumberFormat = NF
        .TrailingCharacter = wdTrailingNone
        .NumberStyle = NS
        .NumberPosition = 0
        .Alignment = wdListLevelAlignLeft
        .TextPosition = 0
        .TabPosition = wdUndefined
        .ResetOnHigher = 0
        .StartAt = 1
        .LinkedStyle = &#34;&#34;
    End With
    ListGalleries(wdNumberGallery).ListTemplates(1).Name = &#34;&#34;
    With r.Find
        .ClearFormatting
        .Text = tpf
        .Forward = True
        .MatchWildcards = True
        Do While .Execute
            With .Parent
                pat = .Text
                If .End &gt; P.End Then Exit Do
                ast = Asc(ActiveDocument.Range(Start:=.Start - 1, End:=.Start))
                If ast = 13 Or ast = 12 Then
                    .ListFormat.ApplyListTemplateWithLevel ListTemplate:= _
                        ListGalleries(wdNumberGallery).ListTemplates(1), ContinuePreviousList:= _
                        True, ApplyTo:=wdListApplyToWholeList, DefaultListBehavior:= _
                        wdWord10ListBehavior
                    With .ParagraphFormat
                        .SpaceBeforeAuto = False
                        .SpaceAfterAuto = False
                        .LeftIndent = LI
                        .FirstLineIndent = FI
                    End With
                    If .Text = pat Then
                        .Delete
                    End If
                End If
                .Start = .End
            End With
        Loop
    End With
    Application.ScreenUpdating = True
    MsgBox &#34;完成&#34;
End Sub
Sub 编号转文本(control As IRibbonControl)
    Dim kgslist As List
    i = MsgBox(&#34;点击确定则将该文档下所有编号转为文本&#34;, 1)
    If i = 1 Then
        For Each kgslist In ActiveDocument.Lists
            kgslist.ConvertNumbersToText
        Next
    End If
End Sub
```




---

> 作者:   
> URL: http://richfan.site/%E7%A8%8B%E6%8A%80/vba/%E5%BA%95%E7%A8%BF%E5%B0%8F%E5%B8%AE%E6%89%8B%E4%BB%A3%E7%A0%81/a%E6%AE%B5%E8%90%BD%E5%A4%84%E7%90%86/  

