# a段落处理


## a段落处理

<!--more-->

```vb
Sub 段落处理(control As IRibbonControl) '段落-段落处理-
    Application.ScreenUpdating = False
    For Each pg In Selection.Paragraphs
        If pg.Range.Information(wdWithInTable) = False Then
            '-------------------------------------------------------------------------
            If pg.OutlineLevel = wdOutlineLevel1 Then '如果是1级标题用此格式
                With pg.Range.ParagraphFormat
                    .LeftIndent = CentimetersToPoints(0)   '左缩进
                    .RightIndent = CentimetersToPoints(0)  '右缩进
                    .SpaceBefore = 2.5        '段前 2.5磅等于0.5行,如果你的单位为行,则这里乘5
                    .SpaceBeforeAuto = False
                    .SpaceAfter = 2.5         '段后
                    .SpaceAfterAuto = False
                    .LineSpacingRule = wdLineSpace1pt5   '行距,单倍 wdLineSpaceSingle   1.5倍 wdLineSpace1pt5     两倍 wdLineSpaceDouble
                    .Alignment = wdAlignParagraphCenter  '对齐方式  左对齐:wdAlignParagraphLeft  右对齐:wdAlignParagraphRight  居中:wdAlignParagraphCenter
                    .FirstLineIndent = CentimetersToPoints(0)  '首行缩进
                    .CharacterUnitLeftIndent = 0
                    .CharacterUnitRightIndent = 0
                    .CharacterUnitFirstLineIndent = 0

                End With
                With pg.Range
                    .Font.NameFarEast = "黑体"   '中文
                    .Font.NameAscii = "Times New Roman"   '西文
                    .Font.NameOther = "Times New Roman"   '西文
                    .Font.Size = 20          '字号三号
                    .Font.Bold = True        '加粗      不加粗填false
                End With
                '-------------------------------------------------------------------------
            ElseIf pg.OutlineLevel = wdOutlineLevel2 Then '如果是2级标题用此格式
                With pg.Range.ParagraphFormat
                    .LeftIndent = CentimetersToPoints(0)
                    .RightIndent = CentimetersToPoints(0)
                    .SpaceBefore = 2.5
                    .SpaceBeforeAuto = False
                    .SpaceAfter = 2.5
                    .SpaceAfterAuto = False
                    .LineSpacingRule = wdLineSpace1pt5
                    .Alignment = wdAlignParagraphJustify '两端对齐
                    .FirstLineIndent = CentimetersToPoints(0)
                    .CharacterUnitLeftIndent = 0
                    .CharacterUnitRightIndent = 0
                    .CharacterUnitFirstLineIndent = 0

                End With
                With pg.Range
                    .Font.NameFarEast = "黑体"
                    .Font.NameAscii = "Times New Roman"
                    .Font.NameOther = "Times New Roman"
                    .Font.Size = 18          '字号四号
                    .Font.Bold = True

                End With
                '-------------------------------------------------------------------------
            ElseIf pg.OutlineLevel = wdOutlineLevel3 Then  '如果是3级标题用此格式
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
                    .Font.NameFarEast = "黑体"
                    .Font.NameAscii = "Times New Roman"
                    .Font.NameOther = "Times New Roman"
                    .Font.Size = 16          '字号小四
                    .Font.Bold = True

                End With
                '-------------------------------------------------------------------------
            ElseIf pg.OutlineLevel = wdOutlineLevel4 Then  '如果是4级标题用此格式
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
                    .Font.NameFarEast = "仿宋"
                    .Font.NameAscii = "Times New Roman"
                    .Font.NameOther = "Times New Roman"
                    .Font.Size = 16          '字号小四
                    .Font.Bold = True
                End With
                '------------------------------------------------------------------------- 下面是5级以及正文的样式设置
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
                    .FirstLineIndent = CentimetersToPoints(0.35)  '首行缩进2
                    .CharacterUnitLeftIndent = 0
                    .CharacterUnitRightIndent = 0
                    .CharacterUnitFirstLineIndent = 2  '首行缩进2

                End With
                With pg.Range
                    .Font.NameFarEast = "仿宋"
                    .Font.NameAscii = "Times New Roman"
                    .Font.NameOther = "Times New Roman"
                    .Font.Size = 14          '字号四号
                    .Font.Bold = False
                End With
            End If
        End If
    Next
    Application.ScreenUpdating = True
End Sub
Sub 表前单位格式(control As IRibbonControl) '
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
            .Font.NameFarEast = "宋体"
            .Font.NameAscii = "Times New Roman"
            .Font.NameOther = "Times New Roman"
            .Font.Size = 10.5         '字号五号
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
Sub 表后注释格式(control As IRibbonControl) '
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
            .Font.NameFarEast = "宋体"
            .Font.NameAscii = "Times New Roman"
            .Font.NameOther = "Times New Roman"
            .Font.Size = 10        '字号10
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
Sub 自动编号(control As IRibbonControl)  '针对样式（一）
    '先选择一片范围再运行代码,会将例如"(一)"此样式的编号换为自动编号,"()"为中文全角符号
    '注意只有段落开头为"(一)"样式的编号会替换,段中的编号则不会
    Dim r As Range, P As Range, tpf, NF, NS, LI, FI
    '================================================== 配置区
    tpf = "（[一二三四五六七八九十]@）"  '通配符
    NF = "（%1）"   '编号格式,%1为编号本身,不能动,只需要编辑%1旁边的格式,比如'(一)'为'（%1）' 或者 '1、'为 '%1、' 或者 '第一章'为'第%1章'
    NS = wdListNumberStyleSimpChinNum3  '编号的样式:wdListNumberStyleArabic阿拉伯数字    wdListNumberStyleSimpChinNum3中文数字
    LI = CentimetersToPoints(0)    '左缩进
    FI = CentimetersToPoints(0.74)   '首行缩进
    '================================================== 配置区
    Application.ScreenUpdating = False
    If Selection.Type = wdSelectionIP Then
        MsgBox "请选择范围！"
        Exit Sub
    Else
        Set r = Selection.Range
        Set P = Selection.Range
    End If
    With ListGalleries(wdNumberGallery).ListTemplates(1).ListLevels(1)  '设置编号格式
        .NumberFormat = NF
        .TrailingCharacter = wdTrailingNone
        .NumberStyle = NS
        .NumberPosition = 0
        .Alignment = wdListLevelAlignLeft
        .TextPosition = 0
        .TabPosition = wdUndefined
        .ResetOnHigher = 0
        .StartAt = 1
        .LinkedStyle = ""
    End With
    ListGalleries(wdNumberGallery).ListTemplates(1).Name = ""
    With r.Find
        .ClearFormatting
        .Text = tpf
        .Forward = True
        .MatchWildcards = True
        Do While .Execute
            With .Parent
                pat = .Text
                If .End > P.End Then Exit Do
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
    MsgBox "完成"
End Sub
Sub 编号转文本(control As IRibbonControl)
    Dim kgslist As List
    i = MsgBox("点击确定则将该文档下所有编号转为文本", 1)
    If i = 1 Then
        For Each kgslist In ActiveDocument.Lists
            kgslist.ConvertNumbersToText
        Next
    End If
End Sub
```




---

> 作者: [richfan](https://richfan.site/)  
> URL: http://richfan.site/program/vba/%E5%BA%95%E7%A8%BF%E5%B0%8F%E5%B8%AE%E6%89%8B%E4%BB%A3%E7%A0%81/a%E6%AE%B5%E8%90%BD%E5%A4%84%E7%90%86/  

