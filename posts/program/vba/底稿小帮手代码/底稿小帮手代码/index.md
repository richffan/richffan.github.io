# 底稿小帮手代码


## a段落处理

<!--more-->

```VBA
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


## b表单处理

```VBA
Dim aarr(1 To 20), bbrr(1 To 30, 1 To 30) '多列调整
Sub 单表_一键调整(control As IRibbonControl) '单表-格式
    '功能：光标在表格中处理当前表格；否则处理所有表格！
    Application.ScreenUpdating = False  '关闭屏幕刷新
    Application.DisplayAlerts = False  '关闭提示
    On Error Resume Next  '忽略错误
    '-------------------------------------------------------------------------
    Dim mytable As Table, i As Long
    For Each mytable In Selection.Tables
        With mytable
            .Shading.ForegroundPatternColor = wdColorAutomatic
            .Shading.BackgroundPatternColor = wdColorAutomatic
            Options.DefaultHighlightColorIndex = wdNoHighlight
            .Range.HighlightColorIndex = wdNoHighlight
            .Style = "表格主题"
            With .Borders(wdBorderLeft)    '左框线
                .LineStyle = wdLineStyleSingle   '设置线条样式    不需要线条则填wdLineStyleNone
                .LineWidth = wdLineWidth150pt    '宽度为1.5
            End With
            With .Borders(wdBorderRight)   '右框线
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth150pt
            End With
            With .Borders(wdBorderTop)     '上框线
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth150pt
            End With
            With .Borders(wdBorderBottom)  '下框线
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth150pt
            End With
            With .Borders(wdBorderVertical)   '内部纵向框线
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth050pt
            End With
            With .Borders(wdBorderHorizontal)   '内部横向框线
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth050pt
            End With
            .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone  '左上的斜线
            .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone    '右上的斜线
            '单元格边距
            .TopPadding = CentimetersToPoints(0) '设置上边距为0
            .BottomPadding = CentimetersToPoints(0) '设置下边距为0
            .LeftPadding = PixelsToPoints(0, True)  '设置左边距为0
            .RightPadding = PixelsToPoints(0, True) '设置右边距为0
            .Spacing = PixelsToPoints(0, True) '允许单元格间距为0
            .AllowPageBreaks = True '允许断页
            '.AllowAutoFit = True '允许自动重调尺寸
            With .Rows
                .WrapAroundText = False '取消文字环绕
                '.Alignment = wdAlignRowCenter '表水平居中  wdAlignRowLeft '左对齐
                .AllowBreakAcrossPages = False '不允许行断页
                .Height = CentimetersToPoints(0.8) '行高0.8
                .HeightRule = wdRowHeightAtLeast '行高设为最小值
                .LeftIndent = CentimetersToPoints(0) '左面缩进量为0
            End With
            With .Range
                With .Font '字体格式
                    .NameFarEast = "宋体"
                    .NameAscii = "Times New Roman"
                    .NameOther = "Times New Roman"
                    .Color = wdColorAutomatic '自动字体颜色
                    .Size = 10.5   '字号
                    .Kerning = 0
                    .DisableCharacterSpaceGrid = True  '选定段落中的字符与行网格对齐
                End With
                With .ParagraphFormat '段落格式
                    .LineUnitBefore = 0
                    .LineUnitAfter = 0
                    .SpaceBefore = 0
                    .SpaceAfter = 0
                    .CharacterUnitFirstLineIndent = 0 '取消首行缩进
                    .FirstLineIndent = CentimetersToPoints(0) '取消首行缩进
                    .LineSpacingRule = wdLineSpaceSingle 'wdLineSpaceSingle '单倍行距  wdLineSpaceExactly '行距固定值
                    ''.LineSpacing = 18 '设置行间距为18磅，配合行距固定值
                    '.Alignment = wdAlignParagraphCenter '单元格水平居中
                    .AutoAdjustRightIndent = False  '自动调整所选段落的右缩进
                    .DisableLineHeightGrid = True   '选定段落中的字符与行网格对齐
                End With
                .Cells.VerticalAlignment = wdCellAlignVerticalCenter  '单元格垂直居中
            End With
            For Each cl In .Range.Cells    '文字靠左，数字靠右,合计居中,序号居中
                Acell = ActiveDocument.Range(cl.Range.Start, cl.Range.End - 1).Text '提取文本
                If IsNumeric(Acell) Then
                    cl.Range.ParagraphFormat.Alignment = wdAlignParagraphRight    '右对齐
                Else
                    cl.Range.ParagraphFormat.Alignment = wdAlignParagraphJustify   '左对齐
                    If Acell = "合计" Or Acell = "总计" Or Acell = "总 计" Or Acell = "合 计" Then
                        cl.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter    '水平居中
                        If cl.ColumnIndex = .Columns.Count Then
                            .Columns(cl.ColumnIndex).Select
                            Selection.Font.Bold = True
                        Else
                            cl.Row.Range.Font.Bold = True
                        End If
                    ElseIf Acell = "序号" Or Acell = "序 号" Then
                        .Columns(cl.ColumnIndex).Select
                        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter    '水平居中
                    End If
                End If
            Next
            '设置首行格式
            .Rows(1).Select ' 选中第一个单元格
            With Selection
                .Rows.HeadingFormat = wdToggle '自动标题行重复
                .ParagraphFormat.Alignment = wdAlignParagraphCenter   '水平居中
                .Range.Font.Bold = True '表头加粗黑体
                .Shading.ForegroundPatternColor = wdColorAutomatic '首行自动颜色
                .Shading.BackgroundPatternColor = -603923969 '首行底纹填充,不要底色则删了这行
                '.Borders(wdBorderBottom).LineStyle = xlContinuous
                '.Borders(wdBorderBottom).LineWidth = wdLineWidth50pt
            End With
            '自动调整表格
            .Columns.PreferredWidthType = wdPreferredWidthAuto
            .AutoFitBehavior (wdAutoFitContent) '根据内容调整表格
            .AutoFitBehavior (wdAutoFitWindow) '根据窗口调整表格
        End With
    Next
    '---------------------------------------------------------------------------------------
    ERR.Clear: On Error GoTo 0 '恢复错误捕捉
    Application.DisplayAlerts = True  '开启提示
    Application.ScreenUpdating = True   '开启屏幕刷新
End Sub
Sub 格宽调整_释放(control As IRibbonControl) '列宽调整-多列加载
    Set mytable = Selection.Tables(1)
    For i = 1 To mytable.Rows.Count
        For j = 1 To mytable.Rows(i).Cells.Count
            mytable.Rows(i).Cells(j).Width = bbrr(i, j)
        Next j
    Next i
End Sub
Sub 格宽调整_读取(control As IRibbonControl)  '列宽调整-多列读取
    Set mytable = Selection.Tables(1)
    mytable.AutoFitBehavior (wdAutoFitFixed)
    For i = 1 To mytable.Rows.Count
        For j = 1 To mytable.Rows(i).Cells.Count
            bbrr(i, j) = mytable.Rows(i).Cells(j).Width
        Next j
    Next i
End Sub
Sub 列宽调整_读取(control As IRibbonControl)   '列宽调整-单列读取
    With Selection.Tables(1)
        ColumnsCounts = .Columns.Count
        For i = 1 To ColumnsCounts
            aarr(i) = .Columns(i).Width
        Next
    End With
End Sub
Sub 列宽调整_释放(control As IRibbonControl) '列宽调整-单列加载
    With Selection.Tables(1)
        .AutoFitBehavior (wdAutoFitFixed)
        ColumnsCounts = .Columns.Count
        For i = 1 To ColumnsCounts
            .Columns(i).Width = aarr(i)
        Next
    End With
End Sub
```


## c数字

```VBA
Sub 千分位符(control As IRibbonControl) '数字-千分位符
    On Error Resume Next
    Dim i As Range, Acell As Cell, CR As Range
    On Error Resume Next
    Application.ScreenUpdating = False
    If Selection.Type = 2 Then '文档选定
        For Each i In Selection.Words
            If IsNumeric(i) Then
                If i Like "####*" = True Then
                    If i.Next Like "." = True And i.Next(wdWord, 2) Like "#*" = True Then
                        i.SetRange Start:=i.Start, End:=i.Next(wdWord, 2).End
                        NC = Format(i, "#,##0.00;-#,##0.00; ")
                        i.Text = NC
                    Else
                        NC = Format(i, "#,##0.00;-#,##0.00; ")
                        i.Text = NC
                    End If
                End If
            End If
        Next i
    ElseIf Selection.Type = 4 Or Selection.Type = 5 Then '竖形表格（5为横形表格）
        For Each Acell In Selection.Cells
            Set CR = ActiveDocument.Range(Acell.Range.Start, Acell.Range.End - 1)
            '            MsgBox CR
            If CR Like "-####*" Or "-####.#*" = True Then
                Yn = Format(CR, "#,##0.00;-#,##0.00; ")
                CR.Text = Yn
            Else
                If CR Like "####*" Or "####.#*" = True Then
                    Yn = Format(CR, "#,##0.00;-#,##0.00; ")
                    CR.Text = Yn
                End If
            End If
        Next Acell
    Else
        MsgBox "您只能选定文本或者表格之一!", vbOK + vbInformation
    End If
    Application.ScreenUpdating = True
    Application.Activate
End Sub
Sub 除以一万(control As IRibbonControl)
    Application.ScreenUpdating = False
    If Selection.Type = 2 Then
        If IsNumeric(Selection.Text) Then
            i = Selection.Text
            P = i / 10000
            q = Format(Round(P, 2), "#,##0.00;-#,##0.00; ")
            Selection.Text = q & "万"
        End If
    ElseIf Selection.Type = 5 Or Selection.Type = 4 Then
        For Each Acell In Selection.Cells
            Set CR = ActiveDocument.Range(Acell.Range.Start, Acell.Range.End - 1)
            'MsgBox CR
            If IsNumeric(CR.Text) Then
                i = CR.Text
                P = i / 10000
                q = Format(Round(P, 2), "#,##0.00;-#,##0.00; ")
                CR.Text = q & "万"
            End If
        Next
    End If
    Application.ScreenUpdating = True
End Sub
Sub 乘百(control As IRibbonControl)
    Application.ScreenUpdating = False
    If Selection.Type = 2 Then
        If IsNumeric(Selection.Text) Then
            i = Selection.Text
            P = i * 100
            q = Format(Round(P, 2), "#,##0.00;-#,##0.00; ")
            Selection.Text = q & "%"
        End If
    ElseIf Selection.Type = 5 Or Selection.Type = 4 Then
        For Each Acell In Selection.Cells
            Set CR = ActiveDocument.Range(Acell.Range.Start, Acell.Range.End - 1)
            'MsgBox CR
            If IsNumeric(CR.Text) Then
                i = CR.Text
                P = i * 100
                q = Format(Round(P, 2), "#,##0.00;-#,##0.00; ")
                CR.Text = q & "%"
            End If
        Next
    End If
        Application.ScreenUpdating = True
End Sub
```

## d文字

```VBA
Sub 宋体宋体(control As IRibbonControl)
    '选中范围字体为宋体+宋体
    With Selection.Font
        .NameFarEast = "宋体"
        .NameAscii = "宋体"
        .NameOther = "宋体"
    End With
End Sub
Sub 宋体罗马(control As IRibbonControl) '文字-宋体罗马
    '选中范围字体为宋体+Times
    With Selection.Font
        .NameFarEast = "仿宋"
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
    End With
End Sub
Sub 楷体加粗(control As IRibbonControl) '文字-楷体加粗
    '选中范围字体为楷体加粗
    With Selection.Font
        .NameFarEast = "楷体"
        .NameAscii = "楷体"
        .NameOther = "楷体"
        .Name = "楷体"
        .Bold = True
    End With
End Sub
Sub 去除空白(control As IRibbonControl) '文字-去除空白
    '删除换行及空格
    
    Selection.Find.ClearFormatting   '删除空格
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = " "
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting   '删除大空格
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "　"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.Replacement.ClearFormatting    '删除连续两个回车
    With Selection.Find
        .Text = "^p^p"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
Sub 中英文标点互换(control As IRibbonControl) ' 文字-中英文标点互换
    Dim ChineseInterpunction() As Variant, EnglishInterpunction() As Variant
    Dim myArray1() As Variant, myArray2() As Variant, strFind As String, strRep As String
    Dim msgResult As VbMsgBoxResult, n As Byte
    ' 定义一个中文标点的数组对象
    ChineseInterpunction = Array("，", "。", "；", "：", "？", "！", "……", "—", "～", "（", "）", "《", "》")
    ' 定义一个英文标点的数组对象
    EnglishInterpunction = Array(",", ".", ";", ":", "?", "!", "…", "-", "~", "(", ")", "&lt;", "&gt;")
    ' 提示用户交互的 MSGBOX 对话框
    msgResult = MsgBox("您想中英标点互换吗？按 Y 将中文标点转为英文标点，按 N 将英文标点转为中文标点！", vbYesNoCancel)
    Select Case msgResult
        Case vbCancel
            Exit Sub ' 如果用户选择了取消按钮，则退出程序运行
        Case vbYes ' 如果用户选择了 YES, 则将中文标点转换为英文标点
            myArray1 = ChineseInterpunction
            myArray2 = EnglishInterpunction
            strFind = "“(*)”"
            strRep = """\1"""
        Case vbNo ' 如果用户选择了 NO, 则将英文标点转换为中文标点
            myArray1 = EnglishInterpunction
            myArray2 = ChineseInterpunction
            strFind = """(*)"""
            strRep = "“\1”"
    End Select
    Application.ScreenUpdating = False ' 关闭屏幕更新
    For n = 0 To UBound(ChineseInterpunction)  ' 从数组的下标到上标间作一个循环
        With Selection.Find
            .ClearFormatting ' 不限定查找格式
            .MatchWildcards = False ' 不使用通配符
            ' 查找相应的英文标点，替换为对应的中文标点
            .Execute findtext:=myArray1(n), replacewith:=myArray2(n), Replace:=wdReplaceAll
        End With
    Next
    With Selection.Find
        .ClearFormatting ' 不限定查找格式
        .MatchWildcards = True ' 使用通配符
        .Execute findtext:=strFind, replacewith:=strRep, Replace:=wdReplaceAll
    End With
    Application.ScreenUpdating = True ' 恢复屏幕更新
End Sub
Sub 高亮(control As IRibbonControl) '文字-HighLight
    If Selection.Range.HighlightColorIndex = 0 Then
        Selection.Range.HighlightColorIndex = wdYellow
    Else
        Selection.Range.HighlightColorIndex = wdNoHighlight
    End If
End Sub
```

## e大纲

```VBA
Sub 大纲一级(control As IRibbonControl) '大纲调整-一级
    With Selection
        .Paragraphs.OutlineLevel = wdOutlineLevel1
    End With
End Sub
Sub 大纲二级(control As IRibbonControl) '大纲调整-二级
    With Selection
        .Paragraphs.OutlineLevel = wdOutlineLevel2
    End With
End Sub
Sub 大纲三级(control As IRibbonControl) '大纲调整-三级
    With Selection
        .Paragraphs.OutlineLevel = wdOutlineLevel3
    End With
End Sub
Sub 大纲四级(control As IRibbonControl) '大纲调整-四级
    With Selection
        .Paragraphs.OutlineLevel = wdOutlineLevel4
    End With
End Sub
Sub 大纲五级(control As IRibbonControl) '大纲调整-五级
    With Selection
        .Paragraphs.OutlineLevel = wdOutlineLevel5
    End With
End Sub
Sub 大纲正文(control As IRibbonControl) '大纲调整-正文
    With Selection
        .Paragraphs.OutlineLevel = wdOutlineLevelBodyText
    End With
End Sub
```

## fExcel贴数

```VBA
Sub 粘贴格式文本(control As IRibbonControl)
    Set xl = GetObject(, "excel.application")
    xlr = xl.Selection.Rows.Count
    xlc = xl.Selection.Columns.Count
    With Selection
        wdc = .Information(16)
        wdr = .Information(13)
        rangeselect wdr, wdc, xlr, xlc
        ReDim arr(1 To 1)
        For Each sht In .Cells
            i = i + 1
            ReDim Preserve arr(1 To i)
            arr(i) = sht.Range.Font.Underline
        Next
        .CopyFormat
        .PasteAndFormat (wdFormatSurroundingFormattingWithEmphasis)
        .PasteFormat
        With .Find
            .Text = " "
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindStop
            .Execute Replace:=wdReplaceAll
        End With
        rangeselect wdr, wdc, xlr, xlc
        For Each sht In .Cells
            j = j + 1
            sht.Range.Font.Underline = arr(j)
        Next
    End With

End Sub
Sub 双下划线(control As IRibbonControl)
    If Selection.Font.Underline = wdUnderlineDouble Then
        Selection.Font.Underline = wdUnderlineNone
    ElseIf Selection.Font.Underline = wdUnderlineNone Then
        Selection.Font.Underline = wdUnderlineDouble
    End If
End Sub
Sub 单下划线(control As IRibbonControl)
    If Selection.Font.Underline = wdUnderlineSingle Then
        Selection.Font.Underline = wdUnderlineNone
    ElseIf Selection.Font.Underline = wdUnderlineNone Then
        Selection.Font.Underline = wdUnderlineSingle
    End If
End Sub
Function rangeselect(wdr, wdc, xlr, xlc)
    With Selection
        .Tables(1).Cell(wdr, wdc).Select
        .Collapse wdCollapseStart
        st = .Start
        .Tables(1).Cell(wdr + xlr - 1, wdc + xlc - 1).Select
        .Collapse wdCollapseEnd
        ed = .End
        ActiveDocument.Range(st, ed).Select
    End With
End Function
```

## g批注

```VBA
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

## h行列校验

```VBA
Sub 选行校验(control As IRibbonControl) '行列校验-选行校验
    i = 0
    x = 0
    n = 0
    If Selection.Type = wdSelectionColumn Then
        For Each Acell In Selection.Cells '求所选列合计数
            If Acell.ColumnIndex > n Then
                n = Acell.ColumnIndex
            End If
            Set CR1 = ActiveDocument.Range(Acell.Range.Start, Acell.Range.End - 1)
            If IsNumeric(CR1.Text) Then
                i = CR1.Text + i
            End If
        Next
        For Each Acell In Selection.Cells '求所选列最后一行合计数
            Set CR2 = ActiveDocument.Range(Acell.Range.Start, Acell.Range.End - 1)
            If Acell.ColumnIndex = n Then
                If IsNumeric(CR2.Text) Then
                    x = CR2.Text + x
                End If
            End If
        Next
        P = i - x
        q = P - x
        y = "列校验:" & Format(P, "Standard") & "-" & Format(x, "Standard") & "=" & Format(q, "Standard")
    Else
        y = "只支持选中表格范围!"
    End If
    Application.StatusBar = y
End Sub
Sub 选列校验(control As IRibbonControl) '行列校验-选列校验
    i = 0
    x = 0
    n = 0
    If Selection.Type = wdSelectionColumn Then
        For Each Acell In Selection.Cells '求所选列合计数
            If Acell.RowIndex > n Then
                n = Acell.RowIndex
            End If
            Set CR1 = ActiveDocument.Range(Acell.Range.Start, Acell.Range.End - 1)
            If IsNumeric(CR1.Text) Then
                i = CR1.Text + i
            End If
        Next
        For Each Acell In Selection.Cells '求所选列最后一行合计数
            Set CR2 = ActiveDocument.Range(Acell.Range.Start, Acell.Range.End - 1)
            If Acell.RowIndex = n Then
                If IsNumeric(CR2.Text) Then
                    x = CR2.Text + x
                End If
            End If
        Next
        P = i - x
        q = P - x
        y = "列校验:" & Format(P, "Standard") & "-" & Format(x, "Standard") & "=" & Format(q, "Standard")
    Else
        y = "只支持选中表格范围!"
    End If
    Application.StatusBar = y
End Sub
Sub 区域求和(control As IRibbonControl) '行列校验-区域求和
    i = 0
    If Selection.Type = 5 Then
        For Each Acell In Selection.Cells
            Set CR = ActiveDocument.Range(Acell.Range.Start, Acell.Range.End - 1)
            If IsNumeric(CR.Text) Then
                i = CR.Text + i
            End If
        Next
        i = "合计:" & Format(i, "#,##0.00;-#,##0.00; ")
    ElseIf Selection.Type = 4 Then
        For Each Acell In Selection.Cells
            Set CR = ActiveDocument.Range(Acell.Range.Start, Acell.Range.End - 1)
            If IsNumeric(CR.Text) Then
                i = CR.Text + i
            End If
        Next
        i = "合计:" & Format(i, "#,##0.00;-#,##0.00; ")
    Else
        i = "只支持表格内行或列求和!"
    End If
    Application.StatusBar = i
End Sub
```

## i多表处理

```VBA
Sub 多表_一键调整(control As IRibbonControl) '批量调整表格格式
    Application.ScreenUpdating = False  '关闭屏幕刷新
    Application.DisplayAlerts = False  '关闭提示
    On Error Resume Next  '忽略错误
    '-------------------------------------------------------------------------
    Dim mytable As Table, i As Long
    For Each mytable In ActiveDocument.Tables
        With mytable
            .Shading.ForegroundPatternColor = wdColorAutomatic
            .Shading.BackgroundPatternColor = wdColorAutomatic
            Options.DefaultHighlightColorIndex = wdNoHighlight
            .Range.HighlightColorIndex = wdNoHighlight
            .Style = "表格主题"
            With .Borders(wdBorderLeft)
                .LineStyle = wdLineStyleSingle: .LineWidth = wdLineWidth150pt
            End With
            With .Borders(wdBorderRight)
                .LineStyle = wdLineStyleSingle: .LineWidth = wdLineWidth150pt
            End With
            With .Borders(wdBorderTop)
                .LineStyle = wdLineStyleSingle: .LineWidth = wdLineWidth150pt
            End With
            With .Borders(wdBorderBottom)
                .LineStyle = wdLineStyleSingle: .LineWidth = wdLineWidth150pt
            End With
            With .Borders(wdBorderVertical)
                .LineStyle = wdLineStyleSingle: .LineWidth = wdLineWidth050pt
            End With
            .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
            .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
            '单元格边距
            .TopPadding = CentimetersToPoints(0) '设置上边距为0
            .BottomPadding = CentimetersToPoints(0) '设置下边距为0
            .LeftPadding = PixelsToPoints(0, True)  '设置左边距为0
            .RightPadding = PixelsToPoints(0, True) '设置右边距为0
            .Spacing = PixelsToPoints(0, True) '允许单元格间距为0
            .AllowPageBreaks = True '允许断页
            '.AllowAutoFit = True '允许自动重调尺寸
            With .Rows
                .WrapAroundText = False '取消文字环绕
                '.Alignment = wdAlignRowCenter '表水平居中  wdAlignRowLeft '左对齐
                .AllowBreakAcrossPages = False '不允许行断页
                .Height = CentimetersToPoints(0.8) '行高0.8
                .HeightRule = wdRowHeightAtLeast '行高设为最小值
                .LeftIndent = CentimetersToPoints(0) '左面缩进量为0
            End With
            With .Range
                With .Font '字体格式
                    .NameFarEast = "宋体"
                    .NameAscii = "Times New Roman"
                    .NameOther = "Times New Roman"
                    .Color = wdColorAutomatic '自动字体颜色
                    .Size = 10.5   '字号
                    .Kerning = 0
                    .DisableCharacterSpaceGrid = True  '选定段落中的字符与行网格对齐
                End With
                With .ParagraphFormat '段落格式
                    .LineUnitBefore = 0
                    .LineUnitAfter = 0
                    .SpaceBefore = 0
                    .SpaceAfter = 0
                    .CharacterUnitFirstLineIndent = 0 '取消首行缩进
                    .FirstLineIndent = CentimetersToPoints(0) '取消首行缩进
                    .LineSpacingRule = wdLineSpaceSingle 'wdLineSpaceSingle '单倍行距  wdLineSpaceExactly '行距固定值
                    ''.LineSpacing = 18 '设置行间距为18磅，配合行距固定值
                    '.Alignment = wdAlignParagraphCenter '单元格水平居中
                    .AutoAdjustRightIndent = False  '自动调整所选段落的右缩进
                    .DisableLineHeightGrid = True   '选定段落中的字符与行网格对齐
                End With
                .Cells.VerticalAlignment = wdCellAlignVerticalCenter  '单元格垂直居中
            End With
            For Each cl In .Range.Cells    '文字靠左，数字靠右
                Acell = ActiveDocument.Range(cl.Range.Start, cl.Range.End - 1).Text '提取文本
                If IsNumeric(Acell) Then
                    cl.Range.ParagraphFormat.Alignment = wdAlignParagraphRight    '右对齐
                Else
                    cl.Range.ParagraphFormat.Alignment = wdAlignParagraphJustify   '左对齐
                    If Acell = "合计" Or Acell = "总计" Or Acell = "总 计" Or Acell = "合 计" Then
                        cl.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter    '水平居中
                        If cl.ColumnIndex = .Columns.Count Then
                            .Columns(cl.ColumnIndex).Select
                            Selection.Font.Bold = True
                        Else
                            cl.Row.Range.Font.Bold = True
                        End If
                    ElseIf Acell = "序号" Or Acell = "序 号" Then
                        .Columns(cl.ColumnIndex).Select
                        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter    '水平居中
                    End If
                End If
            Next
            '设置首行格式
            .Rows(1).Select ' 选中第一个单元格
            With Selection
                .Rows.HeadingFormat = wdToggle '自动标题行重复
                .ParagraphFormat.Alignment = wdAlignParagraphCenter   '水平居中
                .Range.Font.Bold = True '表头加粗黑体
                .Shading.ForegroundPatternColor = wdColorAutomatic '首行自动颜色
                .Shading.BackgroundPatternColor = -603923969 '首行底纹填充
                '.Borders(wdBorderBottom).LineStyle = xlContinuous
                '.Borders(wdBorderBottom).LineWidth = wdLineWidth50pt
            End With
            '自动调整表格
            .Columns.PreferredWidthType = wdPreferredWidthAuto
            .AutoFitBehavior (wdAutoFitContent) '根据内容调整表格
            .AutoFitBehavior (wdAutoFitWindow) '根据窗口调整表格
        End With
    Next
    '---------------------------------------------------------------------------------------
    ERR.Clear: On Error GoTo 0 '恢复错误捕捉
    Application.DisplayAlerts = True  '开启提示
    Application.ScreenUpdating = True   '开启屏幕刷新
End Sub
Sub 多表处理_表格选中(control As IRibbonControl) '
    Dim tempTable As Table
    'Application.ScreenUpdating = False
    '判断文档是否被保护
    If ActiveDocument.ProtectionType = wdAllowOnlyFormFields Then
        MsgBox "文档已保护，此时不能选中多个表格！"
        Exit Sub
    End If
    '删除所有可编辑的区域
    ActiveDocument.DeleteAllEditableRanges wdEditorEveryone
    '添加可编辑区域
    For Each tempTable In ActiveDocument.Tables
        tempTable.Range.Editors.Add wdEditorEveryone
    Next
    '选中所有可编辑区域
    ActiveDocument.SelectAllEditableRanges wdEditorEveryone
    '删除所有可编辑的区域
    ActiveDocument.DeleteAllEditableRanges wdEditorEveryone
    'Application.ScreenUpdating = True
    MsgBox "完成"
End Sub
```

## JsonConverter

```VBA
''
' VBA-JSON v2.3.1
' (c) Tim Hall - https://github.com/VBA-tools/VBA-JSON
'
' JSON Converter for VBA
'
' Errors:
' 10001 - JSON parse error
'
' @class JsonConverter
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
'
' Based originally on vba-json (with extensive changes)
' BSD license included below
'
' JSONLib, http://code.google.com/p/vba-json/
'
' Copyright (c) 2013, Ryo Yokoyama
' All rights reserved.
'
' Redistribution and use in source and binary forms, with or without
' modification, are permitted provided that the following conditions are met:
'     * Redistributions of source code must retain the above copyright
'       notice, this list of conditions and the following disclaimer.
'     * Redistributions in binary form must reproduce the above copyright
'       notice, this list of conditions and the following disclaimer in the
'       documentation and/or other materials provided with the distribution.
'     * Neither the name of the <organization> nor the
'       names of its contributors may be used to endorse or promote products
'       derived from this software without specific prior written permission.
'
' THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
' ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
' WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
' DISCLAIMED. IN NO EVENT SHALL <COPYRIGHT HOLDER> BE LIABLE FOR ANY
' DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
' (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
' LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
' ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
' (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
' SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

' === VBA-UTC Headers
#If Mac Then
    
    #If VBA7 Then
        
        ' 64-bit Mac (2016)
        Private Declare PtrSafe Function utc_popen Lib "/usr/lib/libc.dylib" Alias "popen" _
            (ByVal utc_Command As String, ByVal utc_Mode As String) As LongPtr
        Private Declare PtrSafe Function utc_pclose Lib "/usr/lib/libc.dylib" Alias "pclose" _
            (ByVal utc_File As LongPtr) As LongPtr
        Private Declare PtrSafe Function utc_fread Lib "/usr/lib/libc.dylib" Alias "fread" _
            (ByVal utc_Buffer As String, ByVal utc_Size As LongPtr, ByVal utc_Number As LongPtr, ByVal utc_File As LongPtr) As LongPtr
        Private Declare PtrSafe Function utc_feof Lib "/usr/lib/libc.dylib" Alias "feof" _
            (ByVal utc_File As LongPtr) As LongPtr
        
    #Else
        
        ' 32-bit Mac
        Private Declare Function utc_popen Lib "libc.dylib" Alias "popen" _
            (ByVal utc_Command As String, ByVal utc_Mode As String) As Long
        Private Declare Function utc_pclose Lib "libc.dylib" Alias "pclose" _
            (ByVal utc_File As Long) As Long
        Private Declare Function utc_fread Lib "libc.dylib" Alias "fread" _
            (ByVal utc_Buffer As String, ByVal utc_Size As Long, ByVal utc_Number As Long, ByVal utc_File As Long) As Long
        Private Declare Function utc_feof Lib "libc.dylib" Alias "feof" _
            (ByVal utc_File As Long) As Long
        
    #End If
    
    #ElseIf VBA7 Then
        
        ' http://msdn.microsoft.com/en-us/library/windows/desktop/ms724421.aspx
        ' http://msdn.microsoft.com/en-us/library/windows/desktop/ms724949.aspx
        ' http://msdn.microsoft.com/en-us/library/windows/desktop/ms725485.aspx
        Private Declare PtrSafe Function utc_GetTimeZoneInformation Lib "kernel32" Alias "GetTimeZoneInformation" _
            (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION) As Long
        Private Declare PtrSafe Function utc_SystemTimeToTzSpecificLocalTime Lib "kernel32" Alias "SystemTimeToTzSpecificLocalTime" _
            (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpUniversalTime As utc_SYSTEMTIME, utc_lpLocalTime As utc_SYSTEMTIME) As Long
        Private Declare PtrSafe Function utc_TzSpecificLocalTimeToSystemTime Lib "kernel32" Alias "TzSpecificLocalTimeToSystemTime" _
            (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpLocalTime As utc_SYSTEMTIME, utc_lpUniversalTime As utc_SYSTEMTIME) As Long
        
    #Else
        
        Private Declare Function utc_GetTimeZoneInformation Lib "kernel32" Alias "GetTimeZoneInformation" _
            (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION) As Long
        Private Declare Function utc_SystemTimeToTzSpecificLocalTime Lib "kernel32" Alias "SystemTimeToTzSpecificLocalTime" _
            (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpUniversalTime As utc_SYSTEMTIME, utc_lpLocalTime As utc_SYSTEMTIME) As Long
        Private Declare Function utc_TzSpecificLocalTimeToSystemTime Lib "kernel32" Alias "TzSpecificLocalTimeToSystemTime" _
            (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpLocalTime As utc_SYSTEMTIME, utc_lpUniversalTime As utc_SYSTEMTIME) As Long
        
    #End If
    
    #If Mac Then
        
        #If VBA7 Then
            Private Type utc_ShellResult
                utc_Output As String
                utc_ExitCode As LongPtr
            End Type
            
        #Else
            
            Private Type utc_ShellResult
                utc_Output As String
                utc_ExitCode As Long
            End Type
            
        #End If
        
    #Else
        
        Private Type utc_SYSTEMTIME
            utc_wYear As Integer
            utc_wMonth As Integer
            utc_wDayOfWeek As Integer
            utc_wDay As Integer
            utc_wHour As Integer
            utc_wMinute As Integer
            utc_wSecond As Integer
            utc_wMilliseconds As Integer
        End Type
        
        Private Type utc_TIME_ZONE_INFORMATION
            utc_Bias As Long
            utc_StandardName(0 To 31) As Integer
            utc_StandardDate As utc_SYSTEMTIME
            utc_StandardBias As Long
            utc_DaylightName(0 To 31) As Integer
            utc_DaylightDate As utc_SYSTEMTIME
            utc_DaylightBias As Long
        End Type
        
    #End If
    ' === End VBA-UTC
    
    Private Type json_Options
        ' VBA only stores 15 significant digits, so any numbers larger than that are truncated
        ' This can lead to issues when BIGINT's are used (e.g. for Ids or Credit Cards), as they will be invalid above 15 digits
        ' See: http://support.microsoft.com/kb/269370
        '
        ' By default, VBA-JSON will use String for numbers longer than 15 characters that contain only digits
        ' to override set `JsonConverter.JsonOptions.UseDoubleForLargeNumbers = True`
        UseDoubleForLargeNumbers As Boolean
        
        ' The JSON standard requires object keys to be quoted (" or '), use this option to allow unquoted keys
        AllowUnquotedKeys As Boolean
        
        ' The solidus (/) is not required to be escaped, use this option to escape them as \/ in ConvertToJson
        EscapeSolidus As Boolean
    End Type
    Public JsonOptions As json_Options
    
    ' ============================================= '
    ' Public Methods
    ' ============================================= '
    
    ''
    ' Convert JSON string to object (Dictionary/Collection)
    '
    ' @method ParseJson
    ' @param {String} json_String
    ' @return {Object} (Dictionary or Collection)
    ' @throws 10001 - JSON parse error
    ''
    Public Function ParseJson(ByVal JsonString As String) As Object
        Dim json_Index As Long
        json_Index = 1
        
        ' Remove vbCr, vbLf, and vbTab from json_String
        JsonString = VBA.Replace(VBA.Replace(VBA.Replace(JsonString, VBA.vbCr, ""), VBA.vbLf, ""), VBA.vbTab, "")
        
        json_SkipSpaces JsonString, json_Index
        Select Case VBA.Mid$(JsonString, json_Index, 1)
            Case "{"
                Set ParseJson = json_ParseObject(JsonString, json_Index)
            Case "["
                Set ParseJson = json_ParseArray(JsonString, json_Index)
            Case Else
                ' Error: Invalid JSON string
                ERR.Raise 10001, "JSONConverter", json_ParseErrorMessage(JsonString, json_Index, "Expecting '{' or '['")
        End Select
    End Function
    
    ''
    ' Convert object (Dictionary/Collection/Array) to JSON
    '
    ' @method ConvertToJson
    ' @param {Variant} JsonValue (Dictionary, Collection, or Array)
    ' @param {Integer|String} Whitespace "Pretty" print json with given number of spaces per indentation (Integer) or given string
    ' @return {String}
    ''
    Public Function ConvertToJson(ByVal JsonValue As Variant, Optional ByVal Whitespace As Variant, Optional ByVal json_CurrentIndentation As Long = 0) As String
        Dim json_Buffer As String
        Dim json_BufferPosition As Long
        Dim json_BufferLength As Long
        Dim json_Index As Long
        Dim json_LBound As Long
        Dim json_UBound As Long
        Dim json_IsFirstItem As Boolean
        Dim json_Index2D As Long
        Dim json_LBound2D As Long
        Dim json_UBound2D As Long
        Dim json_IsFirstItem2D As Boolean
        Dim json_Key As Variant
        Dim json_Value As Variant
        Dim json_DateStr As String
        Dim json_Converted As String
        Dim json_SkipItem As Boolean
        Dim json_PrettyPrint As Boolean
        Dim json_Indentation As String
        Dim json_InnerIndentation As String
        
        json_LBound = -1
        json_UBound = -1
        json_IsFirstItem = True
        json_LBound2D = -1
        json_UBound2D = -1
        json_IsFirstItem2D = True
        json_PrettyPrint = Not IsMissing(Whitespace)
        
        Select Case VBA.VarType(JsonValue)
            Case VBA.vbNull
                ConvertToJson = "null"
            Case VBA.vbDate
                ' Date
                json_DateStr = ConvertToIso(VBA.CDate(JsonValue))
                
                ConvertToJson = """" & json_DateStr & """"
            Case VBA.vbString
                ' String (or large number encoded as string)
                If Not JsonOptions.UseDoubleForLargeNumbers And json_StringIsLargeNumber(JsonValue) Then
                    ConvertToJson = JsonValue
                Else
                    ConvertToJson = """" & json_Encode(JsonValue) & """"
                End If
            Case VBA.vbBoolean
                If JsonValue Then
                    ConvertToJson = "true"
                Else
                    ConvertToJson = "false"
                End If
            Case VBA.vbArray To VBA.vbArray + VBA.vbByte
                If json_PrettyPrint Then
                    If VBA.VarType(Whitespace) = VBA.vbString Then
                        json_Indentation = VBA.String$(json_CurrentIndentation + 1, Whitespace)
                        json_InnerIndentation = VBA.String$(json_CurrentIndentation + 2, Whitespace)
                    Else
                        json_Indentation = VBA.Space$((json_CurrentIndentation + 1) * Whitespace)
                        json_InnerIndentation = VBA.Space$((json_CurrentIndentation + 2) * Whitespace)
                    End If
                End If
                
                ' Array
                json_BufferAppend json_Buffer, "[", json_BufferPosition, json_BufferLength
                
                On Error Resume Next
                
                json_LBound = LBound(JsonValue, 1)
                json_UBound = UBound(JsonValue, 1)
                json_LBound2D = LBound(JsonValue, 2)
                json_UBound2D = UBound(JsonValue, 2)
                
                If json_LBound >= 0 And json_UBound >= 0 Then
                    For json_Index = json_LBound To json_UBound
                        If json_IsFirstItem Then
                            json_IsFirstItem = False
                        Else
                            ' Append comma to previous line
                            json_BufferAppend json_Buffer, ",", json_BufferPosition, json_BufferLength
                        End If
                        
                        If json_LBound2D >= 0 And json_UBound2D >= 0 Then
                            ' 2D Array
                            If json_PrettyPrint Then
                                json_BufferAppend json_Buffer, vbNewLine, json_BufferPosition, json_BufferLength
                            End If
                            json_BufferAppend json_Buffer, json_Indentation & "[", json_BufferPosition, json_BufferLength
                            
                            For json_Index2D = json_LBound2D To json_UBound2D
                                If json_IsFirstItem2D Then
                                    json_IsFirstItem2D = False
                                Else
                                    json_BufferAppend json_Buffer, ",", json_BufferPosition, json_BufferLength
                                End If
                                
                                json_Converted = ConvertToJson(JsonValue(json_Index, json_Index2D), Whitespace, json_CurrentIndentation + 2)
                                
                                ' For Arrays/Collections, undefined (Empty/Nothing) is treated as null
                                If json_Converted = "" Then
                                    ' (nest to only check if converted = "")
                                    If json_IsUndefined(JsonValue(json_Index, json_Index2D)) Then
                                        json_Converted = "null"
                                    End If
                                End If
                                
                                If json_PrettyPrint Then
                                    json_Converted = vbNewLine & json_InnerIndentation & json_Converted
                                End If
                                
                                json_BufferAppend json_Buffer, json_Converted, json_BufferPosition, json_BufferLength
                            Next json_Index2D
                            
                            If json_PrettyPrint Then
                                json_BufferAppend json_Buffer, vbNewLine, json_BufferPosition, json_BufferLength
                            End If
                            
                            json_BufferAppend json_Buffer, json_Indentation & "]", json_BufferPosition, json_BufferLength
                            json_IsFirstItem2D = True
                        Else
                            ' 1D Array
                            json_Converted = ConvertToJson(JsonValue(json_Index), Whitespace, json_CurrentIndentation + 1)
                            
                            ' For Arrays/Collections, undefined (Empty/Nothing) is treated as null
                            If json_Converted = "" Then
                                ' (nest to only check if converted = "")
                                If json_IsUndefined(JsonValue(json_Index)) Then
                                    json_Converted = "null"
                                End If
                            End If
                            
                            If json_PrettyPrint Then
                                json_Converted = vbNewLine & json_Indentation & json_Converted
                            End If
                            
                            json_BufferAppend json_Buffer, json_Converted, json_BufferPosition, json_BufferLength
                        End If
                    Next json_Index
                End If
                
                On Error GoTo 0
                
                If json_PrettyPrint Then
                    json_BufferAppend json_Buffer, vbNewLine, json_BufferPosition, json_BufferLength
                    
                    If VBA.VarType(Whitespace) = VBA.vbString Then
                        json_Indentation = VBA.String$(json_CurrentIndentation, Whitespace)
                    Else
                        json_Indentation = VBA.Space$(json_CurrentIndentation * Whitespace)
                    End If
                End If
                
                json_BufferAppend json_Buffer, json_Indentation & "]", json_BufferPosition, json_BufferLength
                
                ConvertToJson = json_BufferToString(json_Buffer, json_BufferPosition)
                
                ' Dictionary or Collection
            Case VBA.vbObject
                If json_PrettyPrint Then
                    If VBA.VarType(Whitespace) = VBA.vbString Then
                        json_Indentation = VBA.String$(json_CurrentIndentation + 1, Whitespace)
                    Else
                        json_Indentation = VBA.Space$((json_CurrentIndentation + 1) * Whitespace)
                    End If
                End If
                
                ' Dictionary
                If VBA.TypeName(JsonValue) = "Dictionary" Then
                    json_BufferAppend json_Buffer, "{", json_BufferPosition, json_BufferLength
                    For Each json_Key In JsonValue.Keys
                        ' For Objects, undefined (Empty/Nothing) is not added to object
                        json_Converted = ConvertToJson(JsonValue(json_Key), Whitespace, json_CurrentIndentation + 1)
                        If json_Converted = "" Then
                            json_SkipItem = json_IsUndefined(JsonValue(json_Key))
                        Else
                            json_SkipItem = False
                        End If
                        
                        If Not json_SkipItem Then
                            If json_IsFirstItem Then
                                json_IsFirstItem = False
                            Else
                                json_BufferAppend json_Buffer, ",", json_BufferPosition, json_BufferLength
                            End If
                            
                            If json_PrettyPrint Then
                                json_Converted = vbNewLine & json_Indentation & """" & json_Key & """: " & json_Converted
                            Else
                                json_Converted = """" & json_Key & """:" & json_Converted
                            End If
                            
                            json_BufferAppend json_Buffer, json_Converted, json_BufferPosition, json_BufferLength
                        End If
                    Next json_Key
                    
                    If json_PrettyPrint Then
                        json_BufferAppend json_Buffer, vbNewLine, json_BufferPosition, json_BufferLength
                        
                        If VBA.VarType(Whitespace) = VBA.vbString Then
                            json_Indentation = VBA.String$(json_CurrentIndentation, Whitespace)
                        Else
                            json_Indentation = VBA.Space$(json_CurrentIndentation * Whitespace)
                        End If
                    End If
                    
                    json_BufferAppend json_Buffer, json_Indentation & "}", json_BufferPosition, json_BufferLength
                    
                    ' Collection
                ElseIf VBA.TypeName(JsonValue) = "Collection" Then
                    json_BufferAppend json_Buffer, "[", json_BufferPosition, json_BufferLength
                    For Each json_Value In JsonValue
                        If json_IsFirstItem Then
                            json_IsFirstItem = False
                        Else
                            json_BufferAppend json_Buffer, ",", json_BufferPosition, json_BufferLength
                        End If
                        
                        json_Converted = ConvertToJson(json_Value, Whitespace, json_CurrentIndentation + 1)
                        
                        ' For Arrays/Collections, undefined (Empty/Nothing) is treated as null
                        If json_Converted = "" Then
                            ' (nest to only check if converted = "")
                            If json_IsUndefined(json_Value) Then
                                json_Converted = "null"
                            End If
                        End If
                        
                        If json_PrettyPrint Then
                            json_Converted = vbNewLine & json_Indentation & json_Converted
                        End If
                        
                        json_BufferAppend json_Buffer, json_Converted, json_BufferPosition, json_BufferLength
                    Next json_Value
                    
                    If json_PrettyPrint Then
                        json_BufferAppend json_Buffer, vbNewLine, json_BufferPosition, json_BufferLength
                        
                        If VBA.VarType(Whitespace) = VBA.vbString Then
                            json_Indentation = VBA.String$(json_CurrentIndentation, Whitespace)
                        Else
                            json_Indentation = VBA.Space$(json_CurrentIndentation * Whitespace)
                        End If
                    End If
                    
                    json_BufferAppend json_Buffer, json_Indentation & "]", json_BufferPosition, json_BufferLength
                End If
                
                ConvertToJson = json_BufferToString(json_Buffer, json_BufferPosition)
            Case VBA.vbInteger, VBA.vbLong, VBA.vbSingle, VBA.vbDouble, VBA.vbCurrency, VBA.vbDecimal
                ' Number (use decimals for numbers)
                ConvertToJson = VBA.Replace(JsonValue, ",", ".")
            Case Else
                ' vbEmpty, vbError, vbDataObject, vbByte, vbUserDefinedType
                ' Use VBA's built-in to-string
                On Error Resume Next
                ConvertToJson = JsonValue
                On Error GoTo 0
        End Select
    End Function
    
    ' ============================================= '
    ' Private Functions
    ' ============================================= '
    
    Private Function json_ParseObject(json_String As String, ByRef json_Index As Long) As Dictionary
        Dim json_Key As String
        Dim json_NextChar As String
        
        Set json_ParseObject = New Dictionary
        json_SkipSpaces json_String, json_Index
        If VBA.Mid$(json_String, json_Index, 1) <> "{" Then
            ERR.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting '{'")
        Else
            json_Index = json_Index + 1
            
            Do
                json_SkipSpaces json_String, json_Index
                If VBA.Mid$(json_String, json_Index, 1) = "}" Then
                    json_Index = json_Index + 1
                    Exit Function
                ElseIf VBA.Mid$(json_String, json_Index, 1) = "," Then
                    json_Index = json_Index + 1
                    json_SkipSpaces json_String, json_Index
                End If
                
                json_Key = json_ParseKey(json_String, json_Index)
                json_NextChar = json_Peek(json_String, json_Index)
                If json_NextChar = "[" Or json_NextChar = "{" Then
                    Set json_ParseObject.Item(json_Key) = json_ParseValue(json_String, json_Index)
                Else
                    json_ParseObject.Item(json_Key) = json_ParseValue(json_String, json_Index)
                End If
            Loop
        End If
    End Function
    
    Private Function json_ParseArray(json_String As String, ByRef json_Index As Long) As Collection
        Set json_ParseArray = New Collection
        
        json_SkipSpaces json_String, json_Index
        If VBA.Mid$(json_String, json_Index, 1) <> "[" Then
            ERR.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting '['")
        Else
            json_Index = json_Index + 1
            
            Do
                json_SkipSpaces json_String, json_Index
                If VBA.Mid$(json_String, json_Index, 1) = "]" Then
                    json_Index = json_Index + 1
                    Exit Function
                ElseIf VBA.Mid$(json_String, json_Index, 1) = "," Then
                    json_Index = json_Index + 1
                    json_SkipSpaces json_String, json_Index
                End If
                
                json_ParseArray.Add json_ParseValue(json_String, json_Index)
            Loop
        End If
    End Function
    
    Private Function json_ParseValue(json_String As String, ByRef json_Index As Long) As Variant
        json_SkipSpaces json_String, json_Index
        Select Case VBA.Mid$(json_String, json_Index, 1)
            Case "{"
                Set json_ParseValue = json_ParseObject(json_String, json_Index)
            Case "["
                Set json_ParseValue = json_ParseArray(json_String, json_Index)
            Case """", "'"
                json_ParseValue = json_ParseString(json_String, json_Index)
            Case Else
                If VBA.Mid$(json_String, json_Index, 4) = "true" Then
                    json_ParseValue = True
                    json_Index = json_Index + 4
                ElseIf VBA.Mid$(json_String, json_Index, 5) = "false" Then
                    json_ParseValue = False
                    json_Index = json_Index + 5
                ElseIf VBA.Mid$(json_String, json_Index, 4) = "null" Then
                    json_ParseValue = Null
                    json_Index = json_Index + 4
                ElseIf VBA.InStr("+-0123456789", VBA.Mid$(json_String, json_Index, 1)) Then
                    json_ParseValue = json_ParseNumber(json_String, json_Index)
                Else
                    ERR.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting 'STRING', 'NUMBER', null, true, false, '{', or '['")
                End If
        End Select
    End Function
    
    Private Function json_ParseString(json_String As String, ByRef json_Index As Long) As String
        Dim json_Quote As String
        Dim json_Char As String
        Dim json_Code As String
        Dim json_Buffer As String
        Dim json_BufferPosition As Long
        Dim json_BufferLength As Long
        
        json_SkipSpaces json_String, json_Index
        
        ' Store opening quote to look for matching closing quote
        json_Quote = VBA.Mid$(json_String, json_Index, 1)
        json_Index = json_Index + 1
        
        Do While json_Index > 0 And json_Index <= Len(json_String)
            json_Char = VBA.Mid$(json_String, json_Index, 1)
            
            Select Case json_Char
                Case "\"
                    ' Escaped string, \\, or \/
                    json_Index = json_Index + 1
                    json_Char = VBA.Mid$(json_String, json_Index, 1)
                    
                    Select Case json_Char
                        Case """", "\", "/", "'"
                            json_BufferAppend json_Buffer, json_Char, json_BufferPosition, json_BufferLength
                            json_Index = json_Index + 1
                        Case "b"
                            json_BufferAppend json_Buffer, vbBack, json_BufferPosition, json_BufferLength
                            json_Index = json_Index + 1
                        Case "f"
                            json_BufferAppend json_Buffer, vbFormFeed, json_BufferPosition, json_BufferLength
                            json_Index = json_Index + 1
                        Case "n"
                            json_BufferAppend json_Buffer, vbCrLf, json_BufferPosition, json_BufferLength
                            json_Index = json_Index + 1
                        Case "r"
                            json_BufferAppend json_Buffer, vbCr, json_BufferPosition, json_BufferLength
                            json_Index = json_Index + 1
                        Case "t"
                            json_BufferAppend json_Buffer, vbTab, json_BufferPosition, json_BufferLength
                            json_Index = json_Index + 1
                        Case "u"
                            ' Unicode character escape (e.g. \u00a9 = Copyright)
                            json_Index = json_Index + 1
                            json_Code = VBA.Mid$(json_String, json_Index, 4)
                            json_BufferAppend json_Buffer, VBA.ChrW(VBA.Val("&h" + json_Code)), json_BufferPosition, json_BufferLength
                            json_Index = json_Index + 4
                    End Select
                Case json_Quote
                    json_ParseString = json_BufferToString(json_Buffer, json_BufferPosition)
                    json_Index = json_Index + 1
                    Exit Function
                Case Else
                    json_BufferAppend json_Buffer, json_Char, json_BufferPosition, json_BufferLength
                    json_Index = json_Index + 1
            End Select
        Loop
    End Function
    
    Private Function json_ParseNumber(json_String As String, ByRef json_Index As Long) As Variant
        Dim json_Char As String
        Dim json_Value As String
        Dim json_IsLargeNumber As Boolean
        
        json_SkipSpaces json_String, json_Index
        
        Do While json_Index > 0 And json_Index <= Len(json_String)
            json_Char = VBA.Mid$(json_String, json_Index, 1)
            
            If VBA.InStr("+-0123456789.eE", json_Char) Then
                ' Unlikely to have massive number, so use simple append rather than buffer here
                json_Value = json_Value & json_Char
                json_Index = json_Index + 1
            Else
                ' Excel only stores 15 significant digits, so any numbers larger than that are truncated
                ' This can lead to issues when BIGINT's are used (e.g. for Ids or Credit Cards), as they will be invalid above 15 digits
                ' See: http://support.microsoft.com/kb/269370
                '
                ' Fix: Parse -> String, Convert -> String longer than 15/16 characters containing only numbers and decimal points -> Number
                ' (decimal doesn't factor into significant digit count, so if present check for 15 digits + decimal = 16)
                json_IsLargeNumber = IIf(InStr(json_Value, "."), Len(json_Value) >= 17, Len(json_Value) >= 16)
                If Not JsonOptions.UseDoubleForLargeNumbers And json_IsLargeNumber Then
                    json_ParseNumber = json_Value
                Else
                    ' VBA.Val does not use regional settings, so guard for comma is not needed
                    json_ParseNumber = VBA.Val(json_Value)
                End If
                Exit Function
            End If
        Loop
    End Function
    
    Private Function json_ParseKey(json_String As String, ByRef json_Index As Long) As String
        ' Parse key with single or double quotes
        If VBA.Mid$(json_String, json_Index, 1) = """" Or VBA.Mid$(json_String, json_Index, 1) = "'" Then
            json_ParseKey = json_ParseString(json_String, json_Index)
        ElseIf JsonOptions.AllowUnquotedKeys Then
            Dim json_Char As String
            Do While json_Index > 0 And json_Index <= Len(json_String)
                json_Char = VBA.Mid$(json_String, json_Index, 1)
                If (json_Char <> " ") And (json_Char <> ":") Then
                    json_ParseKey = json_ParseKey & json_Char
                    json_Index = json_Index + 1
                Else
                    Exit Do
                End If
            Loop
        Else
            ERR.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting '""' or '''")
        End If
        
        ' Check for colon and skip if present or throw if not present
        json_SkipSpaces json_String, json_Index
        If VBA.Mid$(json_String, json_Index, 1) <> ":" Then
            ERR.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting ':'")
        Else
            json_Index = json_Index + 1
        End If
    End Function
    
    Private Function json_IsUndefined(ByVal json_Value As Variant) As Boolean
        ' Empty / Nothing -> undefined
        Select Case VBA.VarType(json_Value)
            Case VBA.vbEmpty
                json_IsUndefined = True
            Case VBA.vbObject
                Select Case VBA.TypeName(json_Value)
                    Case "Empty", "Nothing"
                        json_IsUndefined = True
                End Select
        End Select
    End Function
    
    Private Function json_Encode(ByVal json_Text As Variant) As String
        ' Reference: http://www.ietf.org/rfc/rfc4627.txt
        ' Escape: ", \, /, backspace, form feed, line feed, carriage return, tab
        Dim json_Index As Long
        Dim json_Char As String
        Dim json_AscCode As Long
        Dim json_Buffer As String
        Dim json_BufferPosition As Long
        Dim json_BufferLength As Long
        
        For json_Index = 1 To VBA.Len(json_Text)
            json_Char = VBA.Mid$(json_Text, json_Index, 1)
            json_AscCode = VBA.AscW(json_Char)
            
            ' When AscW returns a negative number, it returns the twos complement form of that number.
            ' To convert the twos complement notation into normal binary notation, add 0xFFF to the return result.
            ' https://support.microsoft.com/en-us/kb/272138
            If json_AscCode < 0 Then
                json_AscCode = json_AscCode + 65536
            End If
            
            ' From spec, ", \, and control characters must be escaped (solidus is optional)
            
            Select Case json_AscCode
                Case 34
                    ' " -> 34 -> \"
                    json_Char = "\"""
                Case 92
                    ' \ -> 92 -> \\
                    json_Char = "\\"
                Case 47
                    ' / -> 47 -> \/ (optional)
                    If JsonOptions.EscapeSolidus Then
                        json_Char = "\/"
                    End If
                Case 8
                    ' backspace -> 8 -> \b
                    json_Char = "\b"
                Case 12
                    ' form feed -> 12 -> \f
                    json_Char = "\f"
                Case 10
                    ' line feed -> 10 -> \n
                    json_Char = "\n"
                Case 13
                    ' carriage return -> 13 -> \r
                    json_Char = "\r"
                Case 9
                    ' tab -> 9 -> \t
                    json_Char = "\t"
                Case 0 To 31, 127 To 65535
                    ' Non-ascii characters -> convert to 4-digit hex
                    json_Char = "\u" & VBA.Right$("0000" & VBA.Hex$(json_AscCode), 4)
            End Select
            
            json_BufferAppend json_Buffer, json_Char, json_BufferPosition, json_BufferLength
        Next json_Index
        
        json_Encode = json_BufferToString(json_Buffer, json_BufferPosition)
    End Function
    
    Private Function json_Peek(json_String As String, ByVal json_Index As Long, Optional json_NumberOfCharacters As Long = 1) As String
        ' "Peek" at the next number of characters without incrementing json_Index (ByVal instead of ByRef)
        json_SkipSpaces json_String, json_Index
        json_Peek = VBA.Mid$(json_String, json_Index, json_NumberOfCharacters)
    End Function
    
    Private Sub json_SkipSpaces(json_String As String, ByRef json_Index As Long)
        ' Increment index to skip over spaces
        Do While json_Index > 0 And json_Index <= VBA.Len(json_String) And VBA.Mid$(json_String, json_Index, 1) = " "
            json_Index = json_Index + 1
        Loop
    End Sub
    
    Private Function json_StringIsLargeNumber(json_String As Variant) As Boolean
        ' Check if the given string is considered a "large number"
        ' (See json_ParseNumber)
        
        Dim json_Length As Long
        Dim json_CharIndex As Long
        json_Length = VBA.Len(json_String)
        
        ' Length with be at least 16 characters and assume will be less than 100 characters
        If json_Length >= 16 And json_Length <= 100 Then
            Dim json_CharCode As String
            
            json_StringIsLargeNumber = True
            
            For json_CharIndex = 1 To json_Length
                json_CharCode = VBA.Asc(VBA.Mid$(json_String, json_CharIndex, 1))
                Select Case json_CharCode
                    ' Look for .|0-9|E|e
                    Case 46, 48 To 57, 69, 101
                        ' Continue through characters
                    Case Else
                        json_StringIsLargeNumber = False
                        Exit Function
                End Select
            Next json_CharIndex
        End If
    End Function
    
    Private Function json_ParseErrorMessage(json_String As String, ByRef json_Index As Long, ErrorMessage As String)
        ' Provide detailed parse error message, including details of where and what occurred
        '
        ' Example:
        ' Error parsing JSON:
        ' {"abcde":True}
        '          ^
        ' Expecting 'STRING', 'NUMBER', null, true, false, '{', or '['
        
        Dim json_StartIndex As Long
        Dim json_StopIndex As Long
        
        ' Include 10 characters before and after error (if possible)
        json_StartIndex = json_Index - 10
        json_StopIndex = json_Index + 10
        If json_StartIndex <= 0 Then
            json_StartIndex = 1
        End If
        If json_StopIndex > VBA.Len(json_String) Then
            json_StopIndex = VBA.Len(json_String)
        End If
        
        json_ParseErrorMessage = "Error parsing JSON:" & VBA.vbNewLine & _
            VBA.Mid$(json_String, json_StartIndex, json_StopIndex - json_StartIndex + 1) & VBA.vbNewLine & _
            VBA.Space$(json_Index - json_StartIndex) & "^" & VBA.vbNewLine & _
            ErrorMessage
    End Function
    
    Private Sub json_BufferAppend(ByRef json_Buffer As String, _
            ByRef json_Append As Variant, _
            ByRef json_BufferPosition As Long, _
            ByRef json_BufferLength As Long)
        ' VBA can be slow to append strings due to allocating a new string for each append
        ' Instead of using the traditional append, allocate a large empty string and then copy string at append position
        '
        ' Example:
        ' Buffer: "abc  "
        ' Append: "def"
        ' Buffer Position: 3
        ' Buffer Length: 5
        '
        ' Buffer position + Append length > Buffer length -> Append chunk of blank space to buffer
        ' Buffer: "abc       "
        ' Buffer Length: 10
        '
        ' Put "def" into buffer at position 3 (0-based)
        ' Buffer: "abcdef    "
        '
        ' Approach based on cStringBuilder from vbAccelerator
        ' http://www.vbaccelerator.com/home/VB/Code/Techniques/RunTime_Debug_Tracing/VB6_Tracer_Utility_zip_cStringBuilder_cls.asp
        '
        ' and clsStringAppend from Philip Swannell
        ' https://github.com/VBA-tools/VBA-JSON/pull/82
        
        Dim json_AppendLength As Long
        Dim json_LengthPlusPosition As Long
        
        json_AppendLength = VBA.Len(json_Append)
        json_LengthPlusPosition = json_AppendLength + json_BufferPosition
        
        If json_LengthPlusPosition > json_BufferLength Then
            ' Appending would overflow buffer, add chunk
            ' (double buffer length or append length, whichever is bigger)
            Dim json_AddedLength As Long
            json_AddedLength = IIf(json_AppendLength > json_BufferLength, json_AppendLength, json_BufferLength)
            
            json_Buffer = json_Buffer & VBA.Space$(json_AddedLength)
            json_BufferLength = json_BufferLength + json_AddedLength
        End If
        
        ' Note: Namespacing with VBA.Mid$ doesn't work properly here, throwing compile error:
        ' Function call on left-hand side of assignment must return Variant or Object
        Mid$(json_Buffer, json_BufferPosition + 1, json_AppendLength) = CStr(json_Append)
        json_BufferPosition = json_BufferPosition + json_AppendLength
    End Sub
    
    Private Function json_BufferToString(ByRef json_Buffer As String, ByVal json_BufferPosition As Long) As String
        If json_BufferPosition > 0 Then
            json_BufferToString = VBA.Left$(json_Buffer, json_BufferPosition)
        End If
    End Function
    
    ''
    ' VBA-UTC v1.0.6
    ' (c) Tim Hall - https://github.com/VBA-tools/VBA-UtcConverter
    '
    ' UTC/ISO 8601 Converter for VBA
    '
    ' Errors:
    ' 10011 - UTC parsing error
    ' 10012 - UTC conversion error
    ' 10013 - ISO 8601 parsing error
    ' 10014 - ISO 8601 conversion error
    '
    ' @module UtcConverter
    ' @author tim.hall.engr@gmail.com
    ' @license MIT (http://www.opensource.org/licenses/mit-license.php)
    '' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
    
    ' (Declarations moved to top)
    
    ' ============================================= '
    ' Public Methods
    ' ============================================= '
    
    ''
    ' Parse UTC date to local date
    '
    ' @method ParseUtc
    ' @param {Date} UtcDate
    ' @return {Date} Local date
    ' @throws 10011 - UTC parsing error
    ''
    Public Function ParseUtc(utc_UtcDate As Date) As Date
        On Error GoTo utc_ErrorHandling
        
        #If Mac Then
            ParseUtc = utc_ConvertDate(utc_UtcDate)
        #Else
            Dim utc_TimeZoneInfo As utc_TIME_ZONE_INFORMATION
            Dim utc_LocalDate As utc_SYSTEMTIME
            
            utc_GetTimeZoneInformation utc_TimeZoneInfo
            utc_SystemTimeToTzSpecificLocalTime utc_TimeZoneInfo, utc_DateToSystemTime(utc_UtcDate), utc_LocalDate
            
            ParseUtc = utc_SystemTimeToDate(utc_LocalDate)
        #End If
        
        Exit Function
        
utc_ErrorHandling:
        ERR.Raise 10011, "UtcConverter.ParseUtc", "UTC parsing error: " & ERR.Number & " - " & ERR.Description
    End Function
    
    ''
    ' Convert local date to UTC date
    '
    ' @method ConvertToUrc
    ' @param {Date} utc_LocalDate
    ' @return {Date} UTC date
    ' @throws 10012 - UTC conversion error
    ''
    Public Function ConvertToUtc(utc_LocalDate As Date) As Date
        On Error GoTo utc_ErrorHandling
        
        #If Mac Then
            ConvertToUtc = utc_ConvertDate(utc_LocalDate, utc_ConvertToUtc:=True)
        #Else
            Dim utc_TimeZoneInfo As utc_TIME_ZONE_INFORMATION
            Dim utc_UtcDate As utc_SYSTEMTIME
            
            utc_GetTimeZoneInformation utc_TimeZoneInfo
            utc_TzSpecificLocalTimeToSystemTime utc_TimeZoneInfo, utc_DateToSystemTime(utc_LocalDate), utc_UtcDate
            
            ConvertToUtc = utc_SystemTimeToDate(utc_UtcDate)
        #End If
        
        Exit Function
        
utc_ErrorHandling:
        ERR.Raise 10012, "UtcConverter.ConvertToUtc", "UTC conversion error: " & ERR.Number & " - " & ERR.Description
    End Function
    
    ''
    ' Parse ISO 8601 date string to local date
    '
    ' @method ParseIso
    ' @param {Date} utc_IsoString
    ' @return {Date} Local date
    ' @throws 10013 - ISO 8601 parsing error
    ''
    Public Function ParseIso(utc_IsoString As String) As Date
        On Error GoTo utc_ErrorHandling
        
        Dim utc_Parts() As String
        Dim utc_DateParts() As String
        Dim utc_TimeParts() As String
        Dim utc_OffsetIndex As Long
        Dim utc_HasOffset As Boolean
        Dim utc_NegativeOffset As Boolean
        Dim utc_OffsetParts() As String
        Dim utc_Offset As Date
        
        utc_Parts = VBA.Split(utc_IsoString, "T")
        utc_DateParts = VBA.Split(utc_Parts(0), "-")
        ParseIso = VBA.DateSerial(VBA.CInt(utc_DateParts(0)), VBA.CInt(utc_DateParts(1)), VBA.CInt(utc_DateParts(2)))
        
        If UBound(utc_Parts) > 0 Then
            If VBA.InStr(utc_Parts(1), "Z") Then
                utc_TimeParts = VBA.Split(VBA.Replace(utc_Parts(1), "Z", ""), ":")
            Else
                utc_OffsetIndex = VBA.InStr(1, utc_Parts(1), "+")
                If utc_OffsetIndex = 0 Then
                    utc_NegativeOffset = True
                    utc_OffsetIndex = VBA.InStr(1, utc_Parts(1), "-")
                End If
                
                If utc_OffsetIndex > 0 Then
                    utc_HasOffset = True
                    utc_TimeParts = VBA.Split(VBA.Left$(utc_Parts(1), utc_OffsetIndex - 1), ":")
                    utc_OffsetParts = VBA.Split(VBA.Right$(utc_Parts(1), Len(utc_Parts(1)) - utc_OffsetIndex), ":")
                    
                    Select Case UBound(utc_OffsetParts)
                        Case 0
                            utc_Offset = TimeSerial(VBA.CInt(utc_OffsetParts(0)), 0, 0)
                        Case 1
                            utc_Offset = TimeSerial(VBA.CInt(utc_OffsetParts(0)), VBA.CInt(utc_OffsetParts(1)), 0)
                        Case 2
                            ' VBA.Val does not use regional settings, use for seconds to avoid decimal/comma issues
                            utc_Offset = TimeSerial(VBA.CInt(utc_OffsetParts(0)), VBA.CInt(utc_OffsetParts(1)), Int(VBA.Val(utc_OffsetParts(2))))
                    End Select
                    
                    If utc_NegativeOffset Then: utc_Offset = -utc_Offset
                Else
                    utc_TimeParts = VBA.Split(utc_Parts(1), ":")
                End If
            End If
            
            Select Case UBound(utc_TimeParts)
                Case 0
                    ParseIso = ParseIso + VBA.TimeSerial(VBA.CInt(utc_TimeParts(0)), 0, 0)
                Case 1
                    ParseIso = ParseIso + VBA.TimeSerial(VBA.CInt(utc_TimeParts(0)), VBA.CInt(utc_TimeParts(1)), 0)
                Case 2
                    ' VBA.Val does not use regional settings, use for seconds to avoid decimal/comma issues
                    ParseIso = ParseIso + VBA.TimeSerial(VBA.CInt(utc_TimeParts(0)), VBA.CInt(utc_TimeParts(1)), Int(VBA.Val(utc_TimeParts(2))))
            End Select
            
            ParseIso = ParseUtc(ParseIso)
            
            If utc_HasOffset Then
                ParseIso = ParseIso - utc_Offset
            End If
        End If
        
        Exit Function
        
utc_ErrorHandling:
        ERR.Raise 10013, "UtcConverter.ParseIso", "ISO 8601 parsing error for " & utc_IsoString & ": " & ERR.Number & " - " & ERR.Description
    End Function
    
    ''
    ' Convert local date to ISO 8601 string
    '
    ' @method ConvertToIso
    ' @param {Date} utc_LocalDate
    ' @return {Date} ISO 8601 string
    ' @throws 10014 - ISO 8601 conversion error
    ''
    Public Function ConvertToIso(utc_LocalDate As Date) As String
        On Error GoTo utc_ErrorHandling
        
        ConvertToIso = VBA.Format$(ConvertToUtc(utc_LocalDate), "yyyy-mm-ddTHH:mm:ss.000Z")
        
        Exit Function
        
utc_ErrorHandling:
        ERR.Raise 10014, "UtcConverter.ConvertToIso", "ISO 8601 conversion error: " & ERR.Number & " - " & ERR.Description
    End Function
    
    ' ============================================= '
    ' Private Functions
    ' ============================================= '
    
    #If Mac Then
        
        Private Function utc_ConvertDate(utc_Value As Date, Optional utc_ConvertToUtc As Boolean = False) As Date
            Dim utc_ShellCommand As String
            Dim utc_Result As utc_ShellResult
            Dim utc_Parts() As String
            Dim utc_DateParts() As String
            Dim utc_TimeParts() As String
            
            If utc_ConvertToUtc Then
                utc_ShellCommand = "date -ur `date -jf '%Y-%m-%d %H:%M:%S' " & _
                    "'" & VBA.Format$(utc_Value, "yyyy-mm-dd HH:mm:ss") & "' " & _
                    " +'%s'` +'%Y-%m-%d %H:%M:%S'"
            Else
                utc_ShellCommand = "date -jf '%Y-%m-%d %H:%M:%S %z' " & _
                    "'" & VBA.Format$(utc_Value, "yyyy-mm-dd HH:mm:ss") & " +0000' " & _
                    "+'%Y-%m-%d %H:%M:%S'"
            End If
            
            utc_Result = utc_ExecuteInShell(utc_ShellCommand)
            
            If utc_Result.utc_Output = "" Then
                ERR.Raise 10015, "UtcConverter.utc_ConvertDate", "'date' command failed"
            Else
                utc_Parts = Split(utc_Result.utc_Output, " ")
                utc_DateParts = Split(utc_Parts(0), "-")
                utc_TimeParts = Split(utc_Parts(1), ":")
                
                utc_ConvertDate = DateSerial(utc_DateParts(0), utc_DateParts(1), utc_DateParts(2)) + _
                    TimeSerial(utc_TimeParts(0), utc_TimeParts(1), utc_TimeParts(2))
            End If
        End Function
        
        Private Function utc_ExecuteInShell(utc_ShellCommand As String) As utc_ShellResult
            #If VBA7 Then
                Dim utc_File As LongPtr
                Dim utc_Read As LongPtr
            #Else
                Dim utc_File As Long
                Dim utc_Read As Long
            #End If
            
            Dim utc_Chunk As String
            
            On Error GoTo utc_ErrorHandling
            utc_File = utc_popen(utc_ShellCommand, "r")
            
            If utc_File = 0 Then: Exit Function
            
            Do While utc_feof(utc_File) = 0
                utc_Chunk = VBA.Space$(50)
                utc_Read = CLng(utc_fread(utc_Chunk, 1, Len(utc_Chunk) - 1, utc_File))
                If utc_Read > 0 Then
                    utc_Chunk = VBA.Left$(utc_Chunk, CLng(utc_Read))
                    utc_ExecuteInShell.utc_Output = utc_ExecuteInShell.utc_Output & utc_Chunk
                End If
            Loop
            
utc_ErrorHandling:
            utc_ExecuteInShell.utc_ExitCode = CLng(utc_pclose(utc_File))
        End Function
        
    #Else
        
        Private Function utc_DateToSystemTime(utc_Value As Date) As utc_SYSTEMTIME
            utc_DateToSystemTime.utc_wYear = VBA.year(utc_Value)
            utc_DateToSystemTime.utc_wMonth = VBA.month(utc_Value)
            utc_DateToSystemTime.utc_wDay = VBA.Day(utc_Value)
            utc_DateToSystemTime.utc_wHour = VBA.Hour(utc_Value)
            utc_DateToSystemTime.utc_wMinute = VBA.Minute(utc_Value)
            utc_DateToSystemTime.utc_wSecond = VBA.Second(utc_Value)
            utc_DateToSystemTime.utc_wMilliseconds = 0
        End Function
        
        Private Function utc_SystemTimeToDate(utc_Value As utc_SYSTEMTIME) As Date
            utc_SystemTimeToDate = DateSerial(utc_Value.utc_wYear, utc_Value.utc_wMonth, utc_Value.utc_wDay) + _
                TimeSerial(utc_Value.utc_wHour, utc_Value.utc_wMinute, utc_Value.utc_wSecond)
        End Function
        
    #End If
```

## j页面设置

```VBA
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

## k访谈提纲

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

## 更新

```VBA
Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Sub 获取标签名(control As IRibbonControl, ByRef returnedVal)
    returnedVal = "报告小帮手 V2.6"
End Sub
Sub 获取标签日期(control As IRibbonControl, ByRef returnedVal)
    returnedVal = "20220814更新"
End Sub
Sub 签名(control As IRibbonControl, ByRef returnedVal)
    returnedVal = "公众号：茶瓜子的休闲馆"
End Sub
Sub 检查更新(control As IRibbonControl)
    本地 = Val(ThisVersion)
    最新 = Val(Getver)
    If 本地 <> 最新 Then
        y = MsgBox("存在新版本，是否进入主页查看最新版？", vbYesNo)
        If y = 6 Then
            OpenWeb
        End If
    Else
        MsgBox "当前版本为最新版"
    End If
End Sub
Public Function ThisVersion()
    ThisVersion = "2.6"
End Function
Public Function Getver()
    Dim Json As Object
    URL = "http://api.gzaudit.com/xbs/wd/"
    res = GetData(URL, "UTF-8")
    Set Json = JsonConverter.ParseJson(res)
    Getver = Json("版本")
End Function
Sub OpenWeb()
  ShellExecute 0&, vbNullString, "www.gzaudit.com", vbNullString, vbNullString, vbNormalFocus
End Sub
Function GetData(StrUrl, CodePageX)
    Dim oHtml As Object
    Set oHtml = VBA.CreateObject("WinHttp.WinHttpRequest.5.1")
    Dim sUrl As String
    sUrl = StrUrl
    Dim sCharset As String
    sCharset = CodePageX
    With oHtml
        .Open "GET", sUrl, False
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
        .Send
        '获取返回的字节数组
        bResult = .ResponseBody
        '按照指定的字符编码显示
        sResult = BytesToStr(bResult, CodePageX)
        'Debug.Print sResult
    End With
    GetData = sResult
    Set oHtml = Nothing
End Function
Public Function BytesToStr(strBody, CodeBase)
    Dim objStream
    Set objStream = CreateObject("Adodb.Stream")
    
    With objStream
        .Type = 1
        .Mode = 3
        .Open
        .Write strBody
        .Position = 0
        .Type = 2
        .Charset = CodeBase '"GB2312" '
        BytesToStr = .ReadText
        .Close
    End With
    Set objStream = Nothing
End Function
```

## 类模块

```VBA
''
' Dictionary v1.2.0
' (c) Tim Hall - https://github.com/timhall/VBA-Dictionary
'
' Drop-in replacement for Scripting.Dictionary on Mac
'
' @author: tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '

#Const UseScriptingDictionaryIfAvailable = True

#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    
    ' KeyValue 0: FormattedKey, 1: OriginalKey, 2: Value
    Private pKeyValues As Collection
    Private pKeys() As Variant
    Private pItems() As Variant
    Private pCompareMode As CompareMethod
    
#Else
    
    Private pDictionary As Object
    
#End If

' --------------------------------------------- '
' Types
' --------------------------------------------- '

Public Enum CompareMethod
    BinaryCompare = vbBinaryCompare
    TextCompare = vbTextCompare
    DatabaseCompare = vbDatabaseCompare
End Enum

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

Public Property Get CompareMode() As CompareMethod
    #If Mac Or Not UseScriptingDictionaryIfAvailable Then
        CompareMode = pCompareMode
    #Else
        CompareMode = pDictionary.CompareMode
    #End If
End Property
Public Property Let CompareMode(Value As CompareMethod)
    #If Mac Or Not UseScriptingDictionaryIfAvailable Then
        If Me.Count > 0 Then
            ' Can't change CompareMode for Dictionary that contains data
            ' http://msdn.microsoft.com/en-us/library/office/gg278481(v=office.15).aspx
            ERR.Raise 5 ' Invalid procedure call or argument
        End If
        
        pCompareMode = Value
    #Else
        pDictionary.CompareMode = Value
    #End If
End Property

Public Property Get Count() As Long
    #If Mac Or Not UseScriptingDictionaryIfAvailable Then
        Count = pKeyValues.Count
    #Else
        Count = pDictionary.Count
    #End If
End Property

Public Property Get Item(Key As Variant) As Variant
    #If Mac Or Not UseScriptingDictionaryIfAvailable Then
        Dim KeyValue As Variant
        KeyValue = GetKeyValue(Key)
        
        If Not IsEmpty(KeyValue) Then
            If IsObject(KeyValue(2)) Then
                Set Item = KeyValue(2)
            Else
                Item = KeyValue(2)
            End If
        Else
            ' Not found -> Returns Empty
        End If
    #Else
        If IsObject(pDictionary.Item(Key)) Then
            Set Item = pDictionary.Item(Key)
        Else
            Item = pDictionary.Item(Key)
        End If
    #End If
End Property
Public Property Let Item(Key As Variant, Value As Variant)
    #If Mac Or Not UseScriptingDictionaryIfAvailable Then
        If Me.Exists(Key) Then
            ReplaceKeyValue GetKeyValue(Key), Key, Value
        Else
            AddKeyValue Key, Value
        End If
    #Else
        pDictionary.Item(Key) = Value
    #End If
End Property
Public Property Set Item(Key As Variant, Value As Variant)
    #If Mac Or Not UseScriptingDictionaryIfAvailable Then
        If Me.Exists(Key) Then
            ReplaceKeyValue GetKeyValue(Key), Key, Value
        Else
            AddKeyValue Key, Value
        End If
    #Else
        Set pDictionary.Item(Key) = Value
    #End If
End Property

Public Property Let Key(Previous As Variant, Updated As Variant)
    #If Mac Or Not UseScriptingDictionaryIfAvailable Then
        Dim KeyValue As Variant
        KeyValue = GetKeyValue(Previous)
        
        If Not IsEmpty(KeyValue) Then
            ReplaceKeyValue KeyValue, Updated, KeyValue(2)
        End If
    #Else
        pDictionary.Key(Previous) = Updated
    #End If
End Property

' ============================================= '
' Public Methods
' ============================================= '

''
' Add an item with the given key
'
' @param {Variant} Key
' @param {Variant} Item
' --------------------------------------------- '
Public Sub Add(Key As Variant, Item As Variant)
    #If Mac Or Not UseScriptingDictionaryIfAvailable Then
        If Not Me.Exists(Key) Then
            AddKeyValue Key, Item
        Else
            ' This key is already associated with an element of this collection
            ERR.Raise 457
        End If
    #Else
        pDictionary.Add Key, Item
    #End If
End Sub

''
' Check if an item exists for the given key
'
' @param {Variant} Key
' @return {Boolean}
' --------------------------------------------- '
Public Function Exists(Key As Variant) As Boolean
    #If Mac Or Not UseScriptingDictionaryIfAvailable Then
        Exists = Not IsEmpty(GetKeyValue(Key))
    #Else
        Exists = pDictionary.Exists(Key)
    #End If
End Function

''
' Get an array of all items
'
' @return {Variant}
' --------------------------------------------- '
Public Function Items() As Variant
    #If Mac Or Not UseScriptingDictionaryIfAvailable Then
        If Me.Count > 0 Then
            Items = pItems
        Else
            ' Split("") creates initialized empty array that matches Dictionary Keys and Items
            Items = Split("")
        End If
    #Else
        Items = pDictionary.Items
    #End If
End Function

''
' Get an array of all keys
'
' @return {Variant}
' --------------------------------------------- '
Public Function Keys() As Variant
    #If Mac Or Not UseScriptingDictionaryIfAvailable Then
        If Me.Count > 0 Then
            Keys = pKeys
        Else
            ' Split("") creates initialized empty array that matches Dictionary Keys and Items
            Keys = Split("")
        End If
    #Else
        Keys = pDictionary.Keys
    #End If
End Function

''
' Remove an item for the given key
'
' @param {Variant} Key
' --------------------------------------------- '
Public Sub Remove(Key As Variant)
    #If Mac Or Not UseScriptingDictionaryIfAvailable Then
        Dim KeyValue As Variant
        KeyValue = GetKeyValue(Key)
        
        If Not IsEmpty(KeyValue) Then
            RemoveKeyValue KeyValue
        Else
            ' Application-defined or object-defined error
            ERR.Raise 32811
        End If
    #Else
        pDictionary.Remove Key
    #End If
End Sub

''
' Remove all items
' --------------------------------------------- '
Public Sub RemoveAll()
    #If Mac Or Not UseScriptingDictionaryIfAvailable Then
        Set pKeyValues = New Collection
        
        Erase pKeys
        Erase pItems
    #Else
        pDictionary.RemoveAll
    #End If
End Sub

' ============================================= '
' Private Functions
' ============================================= '

#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    
    Private Function GetKeyValue(Key As Variant) As Variant
        On Error Resume Next
        GetKeyValue = pKeyValues(GetFormattedKey(Key))
        ERR.Clear
    End Function
    
    Private Sub AddKeyValue(Key As Variant, Value As Variant, Optional Index As Long = -1)
        If Me.Count = 0 Then
            ReDim pKeys(0 To 0)
            ReDim pItems(0 To 0)
        Else
            ReDim Preserve pKeys(0 To UBound(pKeys) + 1)
            ReDim Preserve pItems(0 To UBound(pItems) + 1)
        End If
        
        Dim FormattedKey As String
        FormattedKey = GetFormattedKey(Key)
        
        If Index > 0 And Index <= pKeyValues.Count Then
            Dim i As Long
            For i = UBound(pKeys) To Index Step -1
                pKeys(i) = pKeys(i - 1)
                If IsObject(pItems(i - 1)) Then
                    Set pItems(i) = pItems(i - 1)
                Else
                    pItems(i) = pItems(i - 1)
                End If
            Next i
            
            pKeys(Index - 1) = Key
            If IsObject(Value) Then
                Set pItems(Index - 1) = Value
            Else
                pItems(Index - 1) = Value
            End If
            
            pKeyValues.Add Array(FormattedKey, Key, Value), FormattedKey, before:=Index
        Else
            pKeys(UBound(pKeys)) = Key
            If IsObject(Value) Then
                Set pItems(UBound(pItems)) = Value
            Else
                pItems(UBound(pItems)) = Value
            End If
            
            pKeyValues.Add Array(FormattedKey, Key, Value), FormattedKey
        End If
    End Sub
    
    Private Sub ReplaceKeyValue(KeyValue As Variant, Key As Variant, Value As Variant)
        Dim Index As Long
        Dim i As Integer
        
        For i = 0 To UBound(pKeys)
            If pKeys(i) = KeyValue(1) Then
                Index = i + 1
                Exit For
            End If
        Next i
        
        ' Remove existing value
        RemoveKeyValue KeyValue, Index
        
        ' Add new key value back
        AddKeyValue Key, Value, Index
    End Sub
    
    Private Sub RemoveKeyValue(KeyValue As Variant, Optional ByVal Index As Long = -1)
        Dim i As Long
        If Index = -1 Then
            For i = 0 To UBound(pKeys)
                If pKeys(i) = KeyValue(1) Then
                    Index = i
                End If
            Next i
        Else
            Index = Index - 1
        End If
        
        If Index >= 0 And Index <= UBound(pKeys) Then
            For i = Index To UBound(pKeys) - 1
                pKeys(i) = pKeys(i + 1)
                
                If IsObject(pItems(i + 1)) Then
                    Set pItems(i) = pItems(i + 1)
                Else
                    pItems(i) = pItems(i + 1)
                End If
            Next i
            
            If UBound(pKeys) = 0 Then
                Erase pKeys
                Erase pItems
            Else
                ReDim Preserve pKeys(0 To UBound(pKeys) - 1)
                ReDim Preserve pItems(0 To UBound(pItems) - 1)
            End If
        End If
        
        pKeyValues.Remove KeyValue(0)
    End Sub
    
    Private Function GetFormattedKey(Key As Variant) As String
        GetFormattedKey = CStr(Key)
        If Me.CompareMode = CompareMethod.BinaryCompare Then
            ' Collection does not have method of setting key comparison
            ' So case-sensitive keys aren't supported by default
            ' -> Approach: Append lowercase characters to original key
            '    AbC -> AbC__b, abc -> abc__abc, ABC -> ABC
            '    Won't work in very strange cases, but should work for now
            '    AbBb -> AbBb__bb matches AbbB -> AbbB__bb
            Dim Lowercase As String
            Lowercase = ""
            
            Dim i As Integer
            Dim Ascii As Integer
            Dim Char As String
            For i = 1 To Len(GetFormattedKey)
                Char = VBA.Mid$(GetFormattedKey, i, 1)
                Ascii = Asc(Char)
                If Ascii >= 97 And Ascii <= 122 Then
                    Lowercase = Lowercase & Char
                End If
            Next i
            
            If Lowercase <> "" Then
                GetFormattedKey = GetFormattedKey & "__" & Lowercase
            End If
        End If
    End Function
    
#End If

Private Sub Class_Initialize()
    #If Mac Or Not UseScriptingDictionaryIfAvailable Then
        Set pKeyValues = New Collection
        
        Erase pKeys
        Erase pItems
    #Else
        Set pDictionary = CreateObject("Scripting.Dictionary")
    #End If
End Sub

Private Sub Class_Terminate()
    #If Mac Or Not UseScriptingDictionaryIfAvailable Then
        Set pKeyValues = Nothing
    #Else
        Set pDictionary = Nothing
    #End If
End Sub
```







---

> 作者: [richfan](https://richfan.site/)  
> URL: http://richfan.site/posts/program/vba/%E5%BA%95%E7%A8%BF%E5%B0%8F%E5%B8%AE%E6%89%8B%E4%BB%A3%E7%A0%81/%E5%BA%95%E7%A8%BF%E5%B0%8F%E5%B8%AE%E6%89%8B%E4%BB%A3%E7%A0%81/  

