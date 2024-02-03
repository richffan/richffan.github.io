# b表单处理



## b表单处理


&lt;!--more--&gt;

```vb
Dim aarr(1 To 20), bbrr(1 To 30, 1 To 30) &#39;多列调整
Sub 单表_一键调整(control As IRibbonControl) &#39;单表-格式
    &#39;功能：光标在表格中处理当前表格；否则处理所有表格！
    Application.ScreenUpdating = False  &#39;关闭屏幕刷新
    Application.DisplayAlerts = False  &#39;关闭提示
    On Error Resume Next  &#39;忽略错误
    &#39;-------------------------------------------------------------------------
    Dim mytable As Table, i As Long
    For Each mytable In Selection.Tables
        With mytable
            .Shading.ForegroundPatternColor = wdColorAutomatic
            .Shading.BackgroundPatternColor = wdColorAutomatic
            Options.DefaultHighlightColorIndex = wdNoHighlight
            .Range.HighlightColorIndex = wdNoHighlight
            .Style = &#34;表格主题&#34;
            With .Borders(wdBorderLeft)    &#39;左框线
                .LineStyle = wdLineStyleSingle   &#39;设置线条样式    不需要线条则填wdLineStyleNone
                .LineWidth = wdLineWidth150pt    &#39;宽度为1.5
            End With
            With .Borders(wdBorderRight)   &#39;右框线
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth150pt
            End With
            With .Borders(wdBorderTop)     &#39;上框线
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth150pt
            End With
            With .Borders(wdBorderBottom)  &#39;下框线
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth150pt
            End With
            With .Borders(wdBorderVertical)   &#39;内部纵向框线
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth050pt
            End With
            With .Borders(wdBorderHorizontal)   &#39;内部横向框线
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth050pt
            End With
            .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone  &#39;左上的斜线
            .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone    &#39;右上的斜线
            &#39;单元格边距
            .TopPadding = CentimetersToPoints(0) &#39;设置上边距为0
            .BottomPadding = CentimetersToPoints(0) &#39;设置下边距为0
            .LeftPadding = PixelsToPoints(0, True)  &#39;设置左边距为0
            .RightPadding = PixelsToPoints(0, True) &#39;设置右边距为0
            .Spacing = PixelsToPoints(0, True) &#39;允许单元格间距为0
            .AllowPageBreaks = True &#39;允许断页
            &#39;.AllowAutoFit = True &#39;允许自动重调尺寸
            With .Rows
                .WrapAroundText = False &#39;取消文字环绕
                &#39;.Alignment = wdAlignRowCenter &#39;表水平居中  wdAlignRowLeft &#39;左对齐
                .AllowBreakAcrossPages = False &#39;不允许行断页
                .Height = CentimetersToPoints(0.8) &#39;行高0.8
                .HeightRule = wdRowHeightAtLeast &#39;行高设为最小值
                .LeftIndent = CentimetersToPoints(0) &#39;左面缩进量为0
            End With
            With .Range
                With .Font &#39;字体格式
                    .NameFarEast = &#34;宋体&#34;
                    .NameAscii = &#34;Times New Roman&#34;
                    .NameOther = &#34;Times New Roman&#34;
                    .Color = wdColorAutomatic &#39;自动字体颜色
                    .Size = 10.5   &#39;字号
                    .Kerning = 0
                    .DisableCharacterSpaceGrid = True  &#39;选定段落中的字符与行网格对齐
                End With
                With .ParagraphFormat &#39;段落格式
                    .LineUnitBefore = 0
                    .LineUnitAfter = 0
                    .SpaceBefore = 0
                    .SpaceAfter = 0
                    .CharacterUnitFirstLineIndent = 0 &#39;取消首行缩进
                    .FirstLineIndent = CentimetersToPoints(0) &#39;取消首行缩进
                    .LineSpacingRule = wdLineSpaceSingle &#39;wdLineSpaceSingle &#39;单倍行距  wdLineSpaceExactly &#39;行距固定值
                    &#39;&#39;.LineSpacing = 18 &#39;设置行间距为18磅，配合行距固定值
                    &#39;.Alignment = wdAlignParagraphCenter &#39;单元格水平居中
                    .AutoAdjustRightIndent = False  &#39;自动调整所选段落的右缩进
                    .DisableLineHeightGrid = True   &#39;选定段落中的字符与行网格对齐
                End With
                .Cells.VerticalAlignment = wdCellAlignVerticalCenter  &#39;单元格垂直居中
            End With
            For Each cl In .Range.Cells    &#39;文字靠左，数字靠右,合计居中,序号居中
                Acell = ActiveDocument.Range(cl.Range.Start, cl.Range.End - 1).Text &#39;提取文本
                If IsNumeric(Acell) Then
                    cl.Range.ParagraphFormat.Alignment = wdAlignParagraphRight    &#39;右对齐
                Else
                    cl.Range.ParagraphFormat.Alignment = wdAlignParagraphJustify   &#39;左对齐
                    If Acell = &#34;合计&#34; Or Acell = &#34;总计&#34; Or Acell = &#34;总 计&#34; Or Acell = &#34;合 计&#34; Then
                        cl.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter    &#39;水平居中
                        If cl.ColumnIndex = .Columns.Count Then
                            .Columns(cl.ColumnIndex).Select
                            Selection.Font.Bold = True
                        Else
                            cl.Row.Range.Font.Bold = True
                        End If
                    ElseIf Acell = &#34;序号&#34; Or Acell = &#34;序 号&#34; Then
                        .Columns(cl.ColumnIndex).Select
                        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter    &#39;水平居中
                    End If
                End If
            Next
            &#39;设置首行格式
            .Rows(1).Select &#39; 选中第一个单元格
            With Selection
                .Rows.HeadingFormat = wdToggle &#39;自动标题行重复
                .ParagraphFormat.Alignment = wdAlignParagraphCenter   &#39;水平居中
                .Range.Font.Bold = True &#39;表头加粗黑体
                .Shading.ForegroundPatternColor = wdColorAutomatic &#39;首行自动颜色
                .Shading.BackgroundPatternColor = -603923969 &#39;首行底纹填充,不要底色则删了这行
                &#39;.Borders(wdBorderBottom).LineStyle = xlContinuous
                &#39;.Borders(wdBorderBottom).LineWidth = wdLineWidth50pt
            End With
            &#39;自动调整表格
            .Columns.PreferredWidthType = wdPreferredWidthAuto
            .AutoFitBehavior (wdAutoFitContent) &#39;根据内容调整表格
            .AutoFitBehavior (wdAutoFitWindow) &#39;根据窗口调整表格
        End With
    Next
    &#39;---------------------------------------------------------------------------------------
    ERR.Clear: On Error GoTo 0 &#39;恢复错误捕捉
    Application.DisplayAlerts = True  &#39;开启提示
    Application.ScreenUpdating = True   &#39;开启屏幕刷新
End Sub
Sub 格宽调整_释放(control As IRibbonControl) &#39;列宽调整-多列加载
    Set mytable = Selection.Tables(1)
    For i = 1 To mytable.Rows.Count
        For j = 1 To mytable.Rows(i).Cells.Count
            mytable.Rows(i).Cells(j).Width = bbrr(i, j)
        Next j
    Next i
End Sub
Sub 格宽调整_读取(control As IRibbonControl)  &#39;列宽调整-多列读取
    Set mytable = Selection.Tables(1)
    mytable.AutoFitBehavior (wdAutoFitFixed)
    For i = 1 To mytable.Rows.Count
        For j = 1 To mytable.Rows(i).Cells.Count
            bbrr(i, j) = mytable.Rows(i).Cells(j).Width
        Next j
    Next i
End Sub
Sub 列宽调整_读取(control As IRibbonControl)   &#39;列宽调整-单列读取
    With Selection.Tables(1)
        ColumnsCounts = .Columns.Count
        For i = 1 To ColumnsCounts
            aarr(i) = .Columns(i).Width
        Next
    End With
End Sub
Sub 列宽调整_释放(control As IRibbonControl) &#39;列宽调整-单列加载
    With Selection.Tables(1)
        .AutoFitBehavior (wdAutoFitFixed)
        ColumnsCounts = .Columns.Count
        For i = 1 To ColumnsCounts
            .Columns(i).Width = aarr(i)
        Next
    End With
End Sub
```




---

> 作者:   
> URL: http://richfan.site/%E7%A8%8B%E6%8A%80/vba/%E5%BA%95%E7%A8%BF%E5%B0%8F%E5%B8%AE%E6%89%8B%E4%BB%A3%E7%A0%81/b%E8%A1%A8%E5%8D%95%E5%A4%84%E7%90%86/  

