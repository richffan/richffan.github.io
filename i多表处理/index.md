# i多表处理


## i多表处理

<!--more-->

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



---

> 作者: [richfan](https://richfan.site/)  
> URL: http://richfan.site/i%E5%A4%9A%E8%A1%A8%E5%A4%84%E7%90%86/  

