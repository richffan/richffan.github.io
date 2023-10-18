# d文字


## d文字

<!--more-->

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



---

> 作者: [richfan](https://richfan.site/)  
> URL: http://richfan.site/2023/05/d%E6%96%87%E5%AD%97/  

