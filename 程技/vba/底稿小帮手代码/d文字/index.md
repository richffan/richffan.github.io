# d文字


## d文字

&lt;!--more--&gt;

```vb
Sub 宋体宋体(control As IRibbonControl)
    &#39;选中范围字体为宋体&#43;宋体
    With Selection.Font
        .NameFarEast = &#34;宋体&#34;
        .NameAscii = &#34;宋体&#34;
        .NameOther = &#34;宋体&#34;
    End With
End Sub
Sub 宋体罗马(control As IRibbonControl) &#39;文字-宋体罗马
    &#39;选中范围字体为宋体&#43;Times
    With Selection.Font
        .NameFarEast = &#34;仿宋&#34;
        .NameAscii = &#34;Times New Roman&#34;
        .NameOther = &#34;Times New Roman&#34;
    End With
End Sub
Sub 楷体加粗(control As IRibbonControl) &#39;文字-楷体加粗
    &#39;选中范围字体为楷体加粗
    With Selection.Font
        .NameFarEast = &#34;楷体&#34;
        .NameAscii = &#34;楷体&#34;
        .NameOther = &#34;楷体&#34;
        .Name = &#34;楷体&#34;
        .Bold = True
    End With
End Sub
Sub 去除空白(control As IRibbonControl) &#39;文字-去除空白
    &#39;删除换行及空格
    
    Selection.Find.ClearFormatting   &#39;删除空格
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = &#34; &#34;
        .Replacement.Text = &#34;&#34;
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
    
    Selection.Find.ClearFormatting   &#39;删除大空格
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = &#34;　&#34;
        .Replacement.Text = &#34;&#34;
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
    
    Selection.Find.Replacement.ClearFormatting    &#39;删除连续两个回车
    With Selection.Find
        .Text = &#34;^p^p&#34;
        .Replacement.Text = &#34;^p&#34;
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
Sub 中英文标点互换(control As IRibbonControl) &#39; 文字-中英文标点互换
    Dim ChineseInterpunction() As Variant, EnglishInterpunction() As Variant
    Dim myArray1() As Variant, myArray2() As Variant, strFind As String, strRep As String
    Dim msgResult As VbMsgBoxResult, n As Byte
    &#39; 定义一个中文标点的数组对象
    ChineseInterpunction = Array(&#34;，&#34;, &#34;。&#34;, &#34;；&#34;, &#34;：&#34;, &#34;？&#34;, &#34;！&#34;, &#34;……&#34;, &#34;—&#34;, &#34;～&#34;, &#34;（&#34;, &#34;）&#34;, &#34;《&#34;, &#34;》&#34;)
    &#39; 定义一个英文标点的数组对象
    EnglishInterpunction = Array(&#34;,&#34;, &#34;.&#34;, &#34;;&#34;, &#34;:&#34;, &#34;?&#34;, &#34;!&#34;, &#34;…&#34;, &#34;-&#34;, &#34;~&#34;, &#34;(&#34;, &#34;)&#34;, &#34;&amp;lt;&#34;, &#34;&amp;gt;&#34;)
    &#39; 提示用户交互的 MSGBOX 对话框
    msgResult = MsgBox(&#34;您想中英标点互换吗？按 Y 将中文标点转为英文标点，按 N 将英文标点转为中文标点！&#34;, vbYesNoCancel)
    Select Case msgResult
        Case vbCancel
            Exit Sub &#39; 如果用户选择了取消按钮，则退出程序运行
        Case vbYes &#39; 如果用户选择了 YES, 则将中文标点转换为英文标点
            myArray1 = ChineseInterpunction
            myArray2 = EnglishInterpunction
            strFind = &#34;“(*)”&#34;
            strRep = &#34;&#34;&#34;\1&#34;&#34;&#34;
        Case vbNo &#39; 如果用户选择了 NO, 则将英文标点转换为中文标点
            myArray1 = EnglishInterpunction
            myArray2 = ChineseInterpunction
            strFind = &#34;&#34;&#34;(*)&#34;&#34;&#34;
            strRep = &#34;“\1”&#34;
    End Select
    Application.ScreenUpdating = False &#39; 关闭屏幕更新
    For n = 0 To UBound(ChineseInterpunction)  &#39; 从数组的下标到上标间作一个循环
        With Selection.Find
            .ClearFormatting &#39; 不限定查找格式
            .MatchWildcards = False &#39; 不使用通配符
            &#39; 查找相应的英文标点，替换为对应的中文标点
            .Execute findtext:=myArray1(n), replacewith:=myArray2(n), Replace:=wdReplaceAll
        End With
    Next
    With Selection.Find
        .ClearFormatting &#39; 不限定查找格式
        .MatchWildcards = True &#39; 使用通配符
        .Execute findtext:=strFind, replacewith:=strRep, Replace:=wdReplaceAll
    End With
    Application.ScreenUpdating = True &#39; 恢复屏幕更新
End Sub
Sub 高亮(control As IRibbonControl) &#39;文字-HighLight
    If Selection.Range.HighlightColorIndex = 0 Then
        Selection.Range.HighlightColorIndex = wdYellow
    Else
        Selection.Range.HighlightColorIndex = wdNoHighlight
    End If
End Sub
```



---

> 作者:   
> URL: http://richfan.site/%E7%A8%8B%E6%8A%80/vba/%E5%BA%95%E7%A8%BF%E5%B0%8F%E5%B8%AE%E6%89%8B%E4%BB%A3%E7%A0%81/d%E6%96%87%E5%AD%97/  

