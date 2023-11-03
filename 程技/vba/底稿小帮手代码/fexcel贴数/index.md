# fExcel贴数


## fExcel贴数


<!--more-->

```vb
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



---

> 作者: [richfan](https://richfan.site/)  
> URL: http://richfan.site/%E7%A8%8B%E6%8A%80/vba/%E5%BA%95%E7%A8%BF%E5%B0%8F%E5%B8%AE%E6%89%8B%E4%BB%A3%E7%A0%81/fexcel%E8%B4%B4%E6%95%B0/  

