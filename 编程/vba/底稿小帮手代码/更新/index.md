# 更新


## 更新

<!--more-->

```vb
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


---

> 作者: [richfan](https://richfan.site/)  
> URL: http://richfan.site/%E7%BC%96%E7%A8%8B/vba/%E5%BA%95%E7%A8%BF%E5%B0%8F%E5%B8%AE%E6%89%8B%E4%BB%A3%E7%A0%81/%E6%9B%B4%E6%96%B0/  

