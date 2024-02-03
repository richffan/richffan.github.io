# 更新


## 更新

&lt;!--more--&gt;

```vb
Private Declare PtrSafe Function ShellExecute Lib &#34;shell32.dll&#34; Alias &#34;ShellExecuteA&#34; (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Sub 获取标签名(control As IRibbonControl, ByRef returnedVal)
    returnedVal = &#34;报告小帮手 V2.6&#34;
End Sub
Sub 获取标签日期(control As IRibbonControl, ByRef returnedVal)
    returnedVal = &#34;20220814更新&#34;
End Sub
Sub 签名(control As IRibbonControl, ByRef returnedVal)
    returnedVal = &#34;公众号：茶瓜子的休闲馆&#34;
End Sub
Sub 检查更新(control As IRibbonControl)
    本地 = Val(ThisVersion)
    最新 = Val(Getver)
    If 本地 &lt;&gt; 最新 Then
        y = MsgBox(&#34;存在新版本，是否进入主页查看最新版？&#34;, vbYesNo)
        If y = 6 Then
            OpenWeb
        End If
    Else
        MsgBox &#34;当前版本为最新版&#34;
    End If
End Sub
Public Function ThisVersion()
    ThisVersion = &#34;2.6&#34;
End Function
Public Function Getver()
    Dim Json As Object
    URL = &#34;http://api.gzaudit.com/xbs/wd/&#34;
    res = GetData(URL, &#34;UTF-8&#34;)
    Set Json = JsonConverter.ParseJson(res)
    Getver = Json(&#34;版本&#34;)
End Function
Sub OpenWeb()
  ShellExecute 0&amp;, vbNullString, &#34;www.gzaudit.com&#34;, vbNullString, vbNullString, vbNormalFocus
End Sub
Function GetData(StrUrl, CodePageX)
    Dim oHtml As Object
    Set oHtml = VBA.CreateObject(&#34;WinHttp.WinHttpRequest.5.1&#34;)
    Dim sUrl As String
    sUrl = StrUrl
    Dim sCharset As String
    sCharset = CodePageX
    With oHtml
        .Open &#34;GET&#34;, sUrl, False
        .setRequestHeader &#34;User-Agent&#34;, &#34;Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36&#34;
        .Send
        &#39;获取返回的字节数组
        bResult = .ResponseBody
        &#39;按照指定的字符编码显示
        sResult = BytesToStr(bResult, CodePageX)
        &#39;Debug.Print sResult
    End With
    GetData = sResult
    Set oHtml = Nothing
End Function
Public Function BytesToStr(strBody, CodeBase)
    Dim objStream
    Set objStream = CreateObject(&#34;Adodb.Stream&#34;)
    
    With objStream
        .Type = 1
        .Mode = 3
        .Open
        .Write strBody
        .Position = 0
        .Type = 2
        .Charset = CodeBase &#39;&#34;GB2312&#34; &#39;
        BytesToStr = .ReadText
        .Close
    End With
    Set objStream = Nothing
End Function
```


---

> 作者:   
> URL: http://richfan.site/%E7%A8%8B%E6%8A%80/vba/%E5%BA%95%E7%A8%BF%E5%B0%8F%E5%B8%AE%E6%89%8B%E4%BB%A3%E7%A0%81/%E6%9B%B4%E6%96%B0/  

