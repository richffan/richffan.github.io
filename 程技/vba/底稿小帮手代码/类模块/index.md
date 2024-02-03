# 类模块



## 类模块

&lt;!--more--&gt;


```vb
&#39;&#39;
&#39; Dictionary v1.2.0
&#39; (c) Tim Hall - https://github.com/timhall/VBA-Dictionary
&#39;
&#39; Drop-in replacement for Scripting.Dictionary on Mac
&#39;
&#39; @author: tim.hall.engr@gmail.com
&#39; @license: MIT (http://www.opensource.org/licenses/mit-license.php
&#39;
&#39; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ &#39;
Option Explicit

&#39; --------------------------------------------- &#39;
&#39; Constants and Private Variables
&#39; --------------------------------------------- &#39;

#Const UseScriptingDictionaryIfAvailable = True

#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    
    &#39; KeyValue 0: FormattedKey, 1: OriginalKey, 2: Value
    Private pKeyValues As Collection
    Private pKeys() As Variant
    Private pItems() As Variant
    Private pCompareMode As CompareMethod
    
#Else
    
    Private pDictionary As Object
    
#End If

&#39; --------------------------------------------- &#39;
&#39; Types
&#39; --------------------------------------------- &#39;

Public Enum CompareMethod
    BinaryCompare = vbBinaryCompare
    TextCompare = vbTextCompare
    DatabaseCompare = vbDatabaseCompare
End Enum

&#39; --------------------------------------------- &#39;
&#39; Properties
&#39; --------------------------------------------- &#39;

Public Property Get CompareMode() As CompareMethod
    #If Mac Or Not UseScriptingDictionaryIfAvailable Then
        CompareMode = pCompareMode
    #Else
        CompareMode = pDictionary.CompareMode
    #End If
End Property
Public Property Let CompareMode(Value As CompareMethod)
    #If Mac Or Not UseScriptingDictionaryIfAvailable Then
        If Me.Count &gt; 0 Then
            &#39; Can&#39;t change CompareMode for Dictionary that contains data
            &#39; http://msdn.microsoft.com/en-us/library/office/gg278481(v=office.15).aspx
            ERR.Raise 5 &#39; Invalid procedure call or argument
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
            &#39; Not found -&gt; Returns Empty
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

&#39; ============================================= &#39;
&#39; Public Methods
&#39; ============================================= &#39;

&#39;&#39;
&#39; Add an item with the given key
&#39;
&#39; @param {Variant} Key
&#39; @param {Variant} Item
&#39; --------------------------------------------- &#39;
Public Sub Add(Key As Variant, Item As Variant)
    #If Mac Or Not UseScriptingDictionaryIfAvailable Then
        If Not Me.Exists(Key) Then
            AddKeyValue Key, Item
        Else
            &#39; This key is already associated with an element of this collection
            ERR.Raise 457
        End If
    #Else
        pDictionary.Add Key, Item
    #End If
End Sub

&#39;&#39;
&#39; Check if an item exists for the given key
&#39;
&#39; @param {Variant} Key
&#39; @return {Boolean}
&#39; --------------------------------------------- &#39;
Public Function Exists(Key As Variant) As Boolean
    #If Mac Or Not UseScriptingDictionaryIfAvailable Then
        Exists = Not IsEmpty(GetKeyValue(Key))
    #Else
        Exists = pDictionary.Exists(Key)
    #End If
End Function

&#39;&#39;
&#39; Get an array of all items
&#39;
&#39; @return {Variant}
&#39; --------------------------------------------- &#39;
Public Function Items() As Variant
    #If Mac Or Not UseScriptingDictionaryIfAvailable Then
        If Me.Count &gt; 0 Then
            Items = pItems
        Else
            &#39; Split(&#34;&#34;) creates initialized empty array that matches Dictionary Keys and Items
            Items = Split(&#34;&#34;)
        End If
    #Else
        Items = pDictionary.Items
    #End If
End Function

&#39;&#39;
&#39; Get an array of all keys
&#39;
&#39; @return {Variant}
&#39; --------------------------------------------- &#39;
Public Function Keys() As Variant
    #If Mac Or Not UseScriptingDictionaryIfAvailable Then
        If Me.Count &gt; 0 Then
            Keys = pKeys
        Else
            &#39; Split(&#34;&#34;) creates initialized empty array that matches Dictionary Keys and Items
            Keys = Split(&#34;&#34;)
        End If
    #Else
        Keys = pDictionary.Keys
    #End If
End Function

&#39;&#39;
&#39; Remove an item for the given key
&#39;
&#39; @param {Variant} Key
&#39; --------------------------------------------- &#39;
Public Sub Remove(Key As Variant)
    #If Mac Or Not UseScriptingDictionaryIfAvailable Then
        Dim KeyValue As Variant
        KeyValue = GetKeyValue(Key)
        
        If Not IsEmpty(KeyValue) Then
            RemoveKeyValue KeyValue
        Else
            &#39; Application-defined or object-defined error
            ERR.Raise 32811
        End If
    #Else
        pDictionary.Remove Key
    #End If
End Sub

&#39;&#39;
&#39; Remove all items
&#39; --------------------------------------------- &#39;
Public Sub RemoveAll()
    #If Mac Or Not UseScriptingDictionaryIfAvailable Then
        Set pKeyValues = New Collection
        
        Erase pKeys
        Erase pItems
    #Else
        pDictionary.RemoveAll
    #End If
End Sub

&#39; ============================================= &#39;
&#39; Private Functions
&#39; ============================================= &#39;

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
            ReDim Preserve pKeys(0 To UBound(pKeys) &#43; 1)
            ReDim Preserve pItems(0 To UBound(pItems) &#43; 1)
        End If
        
        Dim FormattedKey As String
        FormattedKey = GetFormattedKey(Key)
        
        If Index &gt; 0 And Index &lt;= pKeyValues.Count Then
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
                Index = i &#43; 1
                Exit For
            End If
        Next i
        
        &#39; Remove existing value
        RemoveKeyValue KeyValue, Index
        
        &#39; Add new key value back
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
        
        If Index &gt;= 0 And Index &lt;= UBound(pKeys) Then
            For i = Index To UBound(pKeys) - 1
                pKeys(i) = pKeys(i &#43; 1)
                
                If IsObject(pItems(i &#43; 1)) Then
                    Set pItems(i) = pItems(i &#43; 1)
                Else
                    pItems(i) = pItems(i &#43; 1)
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
            &#39; Collection does not have method of setting key comparison
            &#39; So case-sensitive keys aren&#39;t supported by default
            &#39; -&gt; Approach: Append lowercase characters to original key
            &#39;    AbC -&gt; AbC__b, abc -&gt; abc__abc, ABC -&gt; ABC
            &#39;    Won&#39;t work in very strange cases, but should work for now
            &#39;    AbBb -&gt; AbBb__bb matches AbbB -&gt; AbbB__bb
            Dim Lowercase As String
            Lowercase = &#34;&#34;
            
            Dim i As Integer
            Dim Ascii As Integer
            Dim Char As String
            For i = 1 To Len(GetFormattedKey)
                Char = VBA.Mid$(GetFormattedKey, i, 1)
                Ascii = Asc(Char)
                If Ascii &gt;= 97 And Ascii &lt;= 122 Then
                    Lowercase = Lowercase &amp; Char
                End If
            Next i
            
            If Lowercase &lt;&gt; &#34;&#34; Then
                GetFormattedKey = GetFormattedKey &amp; &#34;__&#34; &amp; Lowercase
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
        Set pDictionary = CreateObject(&#34;Scripting.Dictionary&#34;)
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

> 作者:   
> URL: http://richfan.site/%E7%A8%8B%E6%8A%80/vba/%E5%BA%95%E7%A8%BF%E5%B0%8F%E5%B8%AE%E6%89%8B%E4%BB%A3%E7%A0%81/%E7%B1%BB%E6%A8%A1%E5%9D%97/  

