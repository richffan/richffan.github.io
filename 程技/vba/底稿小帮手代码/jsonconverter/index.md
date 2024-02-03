# JsonConverter


## JsonConverter


&lt;!--more--&gt;

```vb
&#39;&#39;
&#39; VBA-JSON v2.3.1
&#39; (c) Tim Hall - https://github.com/VBA-tools/VBA-JSON
&#39;
&#39; JSON Converter for VBA
&#39;
&#39; Errors:
&#39; 10001 - JSON parse error
&#39;
&#39; @class JsonConverter
&#39; @author tim.hall.engr@gmail.com
&#39; @license MIT (http://www.opensource.org/licenses/mit-license.php)
&#39;&#39; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ &#39;
&#39;
&#39; Based originally on vba-json (with extensive changes)
&#39; BSD license included below
&#39;
&#39; JSONLib, http://code.google.com/p/vba-json/
&#39;
&#39; Copyright (c) 2013, Ryo Yokoyama
&#39; All rights reserved.
&#39;
&#39; Redistribution and use in source and binary forms, with or without
&#39; modification, are permitted provided that the following conditions are met:
&#39;     * Redistributions of source code must retain the above copyright
&#39;       notice, this list of conditions and the following disclaimer.
&#39;     * Redistributions in binary form must reproduce the above copyright
&#39;       notice, this list of conditions and the following disclaimer in the
&#39;       documentation and/or other materials provided with the distribution.
&#39;     * Neither the name of the &lt;organization&gt; nor the
&#39;       names of its contributors may be used to endorse or promote products
&#39;       derived from this software without specific prior written permission.
&#39;
&#39; THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS &#34;AS IS&#34; AND
&#39; ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
&#39; WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
&#39; DISCLAIMED. IN NO EVENT SHALL &lt;COPYRIGHT HOLDER&gt; BE LIABLE FOR ANY
&#39; DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
&#39; (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
&#39; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
&#39; ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
&#39; (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
&#39; SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
&#39; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ &#39;
Option Explicit

&#39; === VBA-UTC Headers
#If Mac Then
    
    #If VBA7 Then
        
        &#39; 64-bit Mac (2016)
        Private Declare PtrSafe Function utc_popen Lib &#34;/usr/lib/libc.dylib&#34; Alias &#34;popen&#34; _
            (ByVal utc_Command As String, ByVal utc_Mode As String) As LongPtr
        Private Declare PtrSafe Function utc_pclose Lib &#34;/usr/lib/libc.dylib&#34; Alias &#34;pclose&#34; _
            (ByVal utc_File As LongPtr) As LongPtr
        Private Declare PtrSafe Function utc_fread Lib &#34;/usr/lib/libc.dylib&#34; Alias &#34;fread&#34; _
            (ByVal utc_Buffer As String, ByVal utc_Size As LongPtr, ByVal utc_Number As LongPtr, ByVal utc_File As LongPtr) As LongPtr
        Private Declare PtrSafe Function utc_feof Lib &#34;/usr/lib/libc.dylib&#34; Alias &#34;feof&#34; _
            (ByVal utc_File As LongPtr) As LongPtr
        
    #Else
        
        &#39; 32-bit Mac
        Private Declare Function utc_popen Lib &#34;libc.dylib&#34; Alias &#34;popen&#34; _
            (ByVal utc_Command As String, ByVal utc_Mode As String) As Long
        Private Declare Function utc_pclose Lib &#34;libc.dylib&#34; Alias &#34;pclose&#34; _
            (ByVal utc_File As Long) As Long
        Private Declare Function utc_fread Lib &#34;libc.dylib&#34; Alias &#34;fread&#34; _
            (ByVal utc_Buffer As String, ByVal utc_Size As Long, ByVal utc_Number As Long, ByVal utc_File As Long) As Long
        Private Declare Function utc_feof Lib &#34;libc.dylib&#34; Alias &#34;feof&#34; _
            (ByVal utc_File As Long) As Long
        
    #End If
    
    #ElseIf VBA7 Then
        
        &#39; http://msdn.microsoft.com/en-us/library/windows/desktop/ms724421.aspx
        &#39; http://msdn.microsoft.com/en-us/library/windows/desktop/ms724949.aspx
        &#39; http://msdn.microsoft.com/en-us/library/windows/desktop/ms725485.aspx
        Private Declare PtrSafe Function utc_GetTimeZoneInformation Lib &#34;kernel32&#34; Alias &#34;GetTimeZoneInformation&#34; _
            (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION) As Long
        Private Declare PtrSafe Function utc_SystemTimeToTzSpecificLocalTime Lib &#34;kernel32&#34; Alias &#34;SystemTimeToTzSpecificLocalTime&#34; _
            (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpUniversalTime As utc_SYSTEMTIME, utc_lpLocalTime As utc_SYSTEMTIME) As Long
        Private Declare PtrSafe Function utc_TzSpecificLocalTimeToSystemTime Lib &#34;kernel32&#34; Alias &#34;TzSpecificLocalTimeToSystemTime&#34; _
            (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpLocalTime As utc_SYSTEMTIME, utc_lpUniversalTime As utc_SYSTEMTIME) As Long
        
    #Else
        
        Private Declare Function utc_GetTimeZoneInformation Lib &#34;kernel32&#34; Alias &#34;GetTimeZoneInformation&#34; _
            (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION) As Long
        Private Declare Function utc_SystemTimeToTzSpecificLocalTime Lib &#34;kernel32&#34; Alias &#34;SystemTimeToTzSpecificLocalTime&#34; _
            (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpUniversalTime As utc_SYSTEMTIME, utc_lpLocalTime As utc_SYSTEMTIME) As Long
        Private Declare Function utc_TzSpecificLocalTimeToSystemTime Lib &#34;kernel32&#34; Alias &#34;TzSpecificLocalTimeToSystemTime&#34; _
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
    &#39; === End VBA-UTC
    
    Private Type json_Options
        &#39; VBA only stores 15 significant digits, so any numbers larger than that are truncated
        &#39; This can lead to issues when BIGINT&#39;s are used (e.g. for Ids or Credit Cards), as they will be invalid above 15 digits
        &#39; See: http://support.microsoft.com/kb/269370
        &#39;
        &#39; By default, VBA-JSON will use String for numbers longer than 15 characters that contain only digits
        &#39; to override set `JsonConverter.JsonOptions.UseDoubleForLargeNumbers = True`
        UseDoubleForLargeNumbers As Boolean
        
        &#39; The JSON standard requires object keys to be quoted (&#34; or &#39;), use this option to allow unquoted keys
        AllowUnquotedKeys As Boolean
        
        &#39; The solidus (/) is not required to be escaped, use this option to escape them as \/ in ConvertToJson
        EscapeSolidus As Boolean
    End Type
    Public JsonOptions As json_Options
    
    &#39; ============================================= &#39;
    &#39; Public Methods
    &#39; ============================================= &#39;
    
    &#39;&#39;
    &#39; Convert JSON string to object (Dictionary/Collection)
    &#39;
    &#39; @method ParseJson
    &#39; @param {String} json_String
    &#39; @return {Object} (Dictionary or Collection)
    &#39; @throws 10001 - JSON parse error
    &#39;&#39;
    Public Function ParseJson(ByVal JsonString As String) As Object
        Dim json_Index As Long
        json_Index = 1
        
        &#39; Remove vbCr, vbLf, and vbTab from json_String
        JsonString = VBA.Replace(VBA.Replace(VBA.Replace(JsonString, VBA.vbCr, &#34;&#34;), VBA.vbLf, &#34;&#34;), VBA.vbTab, &#34;&#34;)
        
        json_SkipSpaces JsonString, json_Index
        Select Case VBA.Mid$(JsonString, json_Index, 1)
            Case &#34;{&#34;
                Set ParseJson = json_ParseObject(JsonString, json_Index)
            Case &#34;[&#34;
                Set ParseJson = json_ParseArray(JsonString, json_Index)
            Case Else
                &#39; Error: Invalid JSON string
                ERR.Raise 10001, &#34;JSONConverter&#34;, json_ParseErrorMessage(JsonString, json_Index, &#34;Expecting &#39;{&#39; or &#39;[&#39;&#34;)
        End Select
    End Function
    
    &#39;&#39;
    &#39; Convert object (Dictionary/Collection/Array) to JSON
    &#39;
    &#39; @method ConvertToJson
    &#39; @param {Variant} JsonValue (Dictionary, Collection, or Array)
    &#39; @param {Integer|String} Whitespace &#34;Pretty&#34; print json with given number of spaces per indentation (Integer) or given string
    &#39; @return {String}
    &#39;&#39;
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
                ConvertToJson = &#34;null&#34;
            Case VBA.vbDate
                &#39; Date
                json_DateStr = ConvertToIso(VBA.CDate(JsonValue))
                
                ConvertToJson = &#34;&#34;&#34;&#34; &amp; json_DateStr &amp; &#34;&#34;&#34;&#34;
            Case VBA.vbString
                &#39; String (or large number encoded as string)
                If Not JsonOptions.UseDoubleForLargeNumbers And json_StringIsLargeNumber(JsonValue) Then
                    ConvertToJson = JsonValue
                Else
                    ConvertToJson = &#34;&#34;&#34;&#34; &amp; json_Encode(JsonValue) &amp; &#34;&#34;&#34;&#34;
                End If
            Case VBA.vbBoolean
                If JsonValue Then
                    ConvertToJson = &#34;true&#34;
                Else
                    ConvertToJson = &#34;false&#34;
                End If
            Case VBA.vbArray To VBA.vbArray &#43; VBA.vbByte
                If json_PrettyPrint Then
                    If VBA.VarType(Whitespace) = VBA.vbString Then
                        json_Indentation = VBA.String$(json_CurrentIndentation &#43; 1, Whitespace)
                        json_InnerIndentation = VBA.String$(json_CurrentIndentation &#43; 2, Whitespace)
                    Else
                        json_Indentation = VBA.Space$((json_CurrentIndentation &#43; 1) * Whitespace)
                        json_InnerIndentation = VBA.Space$((json_CurrentIndentation &#43; 2) * Whitespace)
                    End If
                End If
                
                &#39; Array
                json_BufferAppend json_Buffer, &#34;[&#34;, json_BufferPosition, json_BufferLength
                
                On Error Resume Next
                
                json_LBound = LBound(JsonValue, 1)
                json_UBound = UBound(JsonValue, 1)
                json_LBound2D = LBound(JsonValue, 2)
                json_UBound2D = UBound(JsonValue, 2)
                
                If json_LBound &gt;= 0 And json_UBound &gt;= 0 Then
                    For json_Index = json_LBound To json_UBound
                        If json_IsFirstItem Then
                            json_IsFirstItem = False
                        Else
                            &#39; Append comma to previous line
                            json_BufferAppend json_Buffer, &#34;,&#34;, json_BufferPosition, json_BufferLength
                        End If
                        
                        If json_LBound2D &gt;= 0 And json_UBound2D &gt;= 0 Then
                            &#39; 2D Array
                            If json_PrettyPrint Then
                                json_BufferAppend json_Buffer, vbNewLine, json_BufferPosition, json_BufferLength
                            End If
                            json_BufferAppend json_Buffer, json_Indentation &amp; &#34;[&#34;, json_BufferPosition, json_BufferLength
                            
                            For json_Index2D = json_LBound2D To json_UBound2D
                                If json_IsFirstItem2D Then
                                    json_IsFirstItem2D = False
                                Else
                                    json_BufferAppend json_Buffer, &#34;,&#34;, json_BufferPosition, json_BufferLength
                                End If
                                
                                json_Converted = ConvertToJson(JsonValue(json_Index, json_Index2D), Whitespace, json_CurrentIndentation &#43; 2)
                                
                                &#39; For Arrays/Collections, undefined (Empty/Nothing) is treated as null
                                If json_Converted = &#34;&#34; Then
                                    &#39; (nest to only check if converted = &#34;&#34;)
                                    If json_IsUndefined(JsonValue(json_Index, json_Index2D)) Then
                                        json_Converted = &#34;null&#34;
                                    End If
                                End If
                                
                                If json_PrettyPrint Then
                                    json_Converted = vbNewLine &amp; json_InnerIndentation &amp; json_Converted
                                End If
                                
                                json_BufferAppend json_Buffer, json_Converted, json_BufferPosition, json_BufferLength
                            Next json_Index2D
                            
                            If json_PrettyPrint Then
                                json_BufferAppend json_Buffer, vbNewLine, json_BufferPosition, json_BufferLength
                            End If
                            
                            json_BufferAppend json_Buffer, json_Indentation &amp; &#34;]&#34;, json_BufferPosition, json_BufferLength
                            json_IsFirstItem2D = True
                        Else
                            &#39; 1D Array
                            json_Converted = ConvertToJson(JsonValue(json_Index), Whitespace, json_CurrentIndentation &#43; 1)
                            
                            &#39; For Arrays/Collections, undefined (Empty/Nothing) is treated as null
                            If json_Converted = &#34;&#34; Then
                                &#39; (nest to only check if converted = &#34;&#34;)
                                If json_IsUndefined(JsonValue(json_Index)) Then
                                    json_Converted = &#34;null&#34;
                                End If
                            End If
                            
                            If json_PrettyPrint Then
                                json_Converted = vbNewLine &amp; json_Indentation &amp; json_Converted
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
                
                json_BufferAppend json_Buffer, json_Indentation &amp; &#34;]&#34;, json_BufferPosition, json_BufferLength
                
                ConvertToJson = json_BufferToString(json_Buffer, json_BufferPosition)
                
                &#39; Dictionary or Collection
            Case VBA.vbObject
                If json_PrettyPrint Then
                    If VBA.VarType(Whitespace) = VBA.vbString Then
                        json_Indentation = VBA.String$(json_CurrentIndentation &#43; 1, Whitespace)
                    Else
                        json_Indentation = VBA.Space$((json_CurrentIndentation &#43; 1) * Whitespace)
                    End If
                End If
                
                &#39; Dictionary
                If VBA.TypeName(JsonValue) = &#34;Dictionary&#34; Then
                    json_BufferAppend json_Buffer, &#34;{&#34;, json_BufferPosition, json_BufferLength
                    For Each json_Key In JsonValue.Keys
                        &#39; For Objects, undefined (Empty/Nothing) is not added to object
                        json_Converted = ConvertToJson(JsonValue(json_Key), Whitespace, json_CurrentIndentation &#43; 1)
                        If json_Converted = &#34;&#34; Then
                            json_SkipItem = json_IsUndefined(JsonValue(json_Key))
                        Else
                            json_SkipItem = False
                        End If
                        
                        If Not json_SkipItem Then
                            If json_IsFirstItem Then
                                json_IsFirstItem = False
                            Else
                                json_BufferAppend json_Buffer, &#34;,&#34;, json_BufferPosition, json_BufferLength
                            End If
                            
                            If json_PrettyPrint Then
                                json_Converted = vbNewLine &amp; json_Indentation &amp; &#34;&#34;&#34;&#34; &amp; json_Key &amp; &#34;&#34;&#34;: &#34; &amp; json_Converted
                            Else
                                json_Converted = &#34;&#34;&#34;&#34; &amp; json_Key &amp; &#34;&#34;&#34;:&#34; &amp; json_Converted
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
                    
                    json_BufferAppend json_Buffer, json_Indentation &amp; &#34;}&#34;, json_BufferPosition, json_BufferLength
                    
                    &#39; Collection
                ElseIf VBA.TypeName(JsonValue) = &#34;Collection&#34; Then
                    json_BufferAppend json_Buffer, &#34;[&#34;, json_BufferPosition, json_BufferLength
                    For Each json_Value In JsonValue
                        If json_IsFirstItem Then
                            json_IsFirstItem = False
                        Else
                            json_BufferAppend json_Buffer, &#34;,&#34;, json_BufferPosition, json_BufferLength
                        End If
                        
                        json_Converted = ConvertToJson(json_Value, Whitespace, json_CurrentIndentation &#43; 1)
                        
                        &#39; For Arrays/Collections, undefined (Empty/Nothing) is treated as null
                        If json_Converted = &#34;&#34; Then
                            &#39; (nest to only check if converted = &#34;&#34;)
                            If json_IsUndefined(json_Value) Then
                                json_Converted = &#34;null&#34;
                            End If
                        End If
                        
                        If json_PrettyPrint Then
                            json_Converted = vbNewLine &amp; json_Indentation &amp; json_Converted
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
                    
                    json_BufferAppend json_Buffer, json_Indentation &amp; &#34;]&#34;, json_BufferPosition, json_BufferLength
                End If
                
                ConvertToJson = json_BufferToString(json_Buffer, json_BufferPosition)
            Case VBA.vbInteger, VBA.vbLong, VBA.vbSingle, VBA.vbDouble, VBA.vbCurrency, VBA.vbDecimal
                &#39; Number (use decimals for numbers)
                ConvertToJson = VBA.Replace(JsonValue, &#34;,&#34;, &#34;.&#34;)
            Case Else
                &#39; vbEmpty, vbError, vbDataObject, vbByte, vbUserDefinedType
                &#39; Use VBA&#39;s built-in to-string
                On Error Resume Next
                ConvertToJson = JsonValue
                On Error GoTo 0
        End Select
    End Function
    
    &#39; ============================================= &#39;
    &#39; Private Functions
    &#39; ============================================= &#39;
    
    Private Function json_ParseObject(json_String As String, ByRef json_Index As Long) As Dictionary
        Dim json_Key As String
        Dim json_NextChar As String
        
        Set json_ParseObject = New Dictionary
        json_SkipSpaces json_String, json_Index
        If VBA.Mid$(json_String, json_Index, 1) &lt;&gt; &#34;{&#34; Then
            ERR.Raise 10001, &#34;JSONConverter&#34;, json_ParseErrorMessage(json_String, json_Index, &#34;Expecting &#39;{&#39;&#34;)
        Else
            json_Index = json_Index &#43; 1
            
            Do
                json_SkipSpaces json_String, json_Index
                If VBA.Mid$(json_String, json_Index, 1) = &#34;}&#34; Then
                    json_Index = json_Index &#43; 1
                    Exit Function
                ElseIf VBA.Mid$(json_String, json_Index, 1) = &#34;,&#34; Then
                    json_Index = json_Index &#43; 1
                    json_SkipSpaces json_String, json_Index
                End If
                
                json_Key = json_ParseKey(json_String, json_Index)
                json_NextChar = json_Peek(json_String, json_Index)
                If json_NextChar = &#34;[&#34; Or json_NextChar = &#34;{&#34; Then
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
        If VBA.Mid$(json_String, json_Index, 1) &lt;&gt; &#34;[&#34; Then
            ERR.Raise 10001, &#34;JSONConverter&#34;, json_ParseErrorMessage(json_String, json_Index, &#34;Expecting &#39;[&#39;&#34;)
        Else
            json_Index = json_Index &#43; 1
            
            Do
                json_SkipSpaces json_String, json_Index
                If VBA.Mid$(json_String, json_Index, 1) = &#34;]&#34; Then
                    json_Index = json_Index &#43; 1
                    Exit Function
                ElseIf VBA.Mid$(json_String, json_Index, 1) = &#34;,&#34; Then
                    json_Index = json_Index &#43; 1
                    json_SkipSpaces json_String, json_Index
                End If
                
                json_ParseArray.Add json_ParseValue(json_String, json_Index)
            Loop
        End If
    End Function
    
    Private Function json_ParseValue(json_String As String, ByRef json_Index As Long) As Variant
        json_SkipSpaces json_String, json_Index
        Select Case VBA.Mid$(json_String, json_Index, 1)
            Case &#34;{&#34;
                Set json_ParseValue = json_ParseObject(json_String, json_Index)
            Case &#34;[&#34;
                Set json_ParseValue = json_ParseArray(json_String, json_Index)
            Case &#34;&#34;&#34;&#34;, &#34;&#39;&#34;
                json_ParseValue = json_ParseString(json_String, json_Index)
            Case Else
                If VBA.Mid$(json_String, json_Index, 4) = &#34;true&#34; Then
                    json_ParseValue = True
                    json_Index = json_Index &#43; 4
                ElseIf VBA.Mid$(json_String, json_Index, 5) = &#34;false&#34; Then
                    json_ParseValue = False
                    json_Index = json_Index &#43; 5
                ElseIf VBA.Mid$(json_String, json_Index, 4) = &#34;null&#34; Then
                    json_ParseValue = Null
                    json_Index = json_Index &#43; 4
                ElseIf VBA.InStr(&#34;&#43;-0123456789&#34;, VBA.Mid$(json_String, json_Index, 1)) Then
                    json_ParseValue = json_ParseNumber(json_String, json_Index)
                Else
                    ERR.Raise 10001, &#34;JSONConverter&#34;, json_ParseErrorMessage(json_String, json_Index, &#34;Expecting &#39;STRING&#39;, &#39;NUMBER&#39;, null, true, false, &#39;{&#39;, or &#39;[&#39;&#34;)
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
        
        &#39; Store opening quote to look for matching closing quote
        json_Quote = VBA.Mid$(json_String, json_Index, 1)
        json_Index = json_Index &#43; 1
        
        Do While json_Index &gt; 0 And json_Index &lt;= Len(json_String)
            json_Char = VBA.Mid$(json_String, json_Index, 1)
            
            Select Case json_Char
                Case &#34;\&#34;
                    &#39; Escaped string, \\, or \/
                    json_Index = json_Index &#43; 1
                    json_Char = VBA.Mid$(json_String, json_Index, 1)
                    
                    Select Case json_Char
                        Case &#34;&#34;&#34;&#34;, &#34;\&#34;, &#34;/&#34;, &#34;&#39;&#34;
                            json_BufferAppend json_Buffer, json_Char, json_BufferPosition, json_BufferLength
                            json_Index = json_Index &#43; 1
                        Case &#34;b&#34;
                            json_BufferAppend json_Buffer, vbBack, json_BufferPosition, json_BufferLength
                            json_Index = json_Index &#43; 1
                        Case &#34;f&#34;
                            json_BufferAppend json_Buffer, vbFormFeed, json_BufferPosition, json_BufferLength
                            json_Index = json_Index &#43; 1
                        Case &#34;n&#34;
                            json_BufferAppend json_Buffer, vbCrLf, json_BufferPosition, json_BufferLength
                            json_Index = json_Index &#43; 1
                        Case &#34;r&#34;
                            json_BufferAppend json_Buffer, vbCr, json_BufferPosition, json_BufferLength
                            json_Index = json_Index &#43; 1
                        Case &#34;t&#34;
                            json_BufferAppend json_Buffer, vbTab, json_BufferPosition, json_BufferLength
                            json_Index = json_Index &#43; 1
                        Case &#34;u&#34;
                            &#39; Unicode character escape (e.g. \u00a9 = Copyright)
                            json_Index = json_Index &#43; 1
                            json_Code = VBA.Mid$(json_String, json_Index, 4)
                            json_BufferAppend json_Buffer, VBA.ChrW(VBA.Val(&#34;&amp;h&#34; &#43; json_Code)), json_BufferPosition, json_BufferLength
                            json_Index = json_Index &#43; 4
                    End Select
                Case json_Quote
                    json_ParseString = json_BufferToString(json_Buffer, json_BufferPosition)
                    json_Index = json_Index &#43; 1
                    Exit Function
                Case Else
                    json_BufferAppend json_Buffer, json_Char, json_BufferPosition, json_BufferLength
                    json_Index = json_Index &#43; 1
            End Select
        Loop
    End Function
    
    Private Function json_ParseNumber(json_String As String, ByRef json_Index As Long) As Variant
        Dim json_Char As String
        Dim json_Value As String
        Dim json_IsLargeNumber As Boolean
        
        json_SkipSpaces json_String, json_Index
        
        Do While json_Index &gt; 0 And json_Index &lt;= Len(json_String)
            json_Char = VBA.Mid$(json_String, json_Index, 1)
            
            If VBA.InStr(&#34;&#43;-0123456789.eE&#34;, json_Char) Then
                &#39; Unlikely to have massive number, so use simple append rather than buffer here
                json_Value = json_Value &amp; json_Char
                json_Index = json_Index &#43; 1
            Else
                &#39; Excel only stores 15 significant digits, so any numbers larger than that are truncated
                &#39; This can lead to issues when BIGINT&#39;s are used (e.g. for Ids or Credit Cards), as they will be invalid above 15 digits
                &#39; See: http://support.microsoft.com/kb/269370
                &#39;
                &#39; Fix: Parse -&gt; String, Convert -&gt; String longer than 15/16 characters containing only numbers and decimal points -&gt; Number
                &#39; (decimal doesn&#39;t factor into significant digit count, so if present check for 15 digits &#43; decimal = 16)
                json_IsLargeNumber = IIf(InStr(json_Value, &#34;.&#34;), Len(json_Value) &gt;= 17, Len(json_Value) &gt;= 16)
                If Not JsonOptions.UseDoubleForLargeNumbers And json_IsLargeNumber Then
                    json_ParseNumber = json_Value
                Else
                    &#39; VBA.Val does not use regional settings, so guard for comma is not needed
                    json_ParseNumber = VBA.Val(json_Value)
                End If
                Exit Function
            End If
        Loop
    End Function
    
    Private Function json_ParseKey(json_String As String, ByRef json_Index As Long) As String
        &#39; Parse key with single or double quotes
        If VBA.Mid$(json_String, json_Index, 1) = &#34;&#34;&#34;&#34; Or VBA.Mid$(json_String, json_Index, 1) = &#34;&#39;&#34; Then
            json_ParseKey = json_ParseString(json_String, json_Index)
        ElseIf JsonOptions.AllowUnquotedKeys Then
            Dim json_Char As String
            Do While json_Index &gt; 0 And json_Index &lt;= Len(json_String)
                json_Char = VBA.Mid$(json_String, json_Index, 1)
                If (json_Char &lt;&gt; &#34; &#34;) And (json_Char &lt;&gt; &#34;:&#34;) Then
                    json_ParseKey = json_ParseKey &amp; json_Char
                    json_Index = json_Index &#43; 1
                Else
                    Exit Do
                End If
            Loop
        Else
            ERR.Raise 10001, &#34;JSONConverter&#34;, json_ParseErrorMessage(json_String, json_Index, &#34;Expecting &#39;&#34;&#34;&#39; or &#39;&#39;&#39;&#34;)
        End If
        
        &#39; Check for colon and skip if present or throw if not present
        json_SkipSpaces json_String, json_Index
        If VBA.Mid$(json_String, json_Index, 1) &lt;&gt; &#34;:&#34; Then
            ERR.Raise 10001, &#34;JSONConverter&#34;, json_ParseErrorMessage(json_String, json_Index, &#34;Expecting &#39;:&#39;&#34;)
        Else
            json_Index = json_Index &#43; 1
        End If
    End Function
    
    Private Function json_IsUndefined(ByVal json_Value As Variant) As Boolean
        &#39; Empty / Nothing -&gt; undefined
        Select Case VBA.VarType(json_Value)
            Case VBA.vbEmpty
                json_IsUndefined = True
            Case VBA.vbObject
                Select Case VBA.TypeName(json_Value)
                    Case &#34;Empty&#34;, &#34;Nothing&#34;
                        json_IsUndefined = True
                End Select
        End Select
    End Function
    
    Private Function json_Encode(ByVal json_Text As Variant) As String
        &#39; Reference: http://www.ietf.org/rfc/rfc4627.txt
        &#39; Escape: &#34;, \, /, backspace, form feed, line feed, carriage return, tab
        Dim json_Index As Long
        Dim json_Char As String
        Dim json_AscCode As Long
        Dim json_Buffer As String
        Dim json_BufferPosition As Long
        Dim json_BufferLength As Long
        
        For json_Index = 1 To VBA.Len(json_Text)
            json_Char = VBA.Mid$(json_Text, json_Index, 1)
            json_AscCode = VBA.AscW(json_Char)
            
            &#39; When AscW returns a negative number, it returns the twos complement form of that number.
            &#39; To convert the twos complement notation into normal binary notation, add 0xFFF to the return result.
            &#39; https://support.microsoft.com/en-us/kb/272138
            If json_AscCode &lt; 0 Then
                json_AscCode = json_AscCode &#43; 65536
            End If
            
            &#39; From spec, &#34;, \, and control characters must be escaped (solidus is optional)
            
            Select Case json_AscCode
                Case 34
                    &#39; &#34; -&gt; 34 -&gt; \&#34;
                    json_Char = &#34;\&#34;&#34;&#34;
                Case 92
                    &#39; \ -&gt; 92 -&gt; \\
                    json_Char = &#34;\\&#34;
                Case 47
                    &#39; / -&gt; 47 -&gt; \/ (optional)
                    If JsonOptions.EscapeSolidus Then
                        json_Char = &#34;\/&#34;
                    End If
                Case 8
                    &#39; backspace -&gt; 8 -&gt; \b
                    json_Char = &#34;\b&#34;
                Case 12
                    &#39; form feed -&gt; 12 -&gt; \f
                    json_Char = &#34;\f&#34;
                Case 10
                    &#39; line feed -&gt; 10 -&gt; \n
                    json_Char = &#34;\n&#34;
                Case 13
                    &#39; carriage return -&gt; 13 -&gt; \r
                    json_Char = &#34;\r&#34;
                Case 9
                    &#39; tab -&gt; 9 -&gt; \t
                    json_Char = &#34;\t&#34;
                Case 0 To 31, 127 To 65535
                    &#39; Non-ascii characters -&gt; convert to 4-digit hex
                    json_Char = &#34;\u&#34; &amp; VBA.Right$(&#34;0000&#34; &amp; VBA.Hex$(json_AscCode), 4)
            End Select
            
            json_BufferAppend json_Buffer, json_Char, json_BufferPosition, json_BufferLength
        Next json_Index
        
        json_Encode = json_BufferToString(json_Buffer, json_BufferPosition)
    End Function
    
    Private Function json_Peek(json_String As String, ByVal json_Index As Long, Optional json_NumberOfCharacters As Long = 1) As String
        &#39; &#34;Peek&#34; at the next number of characters without incrementing json_Index (ByVal instead of ByRef)
        json_SkipSpaces json_String, json_Index
        json_Peek = VBA.Mid$(json_String, json_Index, json_NumberOfCharacters)
    End Function
    
    Private Sub json_SkipSpaces(json_String As String, ByRef json_Index As Long)
        &#39; Increment index to skip over spaces
        Do While json_Index &gt; 0 And json_Index &lt;= VBA.Len(json_String) And VBA.Mid$(json_String, json_Index, 1) = &#34; &#34;
            json_Index = json_Index &#43; 1
        Loop
    End Sub
    
    Private Function json_StringIsLargeNumber(json_String As Variant) As Boolean
        &#39; Check if the given string is considered a &#34;large number&#34;
        &#39; (See json_ParseNumber)
        
        Dim json_Length As Long
        Dim json_CharIndex As Long
        json_Length = VBA.Len(json_String)
        
        &#39; Length with be at least 16 characters and assume will be less than 100 characters
        If json_Length &gt;= 16 And json_Length &lt;= 100 Then
            Dim json_CharCode As String
            
            json_StringIsLargeNumber = True
            
            For json_CharIndex = 1 To json_Length
                json_CharCode = VBA.Asc(VBA.Mid$(json_String, json_CharIndex, 1))
                Select Case json_CharCode
                    &#39; Look for .|0-9|E|e
                    Case 46, 48 To 57, 69, 101
                        &#39; Continue through characters
                    Case Else
                        json_StringIsLargeNumber = False
                        Exit Function
                End Select
            Next json_CharIndex
        End If
    End Function
    
    Private Function json_ParseErrorMessage(json_String As String, ByRef json_Index As Long, ErrorMessage As String)
        &#39; Provide detailed parse error message, including details of where and what occurred
        &#39;
        &#39; Example:
        &#39; Error parsing JSON:
        &#39; {&#34;abcde&#34;:True}
        &#39;          ^
        &#39; Expecting &#39;STRING&#39;, &#39;NUMBER&#39;, null, true, false, &#39;{&#39;, or &#39;[&#39;
        
        Dim json_StartIndex As Long
        Dim json_StopIndex As Long
        
        &#39; Include 10 characters before and after error (if possible)
        json_StartIndex = json_Index - 10
        json_StopIndex = json_Index &#43; 10
        If json_StartIndex &lt;= 0 Then
            json_StartIndex = 1
        End If
        If json_StopIndex &gt; VBA.Len(json_String) Then
            json_StopIndex = VBA.Len(json_String)
        End If
        
        json_ParseErrorMessage = &#34;Error parsing JSON:&#34; &amp; VBA.vbNewLine &amp; _
            VBA.Mid$(json_String, json_StartIndex, json_StopIndex - json_StartIndex &#43; 1) &amp; VBA.vbNewLine &amp; _
            VBA.Space$(json_Index - json_StartIndex) &amp; &#34;^&#34; &amp; VBA.vbNewLine &amp; _
            ErrorMessage
    End Function
    
    Private Sub json_BufferAppend(ByRef json_Buffer As String, _
            ByRef json_Append As Variant, _
            ByRef json_BufferPosition As Long, _
            ByRef json_BufferLength As Long)
        &#39; VBA can be slow to append strings due to allocating a new string for each append
        &#39; Instead of using the traditional append, allocate a large empty string and then copy string at append position
        &#39;
        &#39; Example:
        &#39; Buffer: &#34;abc  &#34;
        &#39; Append: &#34;def&#34;
        &#39; Buffer Position: 3
        &#39; Buffer Length: 5
        &#39;
        &#39; Buffer position &#43; Append length &gt; Buffer length -&gt; Append chunk of blank space to buffer
        &#39; Buffer: &#34;abc       &#34;
        &#39; Buffer Length: 10
        &#39;
        &#39; Put &#34;def&#34; into buffer at position 3 (0-based)
        &#39; Buffer: &#34;abcdef    &#34;
        &#39;
        &#39; Approach based on cStringBuilder from vbAccelerator
        &#39; http://www.vbaccelerator.com/home/VB/Code/Techniques/RunTime_Debug_Tracing/VB6_Tracer_Utility_zip_cStringBuilder_cls.asp
        &#39;
        &#39; and clsStringAppend from Philip Swannell
        &#39; https://github.com/VBA-tools/VBA-JSON/pull/82
        
        Dim json_AppendLength As Long
        Dim json_LengthPlusPosition As Long
        
        json_AppendLength = VBA.Len(json_Append)
        json_LengthPlusPosition = json_AppendLength &#43; json_BufferPosition
        
        If json_LengthPlusPosition &gt; json_BufferLength Then
            &#39; Appending would overflow buffer, add chunk
            &#39; (double buffer length or append length, whichever is bigger)
            Dim json_AddedLength As Long
            json_AddedLength = IIf(json_AppendLength &gt; json_BufferLength, json_AppendLength, json_BufferLength)
            
            json_Buffer = json_Buffer &amp; VBA.Space$(json_AddedLength)
            json_BufferLength = json_BufferLength &#43; json_AddedLength
        End If
        
        &#39; Note: Namespacing with VBA.Mid$ doesn&#39;t work properly here, throwing compile error:
        &#39; Function call on left-hand side of assignment must return Variant or Object
        Mid$(json_Buffer, json_BufferPosition &#43; 1, json_AppendLength) = CStr(json_Append)
        json_BufferPosition = json_BufferPosition &#43; json_AppendLength
    End Sub
    
    Private Function json_BufferToString(ByRef json_Buffer As String, ByVal json_BufferPosition As Long) As String
        If json_BufferPosition &gt; 0 Then
            json_BufferToString = VBA.Left$(json_Buffer, json_BufferPosition)
        End If
    End Function
    
    &#39;&#39;
    &#39; VBA-UTC v1.0.6
    &#39; (c) Tim Hall - https://github.com/VBA-tools/VBA-UtcConverter
    &#39;
    &#39; UTC/ISO 8601 Converter for VBA
    &#39;
    &#39; Errors:
    &#39; 10011 - UTC parsing error
    &#39; 10012 - UTC conversion error
    &#39; 10013 - ISO 8601 parsing error
    &#39; 10014 - ISO 8601 conversion error
    &#39;
    &#39; @module UtcConverter
    &#39; @author tim.hall.engr@gmail.com
    &#39; @license MIT (http://www.opensource.org/licenses/mit-license.php)
    &#39;&#39; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ &#39;
    
    &#39; (Declarations moved to top)
    
    &#39; ============================================= &#39;
    &#39; Public Methods
    &#39; ============================================= &#39;
    
    &#39;&#39;
    &#39; Parse UTC date to local date
    &#39;
    &#39; @method ParseUtc
    &#39; @param {Date} UtcDate
    &#39; @return {Date} Local date
    &#39; @throws 10011 - UTC parsing error
    &#39;&#39;
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
        ERR.Raise 10011, &#34;UtcConverter.ParseUtc&#34;, &#34;UTC parsing error: &#34; &amp; ERR.Number &amp; &#34; - &#34; &amp; ERR.Description
    End Function
    
    &#39;&#39;
    &#39; Convert local date to UTC date
    &#39;
    &#39; @method ConvertToUrc
    &#39; @param {Date} utc_LocalDate
    &#39; @return {Date} UTC date
    &#39; @throws 10012 - UTC conversion error
    &#39;&#39;
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
        ERR.Raise 10012, &#34;UtcConverter.ConvertToUtc&#34;, &#34;UTC conversion error: &#34; &amp; ERR.Number &amp; &#34; - &#34; &amp; ERR.Description
    End Function
    
    &#39;&#39;
    &#39; Parse ISO 8601 date string to local date
    &#39;
    &#39; @method ParseIso
    &#39; @param {Date} utc_IsoString
    &#39; @return {Date} Local date
    &#39; @throws 10013 - ISO 8601 parsing error
    &#39;&#39;
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
        
        utc_Parts = VBA.Split(utc_IsoString, &#34;T&#34;)
        utc_DateParts = VBA.Split(utc_Parts(0), &#34;-&#34;)
        ParseIso = VBA.DateSerial(VBA.CInt(utc_DateParts(0)), VBA.CInt(utc_DateParts(1)), VBA.CInt(utc_DateParts(2)))
        
        If UBound(utc_Parts) &gt; 0 Then
            If VBA.InStr(utc_Parts(1), &#34;Z&#34;) Then
                utc_TimeParts = VBA.Split(VBA.Replace(utc_Parts(1), &#34;Z&#34;, &#34;&#34;), &#34;:&#34;)
            Else
                utc_OffsetIndex = VBA.InStr(1, utc_Parts(1), &#34;&#43;&#34;)
                If utc_OffsetIndex = 0 Then
                    utc_NegativeOffset = True
                    utc_OffsetIndex = VBA.InStr(1, utc_Parts(1), &#34;-&#34;)
                End If
                
                If utc_OffsetIndex &gt; 0 Then
                    utc_HasOffset = True
                    utc_TimeParts = VBA.Split(VBA.Left$(utc_Parts(1), utc_OffsetIndex - 1), &#34;:&#34;)
                    utc_OffsetParts = VBA.Split(VBA.Right$(utc_Parts(1), Len(utc_Parts(1)) - utc_OffsetIndex), &#34;:&#34;)
                    
                    Select Case UBound(utc_OffsetParts)
                        Case 0
                            utc_Offset = TimeSerial(VBA.CInt(utc_OffsetParts(0)), 0, 0)
                        Case 1
                            utc_Offset = TimeSerial(VBA.CInt(utc_OffsetParts(0)), VBA.CInt(utc_OffsetParts(1)), 0)
                        Case 2
                            &#39; VBA.Val does not use regional settings, use for seconds to avoid decimal/comma issues
                            utc_Offset = TimeSerial(VBA.CInt(utc_OffsetParts(0)), VBA.CInt(utc_OffsetParts(1)), Int(VBA.Val(utc_OffsetParts(2))))
                    End Select
                    
                    If utc_NegativeOffset Then: utc_Offset = -utc_Offset
                Else
                    utc_TimeParts = VBA.Split(utc_Parts(1), &#34;:&#34;)
                End If
            End If
            
            Select Case UBound(utc_TimeParts)
                Case 0
                    ParseIso = ParseIso &#43; VBA.TimeSerial(VBA.CInt(utc_TimeParts(0)), 0, 0)
                Case 1
                    ParseIso = ParseIso &#43; VBA.TimeSerial(VBA.CInt(utc_TimeParts(0)), VBA.CInt(utc_TimeParts(1)), 0)
                Case 2
                    &#39; VBA.Val does not use regional settings, use for seconds to avoid decimal/comma issues
                    ParseIso = ParseIso &#43; VBA.TimeSerial(VBA.CInt(utc_TimeParts(0)), VBA.CInt(utc_TimeParts(1)), Int(VBA.Val(utc_TimeParts(2))))
            End Select
            
            ParseIso = ParseUtc(ParseIso)
            
            If utc_HasOffset Then
                ParseIso = ParseIso - utc_Offset
            End If
        End If
        
        Exit Function
        
utc_ErrorHandling:
        ERR.Raise 10013, &#34;UtcConverter.ParseIso&#34;, &#34;ISO 8601 parsing error for &#34; &amp; utc_IsoString &amp; &#34;: &#34; &amp; ERR.Number &amp; &#34; - &#34; &amp; ERR.Description
    End Function
    
    &#39;&#39;
    &#39; Convert local date to ISO 8601 string
    &#39;
    &#39; @method ConvertToIso
    &#39; @param {Date} utc_LocalDate
    &#39; @return {Date} ISO 8601 string
    &#39; @throws 10014 - ISO 8601 conversion error
    &#39;&#39;
    Public Function ConvertToIso(utc_LocalDate As Date) As String
        On Error GoTo utc_ErrorHandling
        
        ConvertToIso = VBA.Format$(ConvertToUtc(utc_LocalDate), &#34;yyyy-mm-ddTHH:mm:ss.000Z&#34;)
        
        Exit Function
        
utc_ErrorHandling:
        ERR.Raise 10014, &#34;UtcConverter.ConvertToIso&#34;, &#34;ISO 8601 conversion error: &#34; &amp; ERR.Number &amp; &#34; - &#34; &amp; ERR.Description
    End Function
    
    &#39; ============================================= &#39;
    &#39; Private Functions
    &#39; ============================================= &#39;
    
    #If Mac Then
        
        Private Function utc_ConvertDate(utc_Value As Date, Optional utc_ConvertToUtc As Boolean = False) As Date
            Dim utc_ShellCommand As String
            Dim utc_Result As utc_ShellResult
            Dim utc_Parts() As String
            Dim utc_DateParts() As String
            Dim utc_TimeParts() As String
            
            If utc_ConvertToUtc Then
                utc_ShellCommand = &#34;date -ur `date -jf &#39;%Y-%m-%d %H:%M:%S&#39; &#34; &amp; _
                    &#34;&#39;&#34; &amp; VBA.Format$(utc_Value, &#34;yyyy-mm-dd HH:mm:ss&#34;) &amp; &#34;&#39; &#34; &amp; _
                    &#34; &#43;&#39;%s&#39;` &#43;&#39;%Y-%m-%d %H:%M:%S&#39;&#34;
            Else
                utc_ShellCommand = &#34;date -jf &#39;%Y-%m-%d %H:%M:%S %z&#39; &#34; &amp; _
                    &#34;&#39;&#34; &amp; VBA.Format$(utc_Value, &#34;yyyy-mm-dd HH:mm:ss&#34;) &amp; &#34; &#43;0000&#39; &#34; &amp; _
                    &#34;&#43;&#39;%Y-%m-%d %H:%M:%S&#39;&#34;
            End If
            
            utc_Result = utc_ExecuteInShell(utc_ShellCommand)
            
            If utc_Result.utc_Output = &#34;&#34; Then
                ERR.Raise 10015, &#34;UtcConverter.utc_ConvertDate&#34;, &#34;&#39;date&#39; command failed&#34;
            Else
                utc_Parts = Split(utc_Result.utc_Output, &#34; &#34;)
                utc_DateParts = Split(utc_Parts(0), &#34;-&#34;)
                utc_TimeParts = Split(utc_Parts(1), &#34;:&#34;)
                
                utc_ConvertDate = DateSerial(utc_DateParts(0), utc_DateParts(1), utc_DateParts(2)) &#43; _
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
            utc_File = utc_popen(utc_ShellCommand, &#34;r&#34;)
            
            If utc_File = 0 Then: Exit Function
            
            Do While utc_feof(utc_File) = 0
                utc_Chunk = VBA.Space$(50)
                utc_Read = CLng(utc_fread(utc_Chunk, 1, Len(utc_Chunk) - 1, utc_File))
                If utc_Read &gt; 0 Then
                    utc_Chunk = VBA.Left$(utc_Chunk, CLng(utc_Read))
                    utc_ExecuteInShell.utc_Output = utc_ExecuteInShell.utc_Output &amp; utc_Chunk
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
            utc_SystemTimeToDate = DateSerial(utc_Value.utc_wYear, utc_Value.utc_wMonth, utc_Value.utc_wDay) &#43; _
                TimeSerial(utc_Value.utc_wHour, utc_Value.utc_wMinute, utc_Value.utc_wSecond)
        End Function
        
    #End If
```



---

> 作者:   
> URL: http://richfan.site/%E7%A8%8B%E6%8A%80/vba/%E5%BA%95%E7%A8%BF%E5%B0%8F%E5%B8%AE%E6%89%8B%E4%BB%A3%E7%A0%81/jsonconverter/  

