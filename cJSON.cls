VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CJSON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Name: cJSON
' Description: JSON parser class for VBA (Visual Basic for Applications)
' Author: José Román Porras Cebriá
' EMail: j.romanporras@gmail.com
'
' Dependencies: Early binding of "Microsoft Scripting Runtime"
'
' Methods:
'   .Encode()         Encodes input Dictionary into JSON
'   .Decode()         Decodes input JSON into Dictionary
' Properties:
'   .toDictionary     Returns the last input as Dictionary
'   .toString         Returns the last input as JSON String
'   .toHumanReadable  Returns the last input as an indented JSON string, for printing or screen visualization
'
Option Compare Binary
Option Explicit

Private FLAG_ERR As Boolean

Private d As New Dictionary
Private s As String
Private n0, n1 As Integer

Public Property Get toDictionary() As Dictionary
    Set toDictionary = d
End Property

Public Property Get toString() As String
    toString = s
End Property

Public Function Decode(json As String) As Dictionary
    FLAG_ERR = False
    s = json
    d.RemoveAll
    
    Dim t As String
    Dim n As Integer
    Dim jsonObj As Dictionary
    
    n = 0
    t = jsonTokenAhead(json, n)
    
    Select Case t
        Case "{"
            Set jsonObj = jsonParseObj(json, n)
        Case vbNullString
            jsonError "Empty string passed as argument"
        Case Else
            jsonError "Missing ""{"" at pos 1"
    End Select
    
    If FLAG_ERR Then
        'Debug.Print "Decode error. Look at key ""ERROR"""
    Else
        Set d = jsonObj
    End If
    
    Set Decode = d
End Function

Private Sub jsonError(error As String)
    FLAG_ERR = True
    
    If Not d.Exists("ERROR") Then
        d.Add "ERROR", error
    End If
End Sub

Private Function jsonTokenAhead(json As String, n As Integer) As String
    jsonTokenAhead = Mid$(json, n + 1, 1)
End Function

Private Function jsonParseObj(json As String, n As Integer) As Dictionary
    Dim t As String, _
        done As Boolean, _
        doneComplete As Boolean, _
        name As String, _
        value As Variant, _
        jsonObj As Dictionary
        
    Set jsonObj = New Dictionary
    n0 = n + 1
    n1 = n0
    
    ' Empty Object
    If jsonTokenAhead(json, n1) = "}" Then
        Set jsonParseObj = jsonObj
        Exit Function
    End If
    
    ' Test all key:value pairs
    doneComplete = False
    Do While Not doneComplete
        ' Test key
        done = False
        Do While Not done
            t = jsonTokenAhead(json, n1)
            
            Select Case t
                Case " ", vbNewLine
                    n1 = n1 + 1
                Case """"
                    name = jsonParseString(json, n1)
                    done = True
                Case Else
                    jsonError "Error at pos " & n1
                    done = True
                    Exit Function
            End Select
        Loop
        
        ' Test colon
        done = False
        n1 = n1 + 1
        
        Do While done = False
            t = jsonTokenAhead(json, n1)
            
            Select Case t
                Case " "
                    n1 = n1 + 1
                Case ":"
                    done = True
                Case Else
                    jsonError "Missing "":"" at pos " & n1
                    done = True
                    Exit Function
            End Select
        Loop
        
        ' Test value
        done = False
        n1 = n1 + 1
        
        Do While done = False
            t = jsonTokenAhead(json, n1)
            
            Select Case t
                Case " "
                    n1 = n1 + 1
                Case "t", "f", "n"
                    value = jsonParseTFN(json, n1)
                    n1 = n1 - 1
                    done = True
                Case 0 To 9
                    value = jsonParseNumber(json, n1)
                    n1 = n1 - 1
                    done = True
                Case """"
                    value = jsonParseString(json, n1)
                    done = True
                Case "["
                    Set value = jsonParseArray(json, n1)
                    done = True
                Case "{"
                    Set value = jsonParseObj(json, n1)
                    done = True
                Case Else
                    jsonError "Error at pos " & n1
                    done = True
                    Exit Function
            End Select
        Loop
        
        If Not name = vbNullString Then
            jsonObj.Add name, value
        Else
            jsonError "Object key name is Null String ("""")"
            Exit Function
        End If
        
        ' Test next value
        done = False
        n1 = n1 + 1
        
        Do While done = False
            t = jsonTokenAhead(json, n1)
            
            Select Case t
                Case " ", vbNewLine
                    n1 = n1 + 1
                Case ","
                    n1 = n1 + 1
                    done = True
                Case "}"
                    Set jsonParseObj = jsonObj
                    done = True
                    doneComplete = True
                Case Else
                    jsonError "Error at pos " & n1
                    done = True
                    doneComplete = True
            End Select
        Loop
    Loop
End Function

Private Function jsonParseString(json As String, n As Integer) As String
    Dim t As String, _
        done As Boolean, _
        value(254) As String, _
        index As Integer
    
    index = 0
    done = False
    n0 = n + 1
    n1 = n0
    
    Do While done = False
        t = jsonTokenAhead(json, n1)
        
        Select Case t
            Case """"
                done = True
            Case vbNewLine, vbNullString
                jsonError "Error at pos " & n1
                Erase value
                done = True
            Case Else
                If index < 255 Then
                    value(index) = t
                    index = index + 1
                    n1 = n1 + 1
                Else
                    jsonError "String too long at pos " & n1
                    done = True
                End If
        End Select
    Loop
    
    jsonParseString = Join(value, "")
End Function

Private Function jsonParseTFN(json As String, n As Integer) As String
    Dim t As String, _
        done As Boolean, _
        value As String
    
    done = False
    n0 = n + 1
    n1 = n0
    
    Do While done = False
        t = jsonTokenAhead(json, n1)
        
        Select Case t
            Case ",", "}", "]"
                done = True
            Case vbNewLine, vbNullString
                jsonError "Error at pos " & n1
                done = True
            Case Else
                n1 = n1 + 1
        End Select
    Loop
    
    value = LCase(Trim(Mid$(json, n0, n1 + 1 - n0)))
    
    Select Case value
    Case "true", "false", "null"
        jsonParseTFN = value
    Case Else
        'jsonParseTFN = Null
        jsonError "Error at pos " & n0
    End Select

End Function

Private Function jsonParseNumber(json As String, n As Integer) As Double
' TODO: Allow exponential numbers like 123.45e+6
' TODO: Control long strings of numbers
    Dim t As String, _
        done As Boolean, _
        value As String, _
        dotcount As Integer
    
    dotcount = 0
    done = False
    n0 = n ' n+1 jumps 1 extra char ahead
    n1 = n0
    
    Do While done = False
        t = jsonTokenAhead(json, n1)
        
        Select Case t
            Case 0 To 9
                n1 = n1 + 1
            Case "."
                dotcount = dotcount + 1
                n1 = n1 + 1
                If dotcount > 1 Then
                    jsonError "Error at pos " & n1
                    done = True
                End If
            Case ",", "}", "]"
                value = CDbl(Trim(Mid$(json, n0 + 1, n1 - n0)))
                jsonParseNumber = value
                done = True
            Case Else
                jsonError "Error at pos " & n1
                done = True
        End Select
    Loop
End Function

Private Function jsonParseArray(json As String, n As Integer) As Collection
    Dim t As String, _
        done As Boolean, _
        doneComplete As Boolean, _
        index As Integer, _
        value As Variant, _
        jsonObj As Collection
        
    Set jsonObj = New Collection
    index = 1
    n0 = n + 1
    n1 = n0
    
    ' Empty Array
    If jsonTokenAhead(json, n1) = "]" Then
        Set jsonParseArray = jsonObj
        Exit Function
    End If
    
    ' Test all values
    doneComplete = False
    Do While Not doneComplete
        ' Test value
        done = False
        
        Do While done = False
            t = jsonTokenAhead(json, n1)
            
            Select Case t
                Case " "
                    n1 = n1 + 1
                Case "t", "f", "n"
                    value = jsonParseTFN(json, n1)
                    n1 = n1 - 1
                    done = True
                Case 0 To 9
                    value = jsonParseNumber(json, n1)
                    n1 = n1 - 1
                    done = True
                Case """"
                    value = jsonParseString(json, n1)
                    done = True
                Case "["
                    Set value = jsonParseArray(json, n1)
                    done = True
                Case "{"
                    Set value = jsonParseObj(json, n1)
                    done = True
                Case Else
                    jsonError "Error at pos " & n1
                    done = True
                    Exit Function
            End Select
        Loop
        
        jsonObj.Add value ', index (index must be String...)
        
        ' Test next value
        done = False
        n1 = n1 + 1
        
        Do While done = False
            t = jsonTokenAhead(json, n1)
            
            Select Case t
                Case " ", vbNewLine
                    n1 = n1 + 1
                Case ","
                    index = index + 1
                    n1 = n1 + 1
                    done = True
                Case "]"
                    Set jsonParseArray = jsonObj
                    done = True
                    doneComplete = True
                Case Else
                    jsonError "Error at pos " & n1
                    done = True
                    doneComplete = True
            End Select
        Loop
    Loop
End Function

Public Function Encode(dict As Dictionary) As String
    Set d = dict
    s = jsonEncodeDict(dict)
    Encode = s
End Function

Private Function jsonEncodeDict(dict As Dictionary) As String
    Dim i, v As Variant, _
        count, index As Integer, _
        obj, objitems As Variant, _
        keyvalue, key, value As Variant

    count = dict.count
    ' Empty object
    If count = 0 Then
        jsonEncodeDict = "{}"
        Exit Function
    End If
    
    index = 0
    ReDim objitems(count - 1) As String
    
    For Each i In dict.keys
    
        If TypeOf dict.item(i) Is Dictionary Then
            Dim tempd As Dictionary
            Set tempd = dict.item(i)
            value = Array("", jsonEncodeDict(tempd), "")
        ElseIf TypeOf dict.item(i) Is Collection Then
            Dim tempc As Collection
            Set tempc = dict.item(i)
            value = Array("", jsonEncodeCol(tempc), "")
        Else
            v = dict.item(i)
            
            If TypeName(v) = "Boolean" Then
                If v Then
                    value = Array("", CStr("true"), "")
                Else
                    value = Array("", CStr("false"), "")
                End If
            ElseIf IsNumeric(v) Then
                value = Array("", v, "")
            ElseIf IsNull(v) Then
                value = Array("", CStr("null"), "")
            Else
                Dim loweri As String
                loweri = LCase(v)
                
                Select Case loweri
                    Case "true", "false", "null"
                        value = Array("", loweri, "")
                    Case Else
                        value = Array("""", v, """")
                End Select
            End If
        End If
        
        key = Array("""", i, """")
        keyvalue = Array(Join(key, ""), Join(value, ""))
        objitems(index) = Join(keyvalue, ":")
        
        index = index + 1
    Next i
    
    obj = Array("{", Join(objitems, ","), "}")
    jsonEncodeDict = Join(obj, "")
End Function

Private Function jsonEncodeCol(coll As Collection) As String
    Dim i, v As Variant, _
        count, index As Integer, _
        obj, objitems As Variant, _
        value As Variant

    count = coll.count
    ' Empty array
    If count = 0 Then
        jsonEncodeCol = "[]"
        Exit Function
    End If
    
    index = 0
    ReDim objitems(count - 1) As String
    
    For Each i In coll
        
        If TypeOf i Is Dictionary Then
            Dim tempd As Dictionary
            Set tempd = i
            value = Array("", jsonEncodeDict(tempd), "")
        ElseIf TypeOf i Is Collection Then
            Dim tempc As Collection
            Set tempc = i
            value = Array("", jsonEncodeCol(tempc), "")
        Else
            v = i
        
            If TypeName(v) = "Boolean" Then
                If v Then
                    value = Array("", CStr("true"), "")
                Else
                    value = Array("", CStr("false"), "")
                End If
            ElseIf IsNumeric(v) Then
                value = Array("", v, "")
            ElseIf IsNull(v) Then
                value = Array("", CStr("null"), "")
            Else
                Dim loweri As String
                loweri = LCase(v)
                
                Select Case loweri
                    Case "true", "false", "null"
                        value = Array("", loweri, "")
                    Case Else
                        value = Array("""", v, """")
                End Select
            End If
        End If
        
        objitems(index) = Join(value, "")
        
        index = index + 1
    Next i
    
    obj = Array("[", Join(objitems, ","), "]")
    jsonEncodeCol = Join(obj, "")
End Function

Public Property Get toHumanReadable() As String
    If d.count = 0 Then
        toHumanReadable = "Nothing to print. Encode or Decode something first."
    Else
        toHumanReadable = jsonPrintDict(d)
    End If
End Property

Private Function jsonPrintDict(dict As Dictionary, Optional Step As Integer = 0) As String
    Dim i, v As Variant, _
        count, index As Integer, _
        obj, objitems As Variant, _
        keyvalue, key, value As Variant

    count = dict.count
    ' Empty object
    If count = 0 Then
        jsonPrintDict = "{}"
        Exit Function
    End If
    
    index = 0
    ReDim objitems(count - 1) As String
    
    For Each i In dict.keys
    
        If TypeOf dict.item(i) Is Dictionary Then
            Dim tempd As Dictionary
            Set tempd = dict.item(i)
            value = Array("", jsonPrintDict(tempd, Step + 1), "")
        ElseIf TypeOf dict.item(i) Is Collection Then
            Dim tempc As Collection
            Set tempc = dict.item(i)
            value = Array("", jsonPrintCol(tempc, Step + 1), "")
        Else
            v = dict.item(i)
            
            If TypeName(v) = "Boolean" Then
                If v Then
                    value = Array("", CStr("true"), "")
                Else
                    value = Array("", CStr("false"), "")
                End If
            ElseIf IsNumeric(v) Then
                value = Array("", v, "")
            ElseIf IsNull(v) Then
                value = Array("", CStr("null"), "")
            Else
                Dim loweri As String
                loweri = LCase(v)
                
                Select Case loweri
                    Case "true", "false", "null"
                        value = Array("", loweri, "")
                    Case Else
                        value = Array("""", v, """")
                End Select
            End If
        End If
        
        key = Array("""", i, """")
        keyvalue = Array(Join(key, ""), Join(value, ""))
        objitems(index) = Space((Step + 1) * 2) & Join(keyvalue, ": ")
        
        index = index + 1
    Next i
    
    obj = Array("{" & vbNewLine, Join(objitems, "," & vbNewLine), vbNewLine & Space(Step * 2) & "}")
    jsonPrintDict = Join(obj, "")
End Function

Private Function jsonPrintCol(coll As Collection, Optional Step As Integer = 0) As String
    Dim i, v As Variant, _
        count, index As Integer, _
        obj, objitems As Variant, _
        value As Variant

    count = coll.count
    ' Empty array
    If count = 0 Then
        jsonPrintCol = "[]"
        Exit Function
    End If
    
    index = 0
    ReDim objitems(count - 1) As String
    
    For Each i In coll
        
        If TypeOf i Is Dictionary Then
            Dim tempd As Dictionary
            Set tempd = i
            value = Array("", jsonPrintDict(tempd, Step + 1), "")
        ElseIf TypeOf i Is Collection Then
            Dim tempc As Collection
            Set tempc = i
            value = Array("", jsonPrintCol(tempc, Step + 1), "")
        Else
            v = i
        
            If TypeName(v) = "Boolean" Then
                If v Then
                    value = Array("", CStr("true"), "")
                Else
                    value = Array("", CStr("false"), "")
                End If
            ElseIf IsNumeric(v) Then
                value = Array("", v, "")
            ElseIf IsNull(v) Then
                value = Array("", CStr("null"), "")
            Else
                Dim loweri As String
                loweri = LCase(v)
                
                Select Case loweri
                    Case "true", "false", "null"
                        value = Array("", loweri, "")
                    Case Else
                        value = Array("""", v, """")
                End Select
            End If
        End If
        
        objitems(index) = Space((Step + 1) * 2) & Join(value, "")
        
        index = index + 1
    Next i
    
    obj = Array("[" & vbNewLine, Join(objitems, "," & vbNewLine), vbNewLine & Space(Step * 2) & "]")
    jsonPrintCol = Join(obj, "")
End Function

