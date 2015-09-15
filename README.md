cJSON - JSON Parser Class for VBA (Visual Basic for Applications)
=================================================================

## Table of Contents:
1. [Initial notes](#user-content-initial-notes)
2. [Dependencies](#user-content-dependencies)
3. [Methods and properties](#user-content-methods-and-properties)
4. [Error handling](#user-content-error-handling)
5. [Code examples](#user-content-code-examples)
6. [Changelog](#user-content-changelog)
7. [License](#user-content-license)


## Initial notes:
In order to use __cJSON__ you must either import the file _cJSON.cls_ into your current MS Access project, or start your project using _cJSON-example.accdb_ as a base.

This class has been developed in my workplace as an in-house solution for an MS Access 2010 project.
Therefore, once the project is fully completed, probably no further development will be done.


## Dependencies:

The only dependency is the __"Microsoft Scripting Runtime"__.
It uses __early binding__, so the reference must be added manually.

## Methods and Properties:

#### Methods:
```vbnet
' Description: 
' 	Encodes input Dictionary into JSON String
' Returns: String 
.Encode(Dict As Dictionary) As String

' Description:
' 	Decodes input JSON String into Dictionary
'   - All numbers are converted to Double
' 	- Objects are converted to Dictionary
' 	- Arrays are converted to Collections
' Returns: Dictionary 
.Decode(JSON As String) As Dictionary
```

#### Properties:
```vbnet
.toDictionary     ' Returns the last input as Dictionary
.toString         ' Returns the last input as JSON String
.toHumanReadable  ' Returns the last input as an indented JSON string, 
                  ' for printing or screen visualization
```


## Error Handling:
During `.Decode()`, the function checks for errors on the JSON string. 
If an error is found, the process is stopped inmediately, and a Dictionary is returned. 

This Dictionary contains only one key named "ERROR", with the position of the error.

No error handling is made during `.Encode()`. 


## Code Examples:

#### Class instantiation

```vbnet
' This will be assumed in the next examples
Dim j As New cJSON, _
	json as String, _
	dict as Dictionary
```

#### Decode JSON String to Dictionary

```vbnet
json = "{""name"":""John"", ""surname"":""Doe""}"
Set dict = j.Decode(json)

Debug.Print j.toHumanReadable
MsgBox "Hello " & " " & dict("name") & " " & dict("surname")
```

#### Encode Dictionary to JSON String

```vbnet
Set dict = New Dictionary
dict.Add "name", "John"
dict.Add "surname", "Doe"

json = j.Encode(dict)

Debug.Print j.toHumanReadable
MsgBox json ' This will show {"name":"John", "surname":"Doe"}
```

#### Decode JSON Error Handling

```vbnet
' Missing comma between elements
json = "{""name"":""John"" ""surname"":""Doe""}
Set dict = j.Decode(json)
MsgBox dict("ERROR")

' Missing final curly brace
json = "{""name"":""John"""
Set dict = j.Decode(json)
MsgBox dict("ERROR")
```


## Changelog:

__v1.0__
- First public release.
- Includes all unit tests used during development.


## License:
This software is provided under the terms and conditions of the MIT License

Copyright (c) 2015 José Román Porras Cebriá

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
