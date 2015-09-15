Attribute VB_Name = "unittest_cJSON"
Option Compare Binary
Option Explicit

Sub jsonUnitTest()
    Dim s As String
    Dim o As New Dictionary
    Dim j As New CJSON
    
    ' Empty string
    Set o = j.Decode("")
    Debug.Assert (o.count = 1)
    Debug.Assert (o.Exists("ERROR"))
    Debug.Assert (o.item("ERROR") = "Empty string passed as argument")
    
    ' No object notation
    Set o = j.Decode("a:1")
    Debug.Assert (o.item("ERROR") = "Missing ""{"" at pos 1")
    
    ' Name without proper string enclosure
    Set o = j.Decode("{  a:1")
    Debug.Assert (o.item("ERROR") = "Error at pos 3")
    
    ' String not closed or wrong value
    Set o = j.Decode("{  ""a:1")
    Debug.Assert (o.item("ERROR") = "Error at pos 7")
    Set o = j.Decode("{  ""a:1" & vbNewLine)
    Debug.Assert (o.item("ERROR") = "Error at pos 9")
    Set o = j.Decode("{  ""a:1" & Space(255))
    Debug.Assert (o.item("ERROR") = "String too long at pos 259")
    
    ' Missing colon
    Set o = j.Decode("{  ""nombre""1")
    Debug.Assert (o.item("ERROR") = "Missing "":"" at pos 11")
    
    ' True, false, null - value not closed
    Set o = j.Decode("{  ""nombre"":   tue  ")
    Debug.Assert (o.item("ERROR") = "Error at pos 20")
    Set o = j.Decode("{  ""nombre"":  flase")
    Debug.Assert (o.item("ERROR") = "Error at pos 19")
    Set o = j.Decode("{  ""nombre"":nul  ")
    Debug.Assert (o.item("ERROR") = "Error at pos 17")
    ' True, false, null - wrong value
    Set o = j.Decode("{  ""nombre"":   tue  }")
    Debug.Assert (o.item("ERROR") = "Error at pos 16")
    Set o = j.Decode("{  ""nombre"":  flase}")
    Debug.Assert (o.item("ERROR") = "Error at pos 15")
    Set o = j.Decode("{  ""nombre"":nul  }")
    Debug.Assert (o.item("ERROR") = "Error at pos 13")
    
    ' Numbers
    ' TODO: Number limits
    Set o = j.Decode("{""número"":1.2.3}")
    Debug.Assert (o.item("ERROR") = "Error at pos 14")
    Set o = j.Decode("{""número"":.2}")
    Debug.Assert (o.item("ERROR") = "Error at pos 10")
    Set o = j.Decode("{""número"":1,2}")
    Debug.Assert (o.item("ERROR") = "Error at pos 12")
    
    ' Arrays
    Set o = j.Decode("{""número"":1,""array"":[1,""z""]}")
    Debug.Assert (Not o.Exists("ERROR") And _
                o.Exists("número") And _
                o.Exists("array"))
    Debug.Assert (o.item("número") = 1)
    Debug.Assert (TypeOf o.item("array") Is Collection)
    Debug.Assert (o.item("array").item(1) = 1)
    Debug.Assert (o.item("array").item(2) = "z")
    Set o = j.Decode("{""número"":1,""array"":[1,,""z""]}")
    Debug.Assert (o.item("ERROR") = "Error at pos 23")
    Set o = j.Decode("{""número"":1,""array"":[1,""z""}")
    Debug.Assert (o.item("ERROR") = "Error at pos 26")
    Set o = j.Decode("{""número"":1,""array"":[1,""z""]")
    Debug.Assert (o.item("ERROR") = "Error at pos 27")
    Set o = j.Decode("{""número"":1,""array"":[,""z""]}")
    Debug.Assert (o.item("ERROR") = "Error at pos 21")
    Set o = j.Decode("{""número"":1,""array"":[1""z""]}")
    Debug.Assert (o.item("ERROR") = "Error at pos 22")
    Set o = j.Decode("{""número"":1,""array"":[]}")
    Debug.Assert (o.count = 2 And _
                    o.item("array").count = 0 And _
                    TypeName(o.item("array")) = "Collection")
    
    ' Objects - simple
    Set o = j.Decode("{"""":1}")
    Debug.Assert (o.item("ERROR") = "Object key name is Null String ("""")")
    Set o = j.Decode("{""número"":1,""object"":{}}")
    Debug.Assert (o.count = 2 And _
                    o.item("object").count = 0)
    
    ' Objects - FULL TEST
    Set o = j.Decode( _
        "{""products"":[" & _
            "{""product"":{" & _
                """id"":1,""sku"":""a"",""visible"":true," & _
                """imgs"":[""1.jpg"",""2.jpg""]}" & _
            "}," & _
            "{""product"":{" & _
                """id"":2,""sku"":""b"",""visible"":null," & _
                """imgs"":[""3.jpg"",""4.jpg""]}" & _
            "}" & _
        "]}")
    Debug.Assert o.count = 1
    Debug.Assert o.item("products").count = 2
    Debug.Assert o.item("products").item(2).item("product").item("visible") = "null"
    Debug.Assert o.item("products").item(2).item("product").item("imgs").item(1) = "3.jpg"
    ' Print las test
    Debug.Print j.toHumanReadable
    Debug.Print ""
    
    ' Encode json - FULL TEST
    ' - True, False, Null and strings "True", "False", "Null"
    ' - Numbers and strings
    ' - Collections to Arrays (with mixed var types)
    ' - Dictionaries to Objects
    ' - Empty Arrays and Objects
    ' - Nested Arrays and Objects
    o.RemoveAll
    o.Add 1, New Collection
    o.item(1).Add "a"
    o.item(1).Add 2
    o.Add 2, New Collection
    o.item(2).Add New Collection
    o.item(2).item(1).Add 1
    o.item(2).item(1).Add 2
    o.item(2).item(1).Add 3
    o.item(2).Add New Dictionary
    o.item(2).item(2).Add "a", 1
    o.item(2).item(2).Add "b", 2
    o.item(2).item(2).Add "c", 3
    o.Add 3, New Collection
    o.Add 4, New Dictionary
    o.Add "a", 1
    o.Add "b", True
    o.Add "c", False
    o.Add "d", Null
    o.Add "e", "True"
    o.Add "f", "False"
    o.Add "g", "Null"
    o.Add "h", "z"
    Debug.Assert j.Encode(o) = "{""1"":[""a"",2],""2"":[[1,2,3],{""a"":1,""b"":2,""c"":3}],""3"":[],""4"":{},""a"":1,""b"":true,""c"":false,""d"":null,""e"":true,""f"":false,""g"":null,""h"":""z""}"
    Debug.Print j.toString
    Debug.Print ""
    Debug.Print j.toHumanReadable
    
End Sub

