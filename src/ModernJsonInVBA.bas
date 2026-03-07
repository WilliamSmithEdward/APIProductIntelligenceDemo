Attribute VB_Name = "ModernJsonInVBA"
' ===========================================================================
' ModernJsonInVBA - Modern JSON Parser for VBA
' ===========================================================================
' A lightweight, zero-dependency JSON parser for Excel VBA.
' Converts any JSON string into native VBA objects:
'
'   JSON object  {}  ->  Scripting.Dictionary  (key/value lookup)
'   JSON array   []  ->  VBA Collection        (1-based index)
'   JSON string      ->  String
'   JSON number      ->  Long  or  Double
'   JSON boolean     ->  Boolean
'   JSON null        ->  Null  (VBA Null variant)
'
' QUICK START
' -----------
'   Dim data As Object
'   Set data = ParseJson("{""name"":""Alice"",""age"":30}")
'   Debug.Print data("name")   ' -> Alice
'   Debug.Print data("age")    ' -> 30
'
'   Dim arr As Collection
'   Set arr = ParseJson("[1,2,3]")
'   Debug.Print arr(1)         ' -> 1  (Collections are 1-based)
'
' HELPER FUNCTIONS (safe, never throw)
' ------------------------------------
'   GetString  dict, "key"         -> String  (returns "" if missing)
'   GetNumber  dict, "key"         -> Double  (returns 0 if missing)
'   GetArray   dict, "key"         -> Collection (returns Nothing if missing)
'   GetObject  dict, "key"         -> Dictionary (returns Nothing if missing)
'   GetNestedValue  node, "a.b[0].c"  -> any nested value via dot/bracket path
' ===========================================================================

Option Explicit

' ---------------------------------------------------------------------------
' Private parsing state  (reset on every ParseJson call)
' ---------------------------------------------------------------------------
Private pInput As String
Private pPos   As Long
Private pLen   As Long

' ===========================================================================
' PUBLIC API
' ===========================================================================

' ---------------------------------------------------------------------------
' ParseJson
' Parses a JSON string and returns the root value.
' Return type is Variant: Object (Dictionary/Collection) or primitive.
' ---------------------------------------------------------------------------
Public Function ParseJson(ByVal jsonStr As String) As Variant
    pInput = jsonStr
    pPos   = 1
    pLen   = Len(jsonStr)
    JsonAssignVariant ParseJson, JsonReadValue()
End Function

' ---------------------------------------------------------------------------
' GetNestedValue
' Navigate a nested JSON graph using dot-notation with optional [n] indexing.
' Example:  GetNestedValue(root, "products[0].reviews[1].rating")
' Returns Null if any segment is missing or the path is invalid.
' ---------------------------------------------------------------------------
Public Function GetNestedValue(ByVal node As Variant, ByVal path As String) As Variant
    ' Normalise: "a[0].b" -> "a.[0].b" so we can split on "."
    Dim normalised As String
    normalised = Replace(path, "[", ".[")

    Dim parts()  As String
    parts = Split(normalised, ".")

    Dim current As Variant
    JsonAssignVariant current, node

    Dim i As Long
    For i = 0 To UBound(parts)
        Dim segment As String
        segment = Trim(parts(i))
        If segment = "" Then GoTo NextSegment

        If Left(segment, 1) = "[" Then
            ' Array index:  [n]  (0-based in JSON -> 1-based in Collection)
            Dim idxStr As String
            idxStr = Mid(segment, 2, Len(segment) - 2)
            If Not IsNumeric(idxStr) Then
                GetNestedValue = Null
                Exit Function
            End If
            Dim idx As Long
            idx = CLng(idxStr) + 1  ' VBA Collection is 1-based
            If TypeName(current) <> "Collection" Then
                GetNestedValue = Null
                Exit Function
            End If
            If idx < 1 Or idx > current.Count Then
                GetNestedValue = Null
                Exit Function
            End If
            JsonAssignVariant current, current(idx)
        Else
            ' Dictionary key
            If TypeName(current) <> "Dictionary" Then
                GetNestedValue = Null
                Exit Function
            End If
            If Not current.Exists(segment) Then
                GetNestedValue = Null
                Exit Function
            End If
            JsonAssignVariant current, current(segment)
        End If

NextSegment:
    Next i

    JsonAssignVariant GetNestedValue, current
End Function

' ---------------------------------------------------------------------------
' GetString  -  safe string accessor on a Dictionary
' ---------------------------------------------------------------------------
Public Function GetString(ByVal dict As Object, ByVal key As String, _
                          Optional ByVal defaultVal As String = "") As String
    If dict Is Nothing Then GetString = defaultVal : Exit Function
    If Not dict.Exists(key) Then GetString = defaultVal : Exit Function
    If IsNull(dict(key)) Then GetString = defaultVal : Exit Function
    GetString = CStr(dict(key))
End Function

' ---------------------------------------------------------------------------
' GetNumber  -  safe numeric accessor on a Dictionary
' ---------------------------------------------------------------------------
Public Function GetNumber(ByVal dict As Object, ByVal key As String, _
                          Optional ByVal defaultVal As Double = 0) As Double
    If dict Is Nothing Then GetNumber = defaultVal : Exit Function
    If Not dict.Exists(key) Then GetNumber = defaultVal : Exit Function
    If IsNull(dict(key)) Then GetNumber = defaultVal : Exit Function
    If Not IsNumeric(dict(key)) Then GetNumber = defaultVal : Exit Function
    GetNumber = CDbl(dict(key))
End Function

' ---------------------------------------------------------------------------
' GetArray  -  safe Collection accessor on a Dictionary (JSON array)
' ---------------------------------------------------------------------------
Public Function GetArray(ByVal dict As Object, ByVal key As String) As Collection
    If dict Is Nothing Then Exit Function
    If Not dict.Exists(key) Then Exit Function
    If TypeName(dict(key)) <> "Collection" Then Exit Function
    Set GetArray = dict(key)
End Function

' ---------------------------------------------------------------------------
' GetObject  -  safe Dictionary accessor on a Dictionary (JSON object)
' ---------------------------------------------------------------------------
Public Function GetObject(ByVal dict As Object, ByVal key As String) As Object
    If dict Is Nothing Then Exit Function
    If Not dict.Exists(key) Then Exit Function
    If TypeName(dict(key)) <> "Dictionary" Then Exit Function
    Set GetObject = dict(key)
End Function

' ===========================================================================
' PRIVATE PARSER
' ===========================================================================

' Read the next JSON value starting at pPos
Private Function JsonReadValue() As Variant
    JsonSkipWhitespace
    If pPos > pLen Then JsonReadValue = Null : Exit Function

    Dim ch As String
    ch = Mid(pInput, pPos, 1)

    Select Case ch
        Case "{"
            Set JsonReadValue = JsonReadObject()
        Case "["
            Set JsonReadValue = JsonReadArray()
        Case """"
            JsonReadValue = JsonReadString()
        Case "-", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
            JsonReadValue = JsonReadNumber()
        Case "t"
            JsonReadValue = JsonReadLiteral("true", True)
        Case "f"
            JsonReadValue = JsonReadLiteral("false", False)
        Case "n"
            JsonReadValue = JsonReadLiteralNull()
        Case Else
            Err.Raise vbObjectError + 1001, "ModernJsonInVBA", _
                "Unexpected character '" & ch & "' at position " & pPos
    End Select
End Function

' Parse a JSON object {}  ->  Scripting.Dictionary
Private Function JsonReadObject() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    pPos = pPos + 1  ' consume "{"
    JsonSkipWhitespace

    If pPos <= pLen And Mid(pInput, pPos, 1) = "}" Then
        pPos = pPos + 1
        Set JsonReadObject = dict
        Exit Function
    End If

    Do
        JsonSkipWhitespace
        ' Key must be a quoted string
        If Mid(pInput, pPos, 1) <> """" Then
            Err.Raise vbObjectError + 1002, "ModernJsonInVBA", _
                "Expected string key at position " & pPos
        End If
        Dim key As String
        key = JsonReadString()

        JsonSkipWhitespace
        If Mid(pInput, pPos, 1) <> ":" Then
            Err.Raise vbObjectError + 1003, "ModernJsonInVBA", _
                "Expected ':' at position " & pPos
        End If
        pPos = pPos + 1  ' consume ":"

        Dim val As Variant
        JsonAssignVariant val, JsonReadValue()

        If IsObject(val) Then
            Set dict(key) = val
        Else
            dict(key) = val
        End If

        JsonSkipWhitespace
        Dim sep As String
        sep = Mid(pInput, pPos, 1)
        If sep = "}" Then
            pPos = pPos + 1 : Exit Do
        ElseIf sep = "," Then
            pPos = pPos + 1
        Else
            Err.Raise vbObjectError + 1004, "ModernJsonInVBA", _
                "Expected ',' or '}' at position " & pPos
        End If
    Loop

    Set JsonReadObject = dict
End Function

' Parse a JSON array []  ->  VBA Collection  (1-based)
Private Function JsonReadArray() As Collection
    Dim arr As New Collection

    pPos = pPos + 1  ' consume "["
    JsonSkipWhitespace

    If pPos <= pLen And Mid(pInput, pPos, 1) = "]" Then
        pPos = pPos + 1
        Set JsonReadArray = arr
        Exit Function
    End If

    Do
        Dim val As Variant
        JsonAssignVariant val, JsonReadValue()

        If IsObject(val) Then
            arr.Add val
        Else
            arr.Add val
        End If

        JsonSkipWhitespace
        Dim sep As String
        sep = Mid(pInput, pPos, 1)
        If sep = "]" Then
            pPos = pPos + 1 : Exit Do
        ElseIf sep = "," Then
            pPos = pPos + 1
        Else
            Err.Raise vbObjectError + 1005, "ModernJsonInVBA", _
                "Expected ',' or ']' at position " & pPos
        End If
    Loop

    Set JsonReadArray = arr
End Function

' Parse a JSON string (handles all standard escape sequences)
Private Function JsonReadString() As String
    pPos = pPos + 1  ' consume opening '"'

    Dim result As String
    result = ""

    Do While pPos <= pLen
        Dim ch As String
        ch = Mid(pInput, pPos, 1)

        If ch = """" Then
            pPos = pPos + 1  ' consume closing '"'
            JsonReadString = result
            Exit Function
        ElseIf ch = "\" Then
            pPos = pPos + 1
            If pPos > pLen Then Exit Do
            Dim esc As String
            esc = Mid(pInput, pPos, 1)
            Select Case esc
                Case """"  : result = result & """"
                Case "\"   : result = result & "\"
                Case "/"   : result = result & "/"
                Case "b"   : result = result & Chr(8)
                Case "f"   : result = result & Chr(12)
                Case "n"   : result = result & Chr(10)
                Case "r"   : result = result & Chr(13)
                Case "t"   : result = result & Chr(9)
                Case "u"
                    ' \uXXXX  Unicode escape
                    Dim hex4 As String
                    hex4 = Mid(pInput, pPos + 1, 4)
                    result = result & ChrW(CLng("&H" & hex4))
                    pPos = pPos + 4
                Case Else
                    result = result & esc
            End Select
        Else
            result = result & ch
        End If

        pPos = pPos + 1
    Loop

    Err.Raise vbObjectError + 1006, "ModernJsonInVBA", _
        "Unterminated string at position " & pPos
End Function

' Parse a JSON number -> Long if integer-valued, else Double
Private Function JsonReadNumber() As Variant
    Dim startPos As Long
    startPos = pPos
    Dim isFloat As Boolean
    isFloat = False

    If Mid(pInput, pPos, 1) = "-" Then pPos = pPos + 1  ' optional minus

    ' Integer digits
    Do While pPos <= pLen
        Dim c As String
        c = Mid(pInput, pPos, 1)
        If c >= "0" And c <= "9" Then : pPos = pPos + 1 : Else : Exit Do : End If
    Loop

    ' Optional fraction
    If pPos <= pLen And Mid(pInput, pPos, 1) = "." Then
        isFloat = True
        pPos = pPos + 1
        Do While pPos <= pLen
            c = Mid(pInput, pPos, 1)
            If c >= "0" And c <= "9" Then : pPos = pPos + 1 : Else : Exit Do : End If
        Loop
    End If

    ' Optional exponent
    If pPos <= pLen Then
        Dim e As String
        e = Mid(pInput, pPos, 1)
        If e = "e" Or e = "E" Then
            isFloat = True
            pPos = pPos + 1
            If pPos <= pLen Then
                Dim sign As String
                sign = Mid(pInput, pPos, 1)
                If sign = "+" Or sign = "-" Then pPos = pPos + 1
            End If
            Do While pPos <= pLen
                c = Mid(pInput, pPos, 1)
                If c >= "0" And c <= "9" Then : pPos = pPos + 1 : Else : Exit Do : End If
            Loop
        End If
    End If

    Dim numStr As String
    numStr = Mid(pInput, startPos, pPos - startPos)

    If isFloat Then
        JsonReadNumber = CDbl(numStr)
    Else
        Dim d As Double
        d = CDbl(numStr)
        If d >= -2147483648# And d <= 2147483647# Then
            JsonReadNumber = CLng(d)
        Else
            JsonReadNumber = d
        End If
    End If
End Function

' Parse  true  /  false
Private Function JsonReadLiteral(ByVal expected As String, ByVal result As Boolean) As Boolean
    If Mid(pInput, pPos, Len(expected)) = expected Then
        pPos = pPos + Len(expected)
        JsonReadLiteral = result
    Else
        Err.Raise vbObjectError + 1007, "ModernJsonInVBA", _
            "Expected '" & expected & "' at position " & pPos
    End If
End Function

' Parse  null  -> VBA Null
Private Function JsonReadLiteralNull() As Variant
    If Mid(pInput, pPos, 4) = "null" Then
        pPos = pPos + 4
        JsonReadLiteralNull = Null
    Else
        Err.Raise vbObjectError + 1008, "ModernJsonInVBA", _
            "Expected 'null' at position " & pPos
    End If
End Function

' Skip JSON whitespace characters
Private Sub JsonSkipWhitespace()
    Do While pPos <= pLen
        Dim ch As String
        ch = Mid(pInput, pPos, 1)
        If ch = " " Or ch = Chr(9) Or ch = Chr(10) Or ch = Chr(13) Then
            pPos = pPos + 1
        Else
            Exit Do
        End If
    Loop
End Sub

' Helper: assign any Variant (Object or primitive) without Set/Let confusion
Private Sub JsonAssignVariant(ByRef target As Variant, ByRef src As Variant)
    If IsObject(src) Then
        Set target = src
    Else
        target = src
    End If
End Sub
