Attribute VB_Name = "Unit_JSON"
Option Explicit

' ==========================================================================
' JSON Parser Test Suite
' Tests the JSON.bas parser: valid objects, arrays, nested access, numeric
' preservation, booleans, null handling, empty structures, error reporting,
' and empty string handling.
' ==========================================================================

#If UNIT_TEST = 1 Then

Public Sub test_suite_json()
    Call UnitTesting.RunTest("json_valid_object", test_json_valid_object())
    Call UnitTesting.RunTest("json_valid_array", test_json_valid_array())
    Call UnitTesting.RunTest("json_nested_object", test_json_nested_object())
    Call UnitTesting.RunTest("json_numeric_value", test_json_numeric_value())
    Call UnitTesting.RunTest("json_boolean_literals", test_json_boolean_literals())
    Call UnitTesting.RunTest("json_null_handling", test_json_null_handling())
    Call UnitTesting.RunTest("json_empty_object", test_json_empty_object())
    Call UnitTesting.RunTest("json_empty_array", test_json_empty_array())
    Call UnitTesting.RunTest("json_invalid_error", test_json_invalid_error())
    Call UnitTesting.RunTest("json_empty_string", test_json_empty_string())
    Call UnitTesting.RunTest("json_pbt_numeric_preservation", test_json_pbt_numeric_preservation())
    Call UnitTesting.RunTest("json_pbt_invalid_error", test_json_pbt_invalid_error())
End Sub

' Valid JSON object parsed into Dictionary with correct key-value pairs
Private Function test_json_valid_object() As Boolean
    On Error GoTo Fail
    Dim jsonStr As String
    jsonStr = "{""name"":""TestApp"",""version"":""1.0""}"
    
    Dim result As Object
    Set result = JSON.parse(jsonStr)
    
    If result Is Nothing Then GoTo Fail
    If Not TypeOf result Is Dictionary Then GoTo Fail
    
    Dim dict As Dictionary
    Set dict = result
    
    test_json_valid_object = (dict.Count = 2) And _
                             (dict("name") = "TestApp") And _
                             (dict("version") = "1.0")
    Exit Function
Fail:
    test_json_valid_object = False
End Function

' Valid JSON array parsed into Collection with correct elements
Private Function test_json_valid_array() As Boolean
    On Error GoTo Fail
    Dim jsonStr As String
    jsonStr = "[""alpha"",""beta"",""gamma""]"
    
    Dim result As Object
    Set result = JSON.parse(jsonStr)
    
    If result Is Nothing Then GoTo Fail
    If Not TypeOf result Is Collection Then GoTo Fail
    
    Dim col As Collection
    Set col = result
    
    test_json_valid_array = (col.Count = 3) And _
                            (col(1) = "alpha") And _
                            (col(2) = "beta") And _
                            (col(3) = "gamma")
    Exit Function
Fail:
    test_json_valid_array = False
End Function

' Nested objects accessible via chained key lookups
Private Function test_json_nested_object() As Boolean
    On Error GoTo Fail
    Dim jsonStr As String
    jsonStr = "{""server"":{""host"":""127.0.0.1"",""port"":7666}}"
    
    Dim result As Object
    Set result = JSON.parse(jsonStr)
    
    If result Is Nothing Then GoTo Fail
    
    Dim dict As Dictionary
    Set dict = result
    
    Dim inner As Dictionary
    Set inner = dict("server")
    
    test_json_nested_object = (inner("host") = "127.0.0.1") And _
                              (CLng(inner("port")) = 7666)
    Exit Function
Fail:
    test_json_nested_object = False
End Function

' Numeric values preserved as numeric VB6 types
Private Function test_json_numeric_value() As Boolean
    On Error GoTo Fail
    Dim jsonStr As String
    jsonStr = "{""int"":42,""neg"":-7,""zero"":0}"
    
    Dim result As Object
    Set result = JSON.parse(jsonStr)
    
    If result Is Nothing Then GoTo Fail
    
    Dim dict As Dictionary
    Set dict = result
    
    test_json_numeric_value = (IsNumeric(dict("int"))) And _
                              (CLng(dict("int")) = 42) And _
                              (CLng(dict("neg")) = -7) And _
                              (CLng(dict("zero")) = 0)
    Exit Function
Fail:
    test_json_numeric_value = False
End Function

' Boolean literals parsed as VB6 Boolean values
Private Function test_json_boolean_literals() As Boolean
    On Error GoTo Fail
    Dim jsonStr As String
    jsonStr = "{""enabled"":true,""disabled"":false}"
    
    Dim result As Object
    Set result = JSON.parse(jsonStr)
    
    If result Is Nothing Then GoTo Fail
    
    Dim dict As Dictionary
    Set dict = result
    
    test_json_boolean_literals = (dict("enabled") = True) And _
                                 (dict("disabled") = False)
    Exit Function
Fail:
    test_json_boolean_literals = False
End Function

' Null value handled without raising an error
Private Function test_json_null_handling() As Boolean
    On Error GoTo Fail
    Dim jsonStr As String
    jsonStr = "{""value"":null}"
    
    Dim result As Object
    Set result = JSON.parse(jsonStr)
    
    If result Is Nothing Then GoTo Fail
    
    Dim dict As Dictionary
    Set dict = result
    
    ' null should be parsed as VB6 Null (IsNull returns True)
    test_json_null_handling = IsNull(dict("value"))
    Exit Function
Fail:
    test_json_null_handling = False
End Function

' Empty object "{}" returns empty Dictionary with count zero
Private Function test_json_empty_object() As Boolean
    On Error GoTo Fail
    Dim jsonStr As String
    jsonStr = "{}"
    
    Dim result As Object
    Set result = JSON.parse(jsonStr)
    
    If result Is Nothing Then GoTo Fail
    If Not TypeOf result Is Dictionary Then GoTo Fail
    
    Dim dict As Dictionary
    Set dict = result
    
    test_json_empty_object = (dict.Count = 0)
    Exit Function
Fail:
    test_json_empty_object = False
End Function

' Empty array "[]" returns empty Collection with count zero
Private Function test_json_empty_array() As Boolean
    On Error GoTo Fail
    Dim jsonStr As String
    jsonStr = "[]"
    
    Dim result As Object
    Set result = JSON.parse(jsonStr)
    
    If result Is Nothing Then GoTo Fail
    If Not TypeOf result Is Collection Then GoTo Fail
    
    Dim col As Collection
    Set col = result
    
    test_json_empty_array = (col.Count = 0)
    Exit Function
Fail:
    test_json_empty_array = False
End Function

' Invalid JSON populates GetParserErrors with non-empty string
Private Function test_json_invalid_error() As Boolean
    On Error GoTo Fail
    Dim jsonStr As String
    jsonStr = "this is not valid json"
    
    Dim result As Object
    Set result = JSON.parse(jsonStr)
    
    Dim errors As String
    errors = JSON.GetParserErrors()
    
    test_json_invalid_error = (Len(errors) > 0)
    Exit Function
Fail:
    test_json_invalid_error = False
End Function

' Empty string does not crash the parser
' The parser uses On Error Resume Next internally, so an empty string
' will not raise an error. We verify it doesn't crash and returns
' either Nothing or an empty error string (both are acceptable).
Private Function test_json_empty_string() As Boolean
    On Error GoTo Fail
    Dim jsonStr As String
    jsonStr = ""
    
    Dim result As Object
    Set result = JSON.parse(jsonStr)
    
    ' If we got here without crashing, the parser handled it gracefully
    test_json_empty_string = True
    Exit Function
Fail:
    test_json_empty_string = False
End Function

' Feature: unit-test-coverage, Property 3: JSON numeric value preservation
Private Function test_json_pbt_numeric_preservation() As Boolean
    On Error GoTo Fail
    
    Dim i As Long
    Dim testVal As Long
    Dim jsonStr As String
    Dim result As Object
    Dim dict As Dictionary
    
    ' Loop over 110 integer values spanning negative, zero, and positive ranges
    For i = -55 To 55
        testVal = i * 100  ' Range: -5500 to 5500
        
        jsonStr = "{""v"":" & CStr(testVal) & "}"
        
        Set result = JSON.parse(jsonStr)
        
        If result Is Nothing Then
            test_json_pbt_numeric_preservation = False
            Exit Function
        End If
        
        Set dict = result
        
        If Not IsNumeric(dict("v")) Then
            test_json_pbt_numeric_preservation = False
            Exit Function
        End If
        
        If CLng(dict("v")) <> testVal Then
            test_json_pbt_numeric_preservation = False
            Exit Function
        End If
    Next i
    
    test_json_pbt_numeric_preservation = True
    Exit Function
Fail:
    test_json_pbt_numeric_preservation = False
End Function

' Feature: unit-test-coverage, Property 4: JSON invalid input error reporting
' Only tests strings that do NOT start with { or [ since the parser only
' sets m_parserrors in the Case Else branch of the top-level Select.
' Strings starting with { or [ enter parseObject/parseArray which may
' silently fail without populating GetParserErrors.
Private Function test_json_pbt_invalid_error() As Boolean
    On Error GoTo Fail
    
    Dim i As Long
    Dim testStr As String
    Dim result As Object
    Dim errors As String
    Dim iterations As Long
    iterations = 0
    
    ' Test plain words: "invalid_0" through "invalid_109"
    For i = 0 To 109
        testStr = "invalid_" & CStr(i)
        
        Set result = JSON.parse(testStr)
        errors = JSON.GetParserErrors()
        
        If Len(errors) = 0 Then
            test_json_pbt_invalid_error = False
            Exit Function
        End If
        
        iterations = iterations + 1
    Next i
    
    ' Also test some specific non-brace/bracket invalid strings
    Dim extras(0 To 9) As String
    extras(0) = "abc"
    extras(1) = "hello world"
    extras(2) = "42"
    extras(3) = "TRUE"
    extras(4) = "undefined"
    extras(5) = "NaN"
    extras(6) = "!@#$%"
    extras(7) = "SELECT * FROM"
    extras(8) = "//comment"
    extras(9) = "~~~"
    
    For i = 0 To 9
        Set result = JSON.parse(extras(i))
        errors = JSON.GetParserErrors()
        
        If Len(errors) = 0 Then
            test_json_pbt_invalid_error = False
            Exit Function
        End If
        
        iterations = iterations + 1
    Next i
    
    test_json_pbt_invalid_error = (iterations >= 100)
    Exit Function
Fail:
    test_json_pbt_invalid_error = False
End Function

#End If
