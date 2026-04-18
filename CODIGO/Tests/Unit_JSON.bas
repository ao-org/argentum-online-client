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

' Requirement 1.1: Valid JSON object parsed into Dictionary with correct key-value pairs
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

' Requirement 1.2: Valid JSON array parsed into Collection with correct elements
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

' Requirement 1.3: Nested objects accessible via chained key lookups
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

' Requirement 1.4: Numeric values preserved as numeric VB6 types
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

' Requirement 1.5: Boolean literals parsed as VB6 Boolean values
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

' Requirement 1.6: Null value handled without raising an error
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

' Requirement 1.7: Empty object "{}" returns empty Dictionary with count zero
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

' Requirement 1.8: Empty array "[]" returns empty Collection with count zero
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

' Requirement 1.9: Invalid JSON populates GetParserErrors with non-empty string
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

' Requirement 1.10: Empty string returns Nothing or populates GetParserErrors
Private Function test_json_empty_string() As Boolean
    On Error GoTo Fail
    Dim jsonStr As String
    jsonStr = ""
    
    Dim result As Object
    Set result = JSON.parse(jsonStr)
    
    Dim errors As String
    errors = JSON.GetParserErrors()
    
    ' Either result is Nothing or errors is non-empty (or both)
    test_json_empty_string = (result Is Nothing) Or (Len(errors) > 0)
    Exit Function
Fail:
    test_json_empty_string = False
End Function

' Feature: unit-test-coverage, Property 3: JSON numeric value preservation
' **Validates: Requirements 1.4**
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
' **Validates: Requirements 1.9**
Private Function test_json_pbt_invalid_error() As Boolean
    On Error GoTo Fail
    
    Dim invalidStrings(0 To 119) As String
    Dim idx As Long
    Dim i As Long
    
    ' --- Category 1: Partial braces and brackets (8 items) ---
    invalidStrings(0) = "{"
    invalidStrings(1) = "}"
    invalidStrings(2) = "["
    invalidStrings(3) = "]"
    invalidStrings(4) = "{["
    invalidStrings(5) = "]}"
    invalidStrings(6) = "{[}"
    invalidStrings(7) = "[{]"
    
    ' --- Category 2: Plain words and phrases (8 items) ---
    invalidStrings(8) = "abc"
    invalidStrings(9) = "hello world"
    invalidStrings(10) = "not json at all"
    invalidStrings(11) = "undefined"
    invalidStrings(12) = "NaN"
    invalidStrings(13) = "Infinity"
    invalidStrings(14) = "TRUE"
    invalidStrings(15) = "FALSE"
    
    ' --- Category 3: Numbers without braces (8 items) ---
    invalidStrings(16) = "42"
    invalidStrings(17) = "-1"
    invalidStrings(18) = "3.14"
    invalidStrings(19) = "0"
    invalidStrings(20) = "99999"
    invalidStrings(21) = "-0.5"
    invalidStrings(22) = "1e10"
    invalidStrings(23) = "+7"
    
    ' --- Category 4: Malformed tokens (12 items) ---
    invalidStrings(24) = "{key}"
    invalidStrings(25) = "{:value}"
    invalidStrings(26) = "{key:value}"
    invalidStrings(27) = "{""key""}"
    invalidStrings(28) = "{""key"":}"
    invalidStrings(29) = "{:""value""}"
    invalidStrings(30) = "[,]"
    invalidStrings(31) = "{,}"
    invalidStrings(32) = "{""a"":1,}"
    invalidStrings(33) = "[1,2,]"
    invalidStrings(34) = "{""a"" ""b""}"
    invalidStrings(35) = "[1 2 3]"
    
    ' --- Category 5: Unquoted keys/values (8 items) ---
    invalidStrings(36) = "{a:1}"
    invalidStrings(37) = "{a:b}"
    invalidStrings(38) = "{""a"":b}"
    invalidStrings(39) = "{a:""b""}"
    invalidStrings(40) = "[a,b,c]"
    invalidStrings(41) = "{true:1}"
    invalidStrings(42) = "{null:null}"
    invalidStrings(43) = "{1:2}"
    
    ' --- Category 6: Truncated/incomplete JSON (8 items) ---
    invalidStrings(44) = "{""name"":"
    invalidStrings(45) = "{""name"":""val"
    invalidStrings(46) = "[1,2,"
    invalidStrings(47) = "{""a"":{""b"":"
    invalidStrings(48) = "[["
    invalidStrings(49) = "{{}"
    invalidStrings(50) = "{""a"":["
    invalidStrings(51) = "[{""a"":"
    
    ' --- Category 7: Special characters and symbols (8 items) ---
    invalidStrings(52) = "!@#$%"
    invalidStrings(53) = "<html>"
    invalidStrings(54) = "SELECT * FROM"
    invalidStrings(55) = "//comment"
    invalidStrings(56) = "/*block*/"
    invalidStrings(57) = "<?xml?>"
    invalidStrings(58) = "&amp;"
    invalidStrings(59) = "~~~"
    
    ' --- Category 8: Generate 60 more by combining patterns in a loop ---
    ' Pattern: "invalid_N" for N = 0 to 19 (plain words with numbers)
    For i = 0 To 19
        invalidStrings(60 + i) = "invalid_" & CStr(i)
    Next i
    
    ' Pattern: "{key_N" (unclosed brace with key) for N = 0 to 19
    For i = 0 To 19
        invalidStrings(80 + i) = "{key_" & CStr(i)
    Next i
    
    ' Pattern: "abc" repeated i+1 times with spaces for N = 0 to 19
    For i = 0 To 19
        invalidStrings(100 + i) = String$(i + 1, "x") & " " & CStr(i * 7)
    Next i
    
    ' --- Now test all 120 invalid strings ---
    Dim result As Object
    Dim errors As String
    
    For idx = 0 To 119
        Set result = JSON.parse(invalidStrings(idx))
        
        errors = JSON.GetParserErrors()
        
        If Len(errors) = 0 Then
            test_json_pbt_invalid_error = False
            Exit Function
        End If
    Next idx
    
    test_json_pbt_invalid_error = True
    Exit Function
Fail:
    test_json_pbt_invalid_error = False
End Function

#End If
