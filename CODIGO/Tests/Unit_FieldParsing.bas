Attribute VB_Name = "Unit_FieldParsing"
Option Explicit

' ==========================================================================
' Field Parsing Test Suite
' Tests the string field parsing functions in Mod_General: ReadField,
' FieldCount, General_Field_Read, and General_Field_Count. Covers
' positional extraction, empty fields, delimiter-free strings, and the
' round-trip property (Property 1).
' ==========================================================================

#If UNIT_TEST = 1 Then

Public Function test_suite_field_parsing() As Boolean
    Call UnitTesting.RunTest("field_read_positions", test_field_read_positions())
    Call UnitTesting.RunTest("field_read_no_delim", test_field_read_no_delim())
    Call UnitTesting.RunTest("field_read_empty_field", test_field_read_empty_field())
    Call UnitTesting.RunTest("field_count_multiple", test_field_count_multiple())
    Call UnitTesting.RunTest("field_count_empty", test_field_count_empty())
    Call UnitTesting.RunTest("field_count_no_delim", test_field_count_no_delim())
    Call UnitTesting.RunTest("general_field_read_multi_char", test_general_field_read_multi_char())
    Call UnitTesting.RunTest("general_field_count_cases", test_general_field_count_cases())
    Call UnitTesting.RunTest("field_round_trip", test_field_round_trip())
    test_suite_field_parsing = True
End Function

' **Validates: Requirements 1.1**
Private Function test_field_read_positions() As Boolean
    On Error GoTo Fail
    Dim text As String
    text = "A,B,C"
    Dim delim As Byte
    delim = 44
    test_field_read_positions = (ReadField(1, text, delim) = "A") And _
                                (ReadField(2, text, delim) = "B") And _
                                (ReadField(3, text, delim) = "C")
    Exit Function
Fail:
    test_field_read_positions = False
End Function

' **Validates: Requirements 1.2**
Private Function test_field_read_no_delim() As Boolean
    On Error GoTo Fail
    Dim text As String
    text = "Hello"
    test_field_read_no_delim = (ReadField(1, text, 44) = "Hello")
    Exit Function
Fail:
    test_field_read_no_delim = False
End Function

' **Validates: Requirements 1.3**
Private Function test_field_read_empty_field() As Boolean
    On Error GoTo Fail
    Dim text As String
    text = "A,,C"
    test_field_read_empty_field = (ReadField(2, text, 44) = "")
    Exit Function
Fail:
    test_field_read_empty_field = False
End Function

' **Validates: Requirements 1.4**
Private Function test_field_count_multiple() As Boolean
    On Error GoTo Fail
    Dim text As String
    text = "A,B,C"
    test_field_count_multiple = (FieldCount(text, 44) = 3)
    Exit Function
Fail:
    test_field_count_multiple = False
End Function

' **Validates: Requirements 1.5**
Private Function test_field_count_empty() As Boolean
    On Error GoTo Fail
    Dim text As String
    text = ""
    test_field_count_empty = (FieldCount(text, 44) = 0)
    Exit Function
Fail:
    test_field_count_empty = False
End Function

' **Validates: Requirements 1.6**
Private Function test_field_count_no_delim() As Boolean
    On Error GoTo Fail
    Dim text As String
    text = "Hello"
    test_field_count_no_delim = (FieldCount(text, 44) = 1)
    Exit Function
Fail:
    test_field_count_no_delim = False
End Function

' **Validates: Requirements 1.7**
Private Function test_general_field_read_multi_char() As Boolean
    On Error GoTo Fail
    test_general_field_read_multi_char = (General_Field_Read(2, "A-B-C", "-") = "B")
    Exit Function
Fail:
    test_general_field_read_multi_char = False
End Function

' **Validates: Requirements 1.8**
Private Function test_general_field_count_cases() As Boolean
    On Error GoTo Fail
    Dim text As String
    text = "A-B-C"
    Dim empty_text As String
    empty_text = ""
    test_general_field_count_cases = (General_Field_Count(text, 45) = 3) And _
                                     (General_Field_Count(empty_text, 45) = 0)
    Exit Function
Fail:
    test_general_field_count_cases = False
End Function

' Feature: unit-test-coverage-tier3, Property 1: ReadField/FieldCount round-trip consistency
' **Validates: Requirements 1.9**
Private Function test_field_round_trip() As Boolean
    On Error GoTo Fail
    Dim delim As Byte
    delim = 44
    Dim sep As String
    sep = Chr$(delim)

    ' Case 1: "A,B,C"
    Dim t1 As String
    t1 = "A,B,C"
    Dim n1 As Long
    n1 = FieldCount(t1, delim)
    Dim rebuilt1 As String
    Dim i As Long
    For i = 1 To n1
        If i > 1 Then rebuilt1 = rebuilt1 & sep
        rebuilt1 = rebuilt1 & ReadField(CInt(i), t1, delim)
    Next i
    If rebuilt1 <> t1 Then
        test_field_round_trip = False
        Exit Function
    End If

    ' Case 2: "X" (single field, no delimiters)
    Dim t2 As String
    t2 = "X"
    Dim n2 As Long
    n2 = FieldCount(t2, delim)
    Dim rebuilt2 As String
    For i = 1 To n2
        If i > 1 Then rebuilt2 = rebuilt2 & sep
        rebuilt2 = rebuilt2 & ReadField(CInt(i), t2, delim)
    Next i
    If rebuilt2 <> t2 Then
        test_field_round_trip = False
        Exit Function
    End If

    ' Case 3: "1,2,,4" (empty field between adjacent delimiters)
    Dim t3 As String
    t3 = "1,2,,4"
    Dim n3 As Long
    n3 = FieldCount(t3, delim)
    Dim rebuilt3 As String
    For i = 1 To n3
        If i > 1 Then rebuilt3 = rebuilt3 & sep
        rebuilt3 = rebuilt3 & ReadField(CInt(i), t3, delim)
    Next i
    If rebuilt3 <> t3 Then
        test_field_round_trip = False
        Exit Function
    End If

    test_field_round_trip = True
    Exit Function
Fail:
    test_field_round_trip = False
End Function

#End If
