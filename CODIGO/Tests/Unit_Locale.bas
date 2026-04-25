Attribute VB_Name = "Unit_Locale"
Option Explicit

' ==========================================================================
' Locale Encoding Test Suite
' Tests string/data conversion round-trips: Integer_To_String <-> String_To_Integer,
' Byte_To_String <-> String_To_Byte, Long_To_String <-> String_To_Long,
' and edge cases for empty strings and out-of-bounds positions.
' ==========================================================================

#If UNIT_TEST = 1 Then

' Runs all locale encoding unit tests.
Public Function test_suite_locale() As Boolean
    Call UnitTesting.RunTest("locale_integer_round_trip", test_integer_round_trip())
    Call UnitTesting.RunTest("locale_byte_round_trip", test_byte_round_trip())
    Call UnitTesting.RunTest("locale_long_round_trip", test_long_round_trip())
    Call UnitTesting.RunTest("locale_str_to_int_empty", test_string_to_integer_empty())
    Call UnitTesting.RunTest("locale_str_to_byte_oob", test_string_to_byte_oob())
    test_suite_locale = True
End Function

' Verifies Integer_To_String followed by String_To_Integer returns
' the original value for representative integers.
Private Function test_integer_round_trip() As Boolean
    On Error GoTo Fail
    Dim ok As Boolean: ok = True
    Dim v As Variant
    Dim vals As Variant: vals = Array(1, 255, 1000)
    Dim i As Long
    For i = 0 To UBound(vals)
        v = vals(i)
        Dim encoded As String
        encoded = Integer_To_String(CInt(v))
        Dim decoded As Integer
        decoded = String_To_Integer(encoded, 1)
        If decoded <> CInt(v) Then ok = False
    Next i
    test_integer_round_trip = ok
    Exit Function
Fail:
    test_integer_round_trip = False
End Function

' Verifies Byte_To_String followed by String_To_Byte returns
' the original value for representative byte values.
Private Function test_byte_round_trip() As Boolean
    On Error GoTo Fail
    Dim ok As Boolean: ok = True
    Dim vals As Variant: vals = Array(1, 127, 255)
    Dim i As Long
    For i = 0 To UBound(vals)
        Dim encoded As String
        encoded = Byte_To_String(CByte(vals(i)))
        Dim decoded As Byte
        decoded = String_To_Byte(encoded, 1)
        If decoded <> CByte(vals(i)) Then ok = False
    Next i
    test_byte_round_trip = ok
    Exit Function
Fail:
    test_byte_round_trip = False
End Function

' Verifies Long_To_String followed by String_To_Long returns
' the original value for a representative non-zero Long.
Private Function test_long_round_trip() As Boolean
    On Error GoTo Fail
    Dim v As Long: v = 123456
    Dim encoded As String
    encoded = Long_To_String(v)
    Dim decoded As Long
    decoded = String_To_Long(encoded, 1)
    test_long_round_trip = (decoded = v)
    Exit Function
Fail:
    test_long_round_trip = False
End Function

' Verifies String_To_Integer returns 0 when given an empty string.
Private Function test_string_to_integer_empty() As Boolean
    On Error GoTo Fail
    test_string_to_integer_empty = (String_To_Integer("", 1) = 0)
    Exit Function
Fail:
    test_string_to_integer_empty = False
End Function

' Verifies String_To_Byte returns 0 when the start position
' exceeds the string length.
Private Function test_string_to_byte_oob() As Boolean
    On Error GoTo Fail
    test_string_to_byte_oob = (String_To_Byte("A", 5) = 0)
    Exit Function
Fail:
    test_string_to_byte_oob = False
End Function

#End If
