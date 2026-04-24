Attribute VB_Name = "Unit_ValidNumber"
Option Explicit

' ==========================================================================
' ValidNumber Test Suite
' Tests the ValidNumber function from ProtocolCmdParse.bas: boundary
' validation for ent_Byte, ent_Integer, ent_Long, and ent_Trigger types,
' plus non-numeric and empty string rejection.
' ==========================================================================

#If UNIT_TEST = 1 Then

Public Function test_suite_valid_number() As Boolean
    Call UnitTesting.RunTest("valid_byte_bounds", test_valid_byte_bounds())
    Call UnitTesting.RunTest("valid_integer_bounds", test_valid_integer_bounds())
    Call UnitTesting.RunTest("valid_long_bounds", test_valid_long_bounds())
    Call UnitTesting.RunTest("valid_trigger_bounds", test_valid_trigger_bounds())
    Call UnitTesting.RunTest("valid_non_numeric", test_valid_non_numeric())
    Call UnitTesting.RunTest("valid_empty_string", test_valid_empty_string())
    test_suite_valid_number = True
End Function

Private Function test_valid_byte_bounds() As Boolean
    On Error GoTo Fail
    Dim okMin As Boolean
    okMin = ValidNumber("0", eNumber_Types.ent_Byte)
    Dim okMax As Boolean
    okMax = ValidNumber("255", eNumber_Types.ent_Byte)
    Dim failBelow As Boolean
    failBelow = ValidNumber("-1", eNumber_Types.ent_Byte)
    Dim failAbove As Boolean
    failAbove = ValidNumber("256", eNumber_Types.ent_Byte)
    test_valid_byte_bounds = (okMin = True And okMax = True And failBelow = False And failAbove = False)
    Exit Function
Fail:
    test_valid_byte_bounds = False
End Function

Private Function test_valid_integer_bounds() As Boolean
    On Error GoTo Fail
    Dim okMin As Boolean
    okMin = ValidNumber("-32768", eNumber_Types.ent_Integer)
    Dim okMax As Boolean
    okMax = ValidNumber("32767", eNumber_Types.ent_Integer)
    Dim failBelow As Boolean
    failBelow = ValidNumber("-32769", eNumber_Types.ent_Integer)
    Dim failAbove As Boolean
    failAbove = ValidNumber("32768", eNumber_Types.ent_Integer)
    test_valid_integer_bounds = (okMin = True And okMax = True And failBelow = False And failAbove = False)
    Exit Function
Fail:
    test_valid_integer_bounds = False
End Function

Private Function test_valid_long_bounds() As Boolean
    On Error GoTo Fail
    Dim okMin As Boolean
    okMin = ValidNumber("-2147483648", eNumber_Types.ent_Long)
    Dim okMax As Boolean
    okMax = ValidNumber("2147483647", eNumber_Types.ent_Long)
    test_valid_long_bounds = (okMin = True And okMax = True)
    Exit Function
Fail:
    test_valid_long_bounds = False
End Function

Private Function test_valid_trigger_bounds() As Boolean
    On Error GoTo Fail
    Dim okMin As Boolean
    okMin = ValidNumber("0", eNumber_Types.ent_Trigger)
    Dim okMax As Boolean
    okMax = ValidNumber("99", eNumber_Types.ent_Trigger)
    Dim failBelow As Boolean
    failBelow = ValidNumber("-1", eNumber_Types.ent_Trigger)
    Dim failAbove As Boolean
    failAbove = ValidNumber("100", eNumber_Types.ent_Trigger)
    test_valid_trigger_bounds = (okMin = True And okMax = True And failBelow = False And failAbove = False)
    Exit Function
Fail:
    test_valid_trigger_bounds = False
End Function

Private Function test_valid_non_numeric() As Boolean
    On Error GoTo Fail
    Dim result As Boolean
    result = ValidNumber("abc", eNumber_Types.ent_Byte)
    test_valid_non_numeric = (result = False)
    Exit Function
Fail:
    test_valid_non_numeric = False
End Function

Private Function test_valid_empty_string() As Boolean
    On Error GoTo Fail
    Dim result As Boolean
    result = ValidNumber("", eNumber_Types.ent_Byte)
    test_valid_empty_string = (result = False)
    Exit Function
Fail:
    test_valid_empty_string = False
End Function

#End If
