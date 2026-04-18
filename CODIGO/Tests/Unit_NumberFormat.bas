Attribute VB_Name = "Unit_NumberFormat"
Option Explicit

' ==========================================================================
' NumberFormat Test Suite
' Tests PonerPuntos function from ModUtils.bas: thousand-separator formatting
' of Long values using dots as separators.
'
' Requirements: 3.1, 3.2, 3.3, 3.4
' ==========================================================================

#If UNIT_TEST = 1 Then

Public Sub test_suite_number_format()
    Call UnitTesting.RunTest("nf_zero", test_zero())
    Call UnitTesting.RunTest("nf_below_thousand", test_below_thousand())
    Call UnitTesting.RunTest("nf_one_thousand", test_one_thousand())
    Call UnitTesting.RunTest("nf_six_digits", test_six_digits())
    Call UnitTesting.RunTest("nf_one_million", test_one_million())
End Sub

' Requirement 3.4: PonerPuntos(0) returns "0"
Private Function test_zero() As Boolean
    On Error GoTo Fail
    test_zero = (PonerPuntos(0) = "0")
    Exit Function
Fail:
    test_zero = False
End Function

' Requirement 3.1: PonerPuntos(999) returns "999" (no separator)
Private Function test_below_thousand() As Boolean
    On Error GoTo Fail
    test_below_thousand = (PonerPuntos(999) = "999")
    Exit Function
Fail:
    test_below_thousand = False
End Function

' Requirement 3.2: PonerPuntos(1000) returns "1.000"
Private Function test_one_thousand() As Boolean
    On Error GoTo Fail
    test_one_thousand = (PonerPuntos(1000) = "1.000")
    Exit Function
Fail:
    test_one_thousand = False
End Function

' Requirement 3.2: PonerPuntos(999999) returns "999.999"
Private Function test_six_digits() As Boolean
    On Error GoTo Fail
    test_six_digits = (PonerPuntos(999999) = "999.999")
    Exit Function
Fail:
    test_six_digits = False
End Function

' Requirement 3.3: PonerPuntos(1000000) returns "1.000.000"
Private Function test_one_million() As Boolean
    On Error GoTo Fail
    test_one_million = (PonerPuntos(1000000) = "1.000.000")
    Exit Function
Fail:
    test_one_million = False
End Function

#End If
