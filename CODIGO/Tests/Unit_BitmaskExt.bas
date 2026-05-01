Attribute VB_Name = "Unit_BitmaskExt"
Option Explicit

' ==========================================================================
' Bitmask Extended Test Suite
' Additional tests for the bitwise flag system in Math.bas: SetMask, IsSet,
' UnsetMask, and ResetMask. Complements Unit_Bitmask.bas with extended
' example coverage and property-based round-trip verification.
' ==========================================================================

#If UNIT_TEST = 1 Then

Public Function test_suite_bitmask_ext() As Boolean
    Call UnitTesting.RunTest("bitmaskext_set", test_set_mask())
    Call UnitTesting.RunTest("bitmaskext_is_set", test_is_set())
    Call UnitTesting.RunTest("bitmaskext_unset", test_unset_mask())
    Call UnitTesting.RunTest("bitmaskext_reset", test_reset_mask())
    Call UnitTesting.RunTest("bitmaskext_prop_set_unset_roundtrip", test_prop_set_unset_roundtrip())
    test_suite_bitmask_ext = True
End Function

' Verifies SetMask on a zero mask with value 4 results in a mask of 4.
' Validates: Requirements 3.1
Private Function test_set_mask() As Boolean
    On Error GoTo Fail
    Dim m As Long: m = 0
    Call SetMask(m, 4)
    test_set_mask = (m = 4)
    Exit Function
Fail:
    test_set_mask = False
End Function

' Verifies IsSet returns True for a set bit and False for an unset bit.
' Validates: Requirements 3.2
Private Function test_is_set() As Boolean
    On Error GoTo Fail
    Dim m As Long: m = 0
    Call SetMask(m, 4)
    test_is_set = (IsSet(m, 4) = True And IsSet(m, 2) = False)
    Exit Function
Fail:
    test_is_set = False
End Function

' Verifies UnsetMask clears a previously set bit while leaving other bits intact.
' Sets bits 2 and 4, unsets bit 2, expects only bit 4 to remain.
' Validates: Requirements 3.3
Private Function test_unset_mask() As Boolean
    On Error GoTo Fail
    Dim m As Long: m = 0
    Call SetMask(m, 2)
    Call SetMask(m, 4)
    Call UnsetMask(m, 2)
    test_unset_mask = (m = 4 And IsSet(m, 4) = True And IsSet(m, 2) = False)
    Exit Function
Fail:
    test_unset_mask = False
End Function

' Verifies ResetMask sets the mask to 0 regardless of prior value.
' Validates: Requirements 3.4
Private Function test_reset_mask() As Boolean
    On Error GoTo Fail
    Dim m As Long: m = 255
    Call ResetMask(m)
    test_reset_mask = (m = 0)
    Exit Function
Fail:
    test_reset_mask = False
End Function

' Feature: full-coverage-unit-tests, Property 4: Long bitmask set/unset round-trip
' For any single-bit power-of-2 Long value v (bit positions 0-30), starting from
' a zero mask, after SetMask then UnsetMask, IsSet returns False.
' Validates: Requirements 3.5
Private Function test_prop_set_unset_roundtrip() As Boolean
    On Error GoTo Fail
    Dim i As Long
    For i = 1 To 120
        Dim bitPos As Long: bitPos = (i - 1) Mod 31  ' 0 to 30
        Dim v As Long: v = 2 ^ bitPos  ' single-bit power-of-2
        Dim m As Long: m = 0
        Call SetMask(m, v)
        Call UnsetMask(m, v)
        If IsSet(m, v) Then
            test_prop_set_unset_roundtrip = False
            Exit Function
        End If
    Next i
    test_prop_set_unset_roundtrip = True
    Exit Function
Fail:
    test_prop_set_unset_roundtrip = False
End Function

#End If
