Attribute VB_Name = "Unit_Bitmask"
Option Explicit

' ==========================================================================
' Bitmask Test Suite
' Tests the bitwise flag system: setting individual bits, querying whether
' a bit is set, unsetting specific bits, resetting the entire mask to zero,
' and verifying correct behavior when multiple bits are set and selectively
' removed.
' ==========================================================================

#If UNIT_TEST = 1 Then

Public Sub test_suite_bitmask()
    Call UnitTesting.RunTest("bitmask_set", test_set_mask())
    Call UnitTesting.RunTest("bitmask_is_set", test_is_set())
    Call UnitTesting.RunTest("bitmask_unset", test_unset_mask())
    Call UnitTesting.RunTest("bitmask_reset", test_reset_mask())
    Call UnitTesting.RunTest("bitmask_multi_set_unset", test_multi_set_unset())
End Sub

' Verifies SetMask turns on the specified bit in a zeroed mask.
Private Function test_set_mask() As Boolean
    On Error GoTo Fail
    ' Start with an empty mask (all bits off)
    Dim m As Long: m = 0
    ' Turn on bit 4 (binary: 100)
    Call SetMask(m, 4)
    ' The mask should now equal 4 since only that bit is on
    test_set_mask = (m = 4)
    Exit Function
Fail:
    test_set_mask = False
End Function

' Verifies IsSet returns True for a bit that was set,
' and False for a bit that was not.
Private Function test_is_set() As Boolean
    On Error GoTo Fail
    ' Start with an empty mask
    Dim m As Long: m = 0
    ' Turn on only bit 4
    Call SetMask(m, 4)
    ' Bit 4 should be detected as set, bit 2 should not
    test_is_set = (IsSet(m, 4) = True And IsSet(m, 2) = False)
    Exit Function
Fail:
    test_is_set = False
End Function

' Verifies UnsetMask clears a specific bit while leaving others intact.
' Sets bits 2 and 4, unsets bit 2, expects only bit 4 to remain.
Private Function test_unset_mask() As Boolean
    On Error GoTo Fail
    ' Start empty
    Dim m As Long: m = 0
    ' Turn on bits 2 and 4 -> mask is now 6 (binary: 110)
    Call SetMask(m, 2)
    Call SetMask(m, 4)
    ' Remove bit 2 -> only bit 4 should remain (binary: 100 = 4)
    Call UnsetMask(m, 2)
    test_unset_mask = (m = 4)
    Exit Function
Fail:
    test_unset_mask = False
End Function

' Verifies ResetMask clears all bits back to zero.
Private Function test_reset_mask() As Boolean
    On Error GoTo Fail
    ' Start with bits 1+2+4 all on (binary: 111 = 7)
    Dim m As Long: m = 7
    ' Reset should clear everything back to zero
    Call ResetMask(m)
    test_reset_mask = (m = 0)
    Exit Function
Fail:
    test_reset_mask = False
End Function

' Verifies that setting two bits and unsetting one leaves only the
' other bit active (combined set/unset workflow).
Private Function test_multi_set_unset() As Boolean
    On Error GoTo Fail
    ' Start empty
    Dim m As Long: m = 0
    ' Turn on bits 2 and 8 -> mask is 10 (binary: 1010)
    Call SetMask(m, 2)
    Call SetMask(m, 8)
    ' Remove bit 2 -> only bit 8 should survive
    Call UnsetMask(m, 2)
    ' Confirm bit 8 is still on and bit 2 is gone
    test_multi_set_unset = (IsSet(m, 8) = True And IsSet(m, 2) = False)
    Exit Function
Fail:
    test_multi_set_unset = False
End Function

#End If
