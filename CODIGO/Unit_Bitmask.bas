Attribute VB_Name = "Unit_Bitmask"
Option Explicit

#If UNIT_TEST = 1 Then

Public Sub test_suite_bitmask()
    Call UnitTesting.RunTest("bitmask_set", test_set_mask())
    Call UnitTesting.RunTest("bitmask_is_set", test_is_set())
    Call UnitTesting.RunTest("bitmask_unset", test_unset_mask())
    Call UnitTesting.RunTest("bitmask_reset", test_reset_mask())
    Call UnitTesting.RunTest("bitmask_multi_set_unset", test_multi_set_unset())
End Sub

Private Function test_set_mask() As Boolean
    On Error GoTo Fail
    Dim m As Long: m = 0
    Call SetMask(m, 4)
    test_set_mask = (m = 4)
    Exit Function
Fail:
    test_set_mask = False
End Function

Private Function test_is_set() As Boolean
    On Error GoTo Fail
    Dim m As Long: m = 0
    Call SetMask(m, 4)
    test_is_set = (IsSet(m, 4) = True And IsSet(m, 2) = False)
    Exit Function
Fail:
    test_is_set = False
End Function

Private Function test_unset_mask() As Boolean
    On Error GoTo Fail
    Dim m As Long: m = 0
    Call SetMask(m, 2)
    Call SetMask(m, 4)
    Call UnsetMask(m, 2)
    test_unset_mask = (m = 4)
    Exit Function
Fail:
    test_unset_mask = False
End Function

Private Function test_reset_mask() As Boolean
    On Error GoTo Fail
    Dim m As Long: m = 7
    Call ResetMask(m)
    test_reset_mask = (m = 0)
    Exit Function
Fail:
    test_reset_mask = False
End Function

Private Function test_multi_set_unset() As Boolean
    On Error GoTo Fail
    Dim m As Long: m = 0
    Call SetMask(m, 2)
    Call SetMask(m, 8)
    Call UnsetMask(m, 2)
    test_multi_set_unset = (IsSet(m, 8) = True And IsSet(m, 2) = False)
    Exit Function
Fail:
    test_multi_set_unset = False
End Function

#End If
