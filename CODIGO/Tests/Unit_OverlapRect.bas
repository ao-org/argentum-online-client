Attribute VB_Name = "Unit_OverlapRect"
Option Explicit

' ==========================================================================
' OverlapRect Test Suite
' Tests Math.bas OverlapRect function: verifies rectangle overlap detection
' for contained, partially overlapping, non-overlapping, and edge-sharing
' cases. Complements Unit_Math.bas with focused OverlapRect coverage.
' ==========================================================================

#If UNIT_TEST = 1 Then

Public Sub test_suite_overlap_rect()
    Call UnitTesting.RunTest("overlap_contained", test_overlap_contained())
    Call UnitTesting.RunTest("overlap_partial", test_overlap_partial())
    Call UnitTesting.RunTest("overlap_none", test_overlap_none())
    Call UnitTesting.RunTest("overlap_edge", test_overlap_edge())
End Sub

' Verifies OverlapRect returns True when the target rectangle fully contains
' the second rectangle (all four corners of the second rect are inside target).
' Validates: Requirements 11.1
Private Function test_overlap_contained() As Boolean
    On Error GoTo Fail
    Dim target As RECT
    target.Left = 0: target.Top = 0: target.Right = 100: target.Bottom = 100
    ' Second rect (10, 10, 20, 20) is fully inside target
    test_overlap_contained = (OverlapRect(target, 10, 10, 20, 20) = True)
    Exit Function
Fail:
    test_overlap_contained = False
End Function

' Verifies OverlapRect returns True when two rectangles partially overlap
' (at least one corner of the second rect is inside the target).
' Validates: Requirements 11.2
Private Function test_overlap_partial() As Boolean
    On Error GoTo Fail
    Dim target As RECT
    target.Left = 10: target.Top = 10: target.Right = 50: target.Bottom = 50
    ' Second rect starts at (40, 40) with size 30x30, so top-left corner (40,40)
    ' is inside target, but bottom-right (70,70) extends beyond
    test_overlap_partial = (OverlapRect(target, 40, 40, 30, 30) = True)
    Exit Function
Fail:
    test_overlap_partial = False
End Function

' Verifies OverlapRect returns False when two rectangles do not overlap at all.
' Validates: Requirements 11.3
Private Function test_overlap_none() As Boolean
    On Error GoTo Fail
    Dim target As RECT
    target.Left = 10: target.Top = 10: target.Right = 50: target.Bottom = 50
    ' Second rect at (200, 200) with size 10x10 is far away from target
    test_overlap_none = (OverlapRect(target, 200, 200, 10, 10) = False)
    Exit Function
Fail:
    test_overlap_none = False
End Function

' Verifies OverlapRect behavior when a rectangle shares only an edge with the
' target, testing the corner-check logic. The second rect is placed so its
' right edge aligns with the target's left edge — the corner (x+Width, y) sits
' exactly on the target boundary, which PointIsInsideRect includes (>=, <=).
' Validates: Requirements 11.4
Private Function test_overlap_edge() As Boolean
    On Error GoTo Fail
    Dim target As RECT
    target.Left = 50: target.Top = 0: target.Right = 100: target.Bottom = 100
    ' Second rect (30, 0, 20, 100): right edge at x=50 touches target.Left=50
    ' Corner (50, 0) is checked by PointIsInsideRect: 50 >= 50 And 50 <= 100
    ' And 0 >= 0 And 0 <= 100 ? True, so OverlapRect returns True
    test_overlap_edge = (OverlapRect(target, 30, 0, 20, 100) = True)
    Exit Function
Fail:
    test_overlap_edge = False
End Function

#End If
