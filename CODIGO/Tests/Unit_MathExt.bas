Attribute VB_Name = "Unit_MathExt"
Option Explicit

' ==========================================================================
' Extended Math Test Suite
' Tests additional Math.bas functions not covered by Unit_Math:
' GetAngle (angle between two points) and OverlapRect (rectangle overlap).
' ==========================================================================

#If UNIT_TEST = 1 Then

' Runs all extended math unit tests.
Public Function test_suite_math_ext() As Boolean
    Call UnitTesting.RunTest("mathext_angle_right", test_angle_right())
    Call UnitTesting.RunTest("mathext_angle_left", test_angle_left())
    Call UnitTesting.RunTest("mathext_angle_same", test_angle_same_point())
    Call UnitTesting.RunTest("mathext_overlap_true", test_overlap_true())
    Call UnitTesting.RunTest("mathext_overlap_false", test_overlap_false())
    test_suite_math_ext = True
End Function

' Verifies GetAngle returns approximately 0 radians when the target
' point is directly to the right (same Y, larger X).
Private Function test_angle_right() As Boolean
    On Error GoTo Fail
    ' Point to the right: (0,0) -> (1,0) should be 0 radians
    Dim angle As Double
    angle = GetAngle(0#, 0#, 1#, 0#)
    test_angle_right = (Abs(angle) < 0.001)
    Exit Function
Fail:
    test_angle_right = False
End Function

' Verifies GetAngle returns approximately pi radians when the target
' point is directly to the left (same Y, smaller X).
Private Function test_angle_left() As Boolean
    On Error GoTo Fail
    ' Point to the left: (0,0) -> (-1,0) should be pi radians
    Dim angle As Double
    angle = GetAngle(0#, 0#, -1#, 0#)
    test_angle_left = (Abs(angle - 3.14159265358979) < 0.001)
    Exit Function
Fail:
    test_angle_left = False
End Function

' Verifies GetAngle returns 0 when both points are identical.
Private Function test_angle_same_point() As Boolean
    On Error GoTo Fail
    Dim angle As Double
    angle = GetAngle(5#, 5#, 5#, 5#)
    test_angle_same_point = (angle = 0#)
    Exit Function
Fail:
    test_angle_same_point = False
End Function

' Verifies OverlapRect returns True for overlapping rectangles.
' Target rect (10,10)-(50,50), test rect at (30,30) size 40x40 overlaps.
Private Function test_overlap_true() As Boolean
    On Error GoTo Fail
    Dim r As RECT
    r.Left = 10: r.Top = 10: r.Right = 50: r.Bottom = 50
    ' Rectangle starting at (30,30) with width=40, height=40 overlaps
    test_overlap_true = (OverlapRect(r, 30, 30, 40, 40) = True)
    Exit Function
Fail:
    test_overlap_true = False
End Function

' Verifies OverlapRect returns False for completely separated rectangles.
' Target rect (10,10)-(50,50), test rect at (100,100) size 10x10 is outside.
Private Function test_overlap_false() As Boolean
    On Error GoTo Fail
    Dim r As RECT
    r.Left = 10: r.Top = 10: r.Right = 50: r.Bottom = 50
    ' Rectangle starting at (100,100) with width=10, height=10 is far away
    test_overlap_false = (OverlapRect(r, 100, 100, 10, 10) = False)
    Exit Function
Fail:
    test_overlap_false = False
End Function

#End If
