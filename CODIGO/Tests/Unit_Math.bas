Attribute VB_Name = "Unit_Math"
Option Explicit

' ==========================================================================
' Math Test Suite
' Tests core math utilities: 2D vector operations (length, squared length,
' normalization, addition, subtraction, scalar multiplication), angle
' wrapping for negative and overflow values, linear interpolation at
' boundary factors, and point-inside-rectangle hit testing.
' ==========================================================================

#If UNIT_TEST = 1 Then

Public Function test_suite_math() As Boolean
    Call UnitTesting.RunTest("math_vec_length", test_vec_length())
    Call UnitTesting.RunTest("math_vec_sq_length", test_vec_sq_length())
    Call UnitTesting.RunTest("math_normalize", test_normalize())
    Call UnitTesting.RunTest("math_vadd", test_vadd())
    Call UnitTesting.RunTest("math_vsubs", test_vsubs())
    Call UnitTesting.RunTest("math_vmul", test_vmul())
    Call UnitTesting.RunTest("math_fix_angle_negative", test_fix_angle_negative())
    Call UnitTesting.RunTest("math_fix_angle_overflow", test_fix_angle_overflow())
    Call UnitTesting.RunTest("math_interpolate_bounds", test_interpolate_bounds())
    Call UnitTesting.RunTest("math_point_inside_rect", test_point_inside_rect())
    test_suite_math = True
End Function

' Verifies VecLength returns the Euclidean length of a vector.
' A (3,4) vector should have length 5.
Private Function test_vec_length() As Boolean
    On Error GoTo Fail
    ' Create a classic 3-4-5 right triangle vector
    Dim v As Vector2
    v.x = 3: v.y = 4
    ' sqrt(3^2 + 4^2) = sqrt(25) = 5
    test_vec_length = (Abs(VecLength(v) - 5!) < 0.0001)
    Exit Function
Fail:
    test_vec_length = False
End Function

' Verifies VecSqLength returns the squared length (avoids sqrt).
' A (3,4) vector should have squared length 25.
Private Function test_vec_sq_length() As Boolean
    On Error GoTo Fail
    ' Same 3-4-5 vector
    Dim v As Vector2
    v.x = 3: v.y = 4
    ' Squared length skips the sqrt: 3^2 + 4^2 = 25
    test_vec_sq_length = (Abs(VecSqLength(v) - 25!) < 0.0001)
    Exit Function
Fail:
    test_vec_sq_length = False
End Function

' Verifies Normalize scales a vector to unit length (~1.0).
Private Function test_normalize() As Boolean
    On Error GoTo Fail
    ' Start with a non-unit vector
    Dim v As Vector2
    v.x = 3: v.y = 4
    ' Normalize should scale it so its length becomes ~1.0
    Call Normalize(v)
    ' Check the resulting length is close to 1
    test_normalize = (Abs(VecLength(v) - 1!) < 0.001)
    Exit Function
Fail:
    test_normalize = False
End Function

' Verifies VAdd returns the component-wise sum of two vectors.
' (1,2) + (3,4) = (4,6)
Private Function test_vadd() As Boolean
    On Error GoTo Fail
    ' Set up two vectors to add
    Dim a As Vector2, b As Vector2, r As Vector2
    a.x = 1: a.y = 2
    b.x = 3: b.y = 4
    ' Add component-wise: (1+3, 2+4) = (4, 6)
    r = VAdd(a, b)
    test_vadd = (Abs(r.x - 4!) < 0.0001 And Abs(r.y - 6!) < 0.0001)
    Exit Function
Fail:
    test_vadd = False
End Function

' Verifies VSubs returns the component-wise difference of two vectors.
' (5,7) - (2,3) = (3,4)
Private Function test_vsubs() As Boolean
    On Error GoTo Fail
    ' Set up two vectors to subtract
    Dim a As Vector2, b As Vector2, r As Vector2
    a.x = 5: a.y = 7
    b.x = 2: b.y = 3
    ' Subtract component-wise: (5-2, 7-3) = (3, 4)
    r = VSubs(a, b)
    test_vsubs = (Abs(r.x - 3!) < 0.0001 And Abs(r.y - 4!) < 0.0001)
    Exit Function
Fail:
    test_vsubs = False
End Function

' Verifies VMul scales a vector by a scalar factor.
' (2,3) * 3 = (6,9)
Private Function test_vmul() As Boolean
    On Error GoTo Fail
    ' Set up a vector and a scalar multiplier
    Dim v As Vector2, r As Vector2
    v.x = 2: v.y = 3
    ' Multiply each component by 3: (2*3, 3*3) = (6, 9)
    r = VMul(v, 3!)
    test_vmul = (Abs(r.x - 6!) < 0.0001 And Abs(r.y - 9!) < 0.0001)
    Exit Function
Fail:
    test_vmul = False
End Function

' Verifies FixAngle wraps a negative angle into the 0-359 range.
' -45 degrees should become 315.
Private Function test_fix_angle_negative() As Boolean
    On Error GoTo Fail
    ' -45 is below 0, so it should wrap to 360 + (-45) = 315
    test_fix_angle_negative = (FixAngle(-45) = 315)
    Exit Function
Fail:
    test_fix_angle_negative = False
End Function

' Verifies FixAngle wraps an angle exceeding 360 back into range.
' 400 degrees should become 40.
Private Function test_fix_angle_overflow() As Boolean
    On Error GoTo Fail
    ' 400 exceeds 360, so it should wrap to 400 - 360 = 40
    test_fix_angle_overflow = (FixAngle(400) = 40)
    Exit Function
Fail:
    test_fix_angle_overflow = False
End Function

' Verifies Interpolate at boundary factors:
'   factor=0 returns the start value, factor=1 returns the end value.
Private Function test_interpolate_bounds() As Boolean
    On Error GoTo Fail
    ' factor=0 should return the start value (10)
    ' factor=1 should return the end value (20)
    test_interpolate_bounds = (Interpolate(10, 20, 0#) = 10 And Interpolate(10, 20, 1#) = 20)
    Exit Function
Fail:
    test_interpolate_bounds = False
End Function

' Verifies PointIsInsideRect returns True for a point within the rect
' and False for a point outside it.
Private Function test_point_inside_rect() As Boolean
    On Error GoTo Fail
    ' Define a rectangle from (10,10) to (50,50)
    Dim r As RECT
    r.Left = 10: r.Top = 10: r.Right = 50: r.Bottom = 50
    ' (25,25) is inside the rect
    Dim inside As Boolean
    inside = PointIsInsideRect(25, 25, r)
    ' (5,5) is outside the rect (top-left of it)
    Dim outside As Boolean
    outside = PointIsInsideRect(5, 5, r)
    ' First should be True, second should be False
    test_point_inside_rect = (inside = True And outside = False)
    Exit Function
Fail:
    test_point_inside_rect = False
End Function

#End If
