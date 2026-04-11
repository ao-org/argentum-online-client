Attribute VB_Name = "Unit_Math"
Option Explicit

#If UNIT_TEST = 1 Then

Public Sub test_suite_math()
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
End Sub

Private Function test_vec_length() As Boolean
    On Error GoTo Fail
    Dim v As Vector2
    v.x = 3: v.y = 4
    test_vec_length = (Abs(VecLength(v) - 5!) < 0.0001)
    Exit Function
Fail:
    test_vec_length = False
End Function

Private Function test_vec_sq_length() As Boolean
    On Error GoTo Fail
    Dim v As Vector2
    v.x = 3: v.y = 4
    test_vec_sq_length = (Abs(VecSqLength(v) - 25!) < 0.0001)
    Exit Function
Fail:
    test_vec_sq_length = False
End Function

Private Function test_normalize() As Boolean
    On Error GoTo Fail
    Dim v As Vector2
    v.x = 3: v.y = 4
    Call Normalize(v)
    test_normalize = (Abs(VecLength(v) - 1!) < 0.001)
    Exit Function
Fail:
    test_normalize = False
End Function

Private Function test_vadd() As Boolean
    On Error GoTo Fail
    Dim a As Vector2, b As Vector2, r As Vector2
    a.x = 1: a.y = 2
    b.x = 3: b.y = 4
    r = VAdd(a, b)
    test_vadd = (Abs(r.x - 4!) < 0.0001 And Abs(r.y - 6!) < 0.0001)
    Exit Function
Fail:
    test_vadd = False
End Function

Private Function test_vsubs() As Boolean
    On Error GoTo Fail
    Dim a As Vector2, b As Vector2, r As Vector2
    a.x = 5: a.y = 7
    b.x = 2: b.y = 3
    r = VSubs(a, b)
    test_vsubs = (Abs(r.x - 3!) < 0.0001 And Abs(r.y - 4!) < 0.0001)
    Exit Function
Fail:
    test_vsubs = False
End Function

Private Function test_vmul() As Boolean
    On Error GoTo Fail
    Dim v As Vector2, r As Vector2
    v.x = 2: v.y = 3
    r = VMul(v, 3!)
    test_vmul = (Abs(r.x - 6!) < 0.0001 And Abs(r.y - 9!) < 0.0001)
    Exit Function
Fail:
    test_vmul = False
End Function

Private Function test_fix_angle_negative() As Boolean
    On Error GoTo Fail
    test_fix_angle_negative = (FixAngle(-45) = 315)
    Exit Function
Fail:
    test_fix_angle_negative = False
End Function

Private Function test_fix_angle_overflow() As Boolean
    On Error GoTo Fail
    test_fix_angle_overflow = (FixAngle(400) = 40)
    Exit Function
Fail:
    test_fix_angle_overflow = False
End Function

Private Function test_interpolate_bounds() As Boolean
    On Error GoTo Fail
    test_interpolate_bounds = (Interpolate(10, 20, 0#) = 10 And Interpolate(10, 20, 1#) = 20)
    Exit Function
Fail:
    test_interpolate_bounds = False
End Function

Private Function test_point_inside_rect() As Boolean
    On Error GoTo Fail
    Dim r As RECT
    r.Left = 10: r.Top = 10: r.Right = 50: r.Bottom = 50
    Dim inside As Boolean
    inside = PointIsInsideRect(25, 25, r)
    Dim outside As Boolean
    outside = PointIsInsideRect(5, 5, r)
    test_point_inside_rect = (inside = True And outside = False)
    Exit Function
Fail:
    test_point_inside_rect = False
End Function

#End If
