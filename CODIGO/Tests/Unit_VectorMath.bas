Attribute VB_Name = "Unit_VectorMath"
Option Explicit

' ==========================================================================
' VectorMath Test Suite
' Tests Math.bas vector and geometry functions: VecLength, VecSqLength,
' Normalize, VAdd, VSubs, VMul, GetAngle, FixAngle, Interpolate, and
' PointIsInsideRect with example-based tests covering known values,
' boundary conditions, and edge cases.
' ==========================================================================

#If UNIT_TEST = 1 Then

Public Sub test_suite_vector_math()
    Call UnitTesting.RunTest("vecmath_length_3_4", test_vec_length_3_4())
    Call UnitTesting.RunTest("vecmath_length_zero", test_vec_length_zero())
    Call UnitTesting.RunTest("vecmath_sq_length", test_vec_sq_length())
    Call UnitTesting.RunTest("vecmath_normalize", test_vec_normalize())
    Call UnitTesting.RunTest("vecmath_vadd", test_vec_add())
    Call UnitTesting.RunTest("vecmath_vsubs", test_vec_subs())
    Call UnitTesting.RunTest("vecmath_vmul", test_vec_mul())
    Call UnitTesting.RunTest("vecmath_getangle_same", test_getangle_same())
    Call UnitTesting.RunTest("vecmath_getangle_pos_x", test_getangle_pos_x())
    Call UnitTesting.RunTest("vecmath_getangle_neg_x", test_getangle_neg_x())
    Call UnitTesting.RunTest("vecmath_fixangle_neg", test_fixangle_neg())
    Call UnitTesting.RunTest("vecmath_fixangle_overflow", test_fixangle_overflow())
    Call UnitTesting.RunTest("vecmath_interpolate_mid", test_interpolate_mid())
    Call UnitTesting.RunTest("vecmath_interpolate_bounds", test_interpolate_bounds())
    Call UnitTesting.RunTest("vecmath_point_inside", test_point_inside())
    Call UnitTesting.RunTest("vecmath_point_outside", test_point_outside())
    Call UnitTesting.RunTest("vecmath_prop_sq_vs_len", test_prop_sq_vs_len())
    Call UnitTesting.RunTest("vecmath_prop_add_sub_roundtrip", test_prop_add_sub_roundtrip())
End Sub

' Verifies VecLength returns 5 for a (3,4) vector (classic 3-4-5 triangle).
' Validates: Requirements 1.1
Private Function test_vec_length_3_4() As Boolean
    On Error GoTo Fail
    Dim v As Vector2
    v.x = 3: v.y = 4
    ' sqrt(9 + 16) = sqrt(25) = 5
    test_vec_length_3_4 = (Abs(VecLength(v) - 5!) < 0.001)
    Exit Function
Fail:
    test_vec_length_3_4 = False
End Function

' Verifies VecLength returns 0 for a zero vector.
' Validates: Requirements 1.1
Private Function test_vec_length_zero() As Boolean
    On Error GoTo Fail
    Dim v As Vector2
    v.x = 0: v.y = 0
    test_vec_length_zero = (Abs(VecLength(v)) < 0.001)
    Exit Function
Fail:
    test_vec_length_zero = False
End Function

' Verifies VecSqLength returns 25 for a (3,4) vector.
' Validates: Requirements 1.2
Private Function test_vec_sq_length() As Boolean
    On Error GoTo Fail
    Dim v As Vector2
    v.x = 3: v.y = 4
    ' 3^2 + 4^2 = 9 + 16 = 25
    test_vec_sq_length = (Abs(VecSqLength(v) - 25!) < 0.001)
    Exit Function
Fail:
    test_vec_sq_length = False
End Function

' Verifies Normalize on (3,4) produces approximately (0.6, 0.8).
' Validates: Requirements 1.3
Private Function test_vec_normalize() As Boolean
    On Error GoTo Fail
    Dim v As Vector2
    v.x = 3: v.y = 4
    Call Normalize(v)
    ' 3/5 = 0.6, 4/5 = 0.8
    test_vec_normalize = (Abs(v.x - 0.6!) < 0.001 And Abs(v.y - 0.8!) < 0.001)
    Exit Function
Fail:
    test_vec_normalize = False
End Function

' Verifies VAdd: (1,2) + (3,4) = (4,6).
' Validates: Requirements 1.4
Private Function test_vec_add() As Boolean
    On Error GoTo Fail
    Dim a As Vector2, b As Vector2, r As Vector2
    a.x = 1: a.y = 2
    b.x = 3: b.y = 4
    r = VAdd(a, b)
    test_vec_add = (Abs(r.x - 4!) < 0.001 And Abs(r.y - 6!) < 0.001)
    Exit Function
Fail:
    test_vec_add = False
End Function

' Verifies VSubs: (5,7) - (3,4) = (2,3).
' Validates: Requirements 1.5
Private Function test_vec_subs() As Boolean
    On Error GoTo Fail
    Dim a As Vector2, b As Vector2, r As Vector2
    a.x = 5: a.y = 7
    b.x = 3: b.y = 4
    r = VSubs(a, b)
    test_vec_subs = (Abs(r.x - 2!) < 0.001 And Abs(r.y - 3!) < 0.001)
    Exit Function
Fail:
    test_vec_subs = False
End Function

' Verifies VMul: (2,3) * 4 = (8,12).
' Validates: Requirements 1.6
Private Function test_vec_mul() As Boolean
    On Error GoTo Fail
    Dim v As Vector2, r As Vector2
    v.x = 2: v.y = 3
    r = VMul(v, 4!)
    test_vec_mul = (Abs(r.x - 8!) < 0.001 And Abs(r.y - 12!) < 0.001)
    Exit Function
Fail:
    test_vec_mul = False
End Function

' Verifies GetAngle returns 0 when both points are the same.
' Validates: Requirements 1.7
Private Function test_getangle_same() As Boolean
    On Error GoTo Fail
    ' Same point should return 0 (early exit in GetAngle)
    Dim result As Double
    result = GetAngle(5#, 5#, 5#, 5#)
    test_getangle_same = (Abs(result) < 0.001)
    Exit Function
Fail:
    test_getangle_same = False
End Function

' Verifies GetAngle returns 0 for a point along the positive X axis.
' Validates: Requirements 1.7
Private Function test_getangle_pos_x() As Boolean
    On Error GoTo Fail
    ' Point to the right on the X axis should return 0
    Dim result As Double
    result = GetAngle(0#, 0#, 10#, 0#)
    test_getangle_pos_x = (Abs(result) < 0.001)
    Exit Function
Fail:
    test_getangle_pos_x = False
End Function

' Verifies GetAngle returns approximately PI for a point along the negative X axis.
' Validates: Requirements 1.7
Private Function test_getangle_neg_x() As Boolean
    On Error GoTo Fail
    ' Point to the left on the X axis should return PI
    Dim result As Double
    result = GetAngle(0#, 0#, -10#, 0#)
    test_getangle_neg_x = (Abs(result - 3.14159265358979) < 0.001)
    Exit Function
Fail:
    test_getangle_neg_x = False
End Function

' Verifies FixAngle wraps -45 to 315.
' Validates: Requirements 1.8
Private Function test_fixangle_neg() As Boolean
    On Error GoTo Fail
    test_fixangle_neg = (FixAngle(-45) = 315)
    Exit Function
Fail:
    test_fixangle_neg = False
End Function

' Verifies FixAngle wraps 400 to 40.
' Validates: Requirements 1.8
Private Function test_fixangle_overflow() As Boolean
    On Error GoTo Fail
    test_fixangle_overflow = (FixAngle(400) = 40)
    Exit Function
Fail:
    test_fixangle_overflow = False
End Function

' Verifies Interpolate(0, 100, 0.5) returns 50.
' Validates: Requirements 1.9
Private Function test_interpolate_mid() As Boolean
    On Error GoTo Fail
    test_interpolate_mid = (Interpolate(0, 100, 0.5) = 50)
    Exit Function
Fail:
    test_interpolate_mid = False
End Function

' Verifies Interpolate boundary values: t=0 returns A, t=1 returns B.
' Validates: Requirements 1.9
Private Function test_interpolate_bounds() As Boolean
    On Error GoTo Fail
    Dim atZero As Integer
    Dim atOne As Integer
    atZero = Interpolate(0, 100, 0#)
    atOne = Interpolate(0, 100, 1#)
    test_interpolate_bounds = (atZero = 0 And atOne = 100)
    Exit Function
Fail:
    test_interpolate_bounds = False
End Function

' Verifies PointIsInsideRect returns True for a point inside the rectangle.
' Validates: Requirements 1.10
Private Function test_point_inside() As Boolean
    On Error GoTo Fail
    Dim r As RECT
    r.Left = 10: r.Top = 10: r.Right = 50: r.Bottom = 50
    ' (25,25) is clearly inside the rectangle
    test_point_inside = (PointIsInsideRect(25, 25, r) = True)
    Exit Function
Fail:
    test_point_inside = False
End Function

' Verifies PointIsInsideRect returns False for a point outside the rectangle.
' Validates: Requirements 1.10
Private Function test_point_outside() As Boolean
    On Error GoTo Fail
    Dim r As RECT
    r.Left = 10: r.Top = 10: r.Right = 50: r.Bottom = 50
    ' (5,5) is outside the rectangle (above and to the left)
    test_point_outside = (PointIsInsideRect(5, 5, r) = False)
    Exit Function
Fail:
    test_point_outside = False
End Function

' Feature: full-coverage-unit-tests, Property 1: Squared length equals length squared
' **Validates: Requirements 1.11**
Private Function test_prop_sq_vs_len() As Boolean
    On Error GoTo Fail
    Dim iterations As Long
    Dim i As Long
    Dim v As Vector2
    Dim sqLen As Single
    Dim vecLen As Single
    iterations = 0
    For i = 1 To 120
        ' Generate deterministic values from i
        v.x = CSng(i * 1.5 - 90)
        v.y = CSng(i * 0.7 - 42)
        sqLen = VecSqLength(v)
        vecLen = VecLength(v)
        If Abs(sqLen - vecLen * vecLen) > 0.01 Then
            test_prop_sq_vs_len = False
            Exit Function
        End If
        iterations = iterations + 1
    Next i
    test_prop_sq_vs_len = (iterations >= 100)
    Exit Function
Fail:
    test_prop_sq_vs_len = False
End Function

' Feature: full-coverage-unit-tests, Property 2: Vector addition/subtraction round-trip
' **Validates: Requirements 1.12**
Private Function test_prop_add_sub_roundtrip() As Boolean
    On Error GoTo Fail
    Dim iterations As Long
    Dim i As Long
    Dim a As Vector2
    Dim b As Vector2
    Dim roundtrip As Vector2
    iterations = 0
    For i = 1 To 120
        ' Generate two Vector2 pairs deterministically from loop index
        a.x = CSng(i * 1.3 - 78)
        a.y = CSng(i * 0.9 - 54)
        b.x = CSng(i * 2.1 - 126)
        b.y = CSng(i * 0.4 - 24)
        ' Round-trip: VSubs(VAdd(a, b), a) should equal b
        roundtrip = VSubs(VAdd(a, b), a)
        If Abs(roundtrip.x - b.x) > 0.01 Then
            test_prop_add_sub_roundtrip = False
            Exit Function
        End If
        If Abs(roundtrip.y - b.y) > 0.01 Then
            test_prop_add_sub_roundtrip = False
            Exit Function
        End If
        iterations = iterations + 1
    Next i
    test_prop_add_sub_roundtrip = (iterations >= 100)
    Exit Function
Fail:
    test_prop_add_sub_roundtrip = False
End Function

#End If
