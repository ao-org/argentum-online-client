Attribute VB_Name = "Unit_ShadowProjection"
Option Explicit

#If UNIT_TEST = 1 Then

Public Sub test_suite_shadow_projection()
    Call UnitTesting.RunTest("shadow_base_horizontal", test_base_horizontal())
    Call UnitTesting.RunTest("shadow_body_pivot_fallback", test_body_pivot_fallback())
    Call UnitTesting.RunTest("shadow_near_horizontal_depth", test_near_horizontal_depth())
    Call UnitTesting.RunTest("shadow_sun_06_west", test_sun_06_west())
    Call UnitTesting.RunTest("shadow_sun_18_east", test_sun_18_east())
    Call UnitTesting.RunTest("shadow_sun_fractional_change", test_sun_fractional_change())
    Call UnitTesting.RunTest("shadow_local_isometric_direction", test_local_isometric_direction())
End Sub

Private Function test_base_horizontal() As Boolean
    Dim west As ShadowProjectionData
    Dim east As ShadowProjectionData
    west = Shadow_Project(10, 20, 30, 24, -48, 16)
    east = Shadow_Project(10, 20, 30, 24, 48, 16)
    test_base_horizontal = (west.BaseY = east.BaseY And west.BaseRightX - west.BaseLeftX = east.BaseRightX - east.BaseLeftX)
End Function

Private Function test_body_pivot_fallback() As Boolean
    test_body_pivot_fallback = (Shadow_BodyPivot(0, 48) = 47 And Shadow_BodyPivot(30, 48) = 30)
End Function

Private Function test_near_horizontal_depth() As Boolean
    Dim projection As ShadowProjectionData
    projection = Shadow_Project(0, 32, 40, 39, 20, 0.01)
    test_near_horizontal_depth = (projection.TopY < projection.BaseY)
End Function

Private Function test_sun_06_west() As Boolean
    test_sun_06_west = (Shadow_SunDirectionX(6) < 0)
End Function

Private Function test_sun_18_east() As Boolean
    test_sun_18_east = (Shadow_SunDirectionX(18) > 0)
End Function

Private Function test_sun_fractional_change() As Boolean
    test_sun_fractional_change = (Shadow_SunDirectionX(14.9) > Shadow_SunDirectionX(10.9))
End Function

Private Function test_local_isometric_direction() As Boolean
    test_local_isometric_direction = (Shadow_LocalDirectionX(10, 10, 11, 10, 3) > 0 And _
            Shadow_LocalDirectionX(10, 10, 10, 11, 3) < 0 And Shadow_LocalDepth(10, 10, 10, 11, 3) > 0)
End Function

#End If