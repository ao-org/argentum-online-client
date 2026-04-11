Attribute VB_Name = "Unit_Color"
Option Explicit

#If UNIT_TEST = 1 Then

Public Sub test_suite_color()
    Call UnitTesting.RunTest("color_round_trip", test_rgba_round_trip())
    Call UnitTesting.RunTest("color_from_comp", test_rgba_from_comp())
    Call UnitTesting.RunTest("color_lerp_bounds", test_lerp_bounds())
    Call UnitTesting.RunTest("color_modulate", test_modulate())
    Call UnitTesting.RunTest("color_add_clamp", test_add_clamp())
End Sub

Private Function test_rgba_round_trip() As Boolean
    On Error GoTo Fail
    Dim original As RGBA
    original.R = 100: original.G = 150: original.B = 200: original.A = 255
    Dim lng As Long
    lng = RGBA_2_Long(original)
    Dim restored As RGBA
    Call Long_2_RGBA(restored, lng)
    test_rgba_round_trip = (restored.R = 100 And restored.G = 150 _
                            And restored.B = 200 And restored.A = 255)
    Exit Function
Fail:
    test_rgba_round_trip = False
End Function

Private Function test_rgba_from_comp() As Boolean
    On Error GoTo Fail
    Dim c As RGBA
    c = RGBA_From_Comp(10, 20, 30, 40)
    test_rgba_from_comp = (c.R = 10 And c.G = 20 And c.B = 30 And c.A = 40)
    Exit Function
Fail:
    test_rgba_from_comp = False
End Function

Private Function test_lerp_bounds() As Boolean
    On Error GoTo Fail
    Dim a As RGBA, b As RGBA, dest As RGBA
    a.R = 100: a.G = 100: a.B = 100: a.A = 255
    b.R = 200: b.G = 200: b.B = 200: b.A = 255
    ' Factor = 0 should return color A
    Call LerpRGBA(dest, a, b, 0!)
    Dim okA As Boolean
    okA = (dest.R = 100 And dest.G = 100 And dest.B = 100 And dest.A = 255)
    
    ' Factor = 1 should return color B
    Call LerpRGBA(dest, a, b, 1!)
    Dim okB As Boolean
    okB = (dest.R = 200 And dest.G = 200 And dest.B = 200 And dest.A = 255)
    
    test_lerp_bounds = (okA And okB)
    Exit Function
Fail:
    test_lerp_bounds = False
End Function

Private Function test_modulate() As Boolean
    On Error GoTo Fail
    Dim a As RGBA, b As RGBA, dest As RGBA
    a.R = 255: a.G = 128: a.B = 0: a.A = 255
    b.R = 128: b.G = 255: b.B = 0: b.A = 255
    Call ModulateRGBA(dest, a, b)
    ' 255*128\255 = 128, 128*255\255 = 128, 0*0\255 = 0, 255*255\255 = 255
    test_modulate = (dest.R = 128 And dest.G = 128 And dest.B = 0 And dest.A = 255)
    Exit Function
Fail:
    test_modulate = False
End Function

Private Function test_add_clamp() As Boolean
    On Error GoTo Fail
    Dim a As RGBA, b As RGBA, dest As RGBA
    a.R = 200: a.G = 100: a.B = 0: a.A = 255
    b.R = 100: b.G = 200: b.B = 0: b.A = 10
    Call AddRGBA(dest, a, b)
    ' 200+100=300 clamped to 255, 100+200=300 clamped to 255, 0+0=0, 255+10=265 clamped to 255
    test_add_clamp = (dest.R = 255 And dest.G = 255 And dest.B = 0 And dest.A = 255)
    Exit Function
Fail:
    test_add_clamp = False
End Function

#End If
