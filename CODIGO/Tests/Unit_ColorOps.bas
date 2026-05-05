Attribute VB_Name = "Unit_ColorOps"
Option Explicit

' ==========================================================================
' ColorOps Test Suite
' Tests Graficos_Color.bas color manipulation functions: RGBA_From_Comp,
' SetRGBA, LerpRGBA, ModulateRGBA, AddRGBA, and RGBA_ToString with
' example-based tests covering construction, interpolation, modulation,
' clamped addition, and string formatting.
' ==========================================================================

#If UNIT_TEST = 1 Then

Public Sub test_suite_color_ops()
    Call UnitTesting.RunTest("colorops_from_comp", test_from_comp())
    Call UnitTesting.RunTest("colorops_set_rgba", test_set_rgba())
    Call UnitTesting.RunTest("colorops_lerp_mid", test_lerp_mid())
    Call UnitTesting.RunTest("colorops_lerp_bounds", test_lerp_bounds())
    Call UnitTesting.RunTest("colorops_modulate", test_modulate())
    Call UnitTesting.RunTest("colorops_add_clamp", test_add_clamp())
    Call UnitTesting.RunTest("colorops_to_string", test_to_string())
    Call UnitTesting.RunTest("colorops_prop_long_roundtrip", test_prop_long_roundtrip())
End Sub

' Verifies RGBA_From_Comp stores each component correctly.
' Validates: Requirements 2.1
Private Function test_from_comp() As Boolean
    On Error GoTo Fail
    Dim c As RGBA
    c = RGBA_From_Comp(128, 64, 32, 255)
    test_from_comp = (c.R = 128 And c.G = 64 And c.B = 32 And c.A = 255)
    Exit Function
Fail:
    test_from_comp = False
End Function

' Verifies SetRGBA writes R=10, G=20, B=30, A=40 into an RGBA structure.
' Validates: Requirements 2.2
Private Function test_set_rgba() As Boolean
    On Error GoTo Fail
    Dim c As RGBA
    Call SetRGBA(c, 10, 20, 30, 40)
    test_set_rgba = (c.R = 10 And c.G = 20 And c.B = 30 And c.A = 40)
    Exit Function
Fail:
    test_set_rgba = False
End Function

' Verifies LerpRGBA at factor 0.5 between (0,0,0,0) and (100,200,100,200)
' produces (50,100,50,100).
' Validates: Requirements 2.3
Private Function test_lerp_mid() As Boolean
    On Error GoTo Fail
    Dim a As RGBA, b As RGBA, dest As RGBA
    Call SetRGBA(a, 0, 0, 0, 0)
    Call SetRGBA(b, 100, 200, 100, 200)
    Call LerpRGBA(dest, a, b, 0.5!)
    test_lerp_mid = (dest.R = 50 And dest.G = 100 And dest.B = 50 And dest.A = 100)
    Exit Function
Fail:
    test_lerp_mid = False
End Function

' Verifies LerpRGBA boundary conditions: factor 0 returns color A,
' factor 1 returns color B.
' Validates: Requirements 2.4
Private Function test_lerp_bounds() As Boolean
    On Error GoTo Fail
    Dim a As RGBA, b As RGBA, dest As RGBA
    Call SetRGBA(a, 10, 20, 30, 40)
    Call SetRGBA(b, 200, 210, 220, 230)
    
    ' Factor = 0 should return color A
    Call LerpRGBA(dest, a, b, 0!)
    Dim okA As Boolean
    okA = (dest.R = 10 And dest.G = 20 And dest.B = 30 And dest.A = 40)
    
    ' Factor = 1 should return color B
    Call LerpRGBA(dest, a, b, 1!)
    Dim okB As Boolean
    okB = (dest.R = 200 And dest.G = 210 And dest.B = 220 And dest.A = 230)
    
    test_lerp_bounds = (okA And okB)
    Exit Function
Fail:
    test_lerp_bounds = False
End Function

' Verifies ModulateRGBA: (255,255,255,255) modulated with (128,128,128,128)
' produces approximately (128,128,128,128).
' Implementation uses integer division: CLng(255)*128\255 = 128.
' Validates: Requirements 2.5
Private Function test_modulate() As Boolean
    On Error GoTo Fail
    Dim a As RGBA, b As RGBA, dest As RGBA
    Call SetRGBA(a, 255, 255, 255, 255)
    Call SetRGBA(b, 128, 128, 128, 128)
    Call ModulateRGBA(dest, a, b)
    ' 255*128\255 = 128 for each channel
    test_modulate = (dest.R = 128 And dest.G = 128 And dest.B = 128 And dest.A = 128)
    Exit Function
Fail:
    test_modulate = False
End Function

' Verifies AddRGBA clamps each component to 255 when the sum exceeds it.
' (200,200,200,200) + (100,100,100,100) = (300,300,300,300) clamped to (255,255,255,255).
' Validates: Requirements 2.6
Private Function test_add_clamp() As Boolean
    On Error GoTo Fail
    Dim a As RGBA, b As RGBA, dest As RGBA
    Call SetRGBA(a, 200, 200, 200, 200)
    Call SetRGBA(b, 100, 100, 100, 100)
    Call AddRGBA(dest, a, b)
    ' Each channel: 200+100=300, clamped to 255
    test_add_clamp = (dest.R = 255 And dest.G = 255 And dest.B = 255 And dest.A = 255)
    Exit Function
Fail:
    test_add_clamp = False
End Function

' Verifies RGBA_ToString output format matches "RGBA(R, G, B, A)".
' Validates: Requirements 2.7
Private Function test_to_string() As Boolean
    On Error GoTo Fail
    Dim c As RGBA
    Call SetRGBA(c, 255, 128, 0, 200)
    test_to_string = (RGBA_ToString(c) = "RGBA(255, 128, 0, 200)")
    Exit Function
Fail:
    test_to_string = False
End Function

' Feature: full-coverage-unit-tests, Property 3: RGBA Long round-trip
' For any RGBA color c constructed from byte components (R, G, B, A each
' in 0-255), RGBA_From_Long(RGBA_2_Long(c)) produces identical components.
' Validates: Requirements 2.8
Private Function test_prop_long_roundtrip() As Boolean
    On Error GoTo Fail
    Dim iterations As Long
    Dim i As Long
    Dim R As Byte
    Dim G As Byte
    Dim B As Byte
    Dim A As Byte
    Dim c As RGBA
    Dim lng As Long
    Dim result As RGBA
    iterations = 0
    For i = 1 To 120
        R = CByte(i Mod 256)
        G = CByte((i * 3) Mod 256)
        B = CByte((i * 7) Mod 256)
        A = CByte((i * 11) Mod 256)
        
        c = RGBA_From_Comp(R, G, B, A)
        
        lng = RGBA_2_Long(c)
        
        result = RGBA_From_Long(lng)
        
        If result.R <> R Or result.G <> G Or result.B <> B Or result.A <> A Then
            test_prop_long_roundtrip = False
            Exit Function
        End If
        
        iterations = iterations + 1
    Next i
    test_prop_long_roundtrip = (iterations >= 100)
    Exit Function
Fail:
    test_prop_long_roundtrip = False
End Function

#End If
