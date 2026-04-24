Attribute VB_Name = "Unit_Color"
Option Explicit

' ==========================================================================
' Color Test Suite
' Tests the RGBA color system: Long<->RGBA round-trip conversions,
' component construction (with default alpha), VB BGR<->D3D RGB swapping,
' linear interpolation (RGBA and RGB-only), channel modulation,
' additive blending with clamping, 4-element array helpers, and
' string formatting.
' ==========================================================================

#If UNIT_TEST = 1 Then

' Runs all color-related unit tests.
Public Function test_suite_color() As Boolean
    Call UnitTesting.RunTest("color_round_trip", test_rgba_round_trip())
    Call UnitTesting.RunTest("color_from_comp", test_rgba_from_comp())
    Call UnitTesting.RunTest("color_from_comp_default_alpha", test_rgba_from_comp_default_alpha())
    Call UnitTesting.RunTest("color_from_long", test_rgba_from_long())
    Call UnitTesting.RunTest("color_from_vbcolor", test_rgba_from_vbcolor())
    Call UnitTesting.RunTest("color_set_rgba", test_set_rgba())
    Call UnitTesting.RunTest("color_lerp_bounds", test_lerp_bounds())
    Call UnitTesting.RunTest("color_lerp_midpoint", test_lerp_midpoint())
    Call UnitTesting.RunTest("color_lerp_rgb_preserves_alpha", test_lerp_rgb_preserves_alpha())
    Call UnitTesting.RunTest("color_modulate", test_modulate())
    Call UnitTesting.RunTest("color_add_clamp", test_add_clamp())
    Call UnitTesting.RunTest("color_long_2_rgba_list", test_long_2_rgba_list())
    Call UnitTesting.RunTest("color_rgba_list", test_rgba_list())
    Call UnitTesting.RunTest("color_rgba_to_list", test_rgba_to_list())
    Call UnitTesting.RunTest("color_copy_rgba_list", test_copy_rgba_list())
    Call UnitTesting.RunTest("color_copy_rgba_list_with_alpha", test_copy_rgba_list_with_alpha())
    Call UnitTesting.RunTest("color_vbcolor_2_long", test_vbcolor_2_long())
    Call UnitTesting.RunTest("color_to_string", test_rgba_to_string())
    test_suite_color = True
End Function

' Verifies that converting an RGBA color to a Long and back
' preserves all four channel values (R, G, B, A) exactly.
Private Function test_rgba_round_trip() As Boolean
    On Error GoTo Fail
    ' Create a color with known channel values
    Dim original As RGBA
    original.R = 100: original.G = 150: original.B = 200: original.A = 255
    ' Pack it into a single Long
    Dim lng As Long
    lng = RGBA_2_Long(original)
    ' Unpack it back into an RGBA struct
    Dim restored As RGBA
    Call Long_2_RGBA(restored, lng)
    ' All four channels should survive the round trip unchanged
    test_rgba_round_trip = (restored.R = 100 And restored.G = 150 _
                            And restored.B = 200 And restored.A = 255)
    Exit Function
Fail:
    test_rgba_round_trip = False
End Function

' Verifies that RGBA_From_Comp correctly builds an RGBA struct
' from individual R, G, B, A component values.
Private Function test_rgba_from_comp() As Boolean
    On Error GoTo Fail
    ' Build an RGBA from explicit R=10, G=20, B=30, A=40
    Dim c As RGBA
    c = RGBA_From_Comp(10, 20, 30, 40)
    ' Each channel should match the input values
    test_rgba_from_comp = (c.R = 10 And c.G = 20 And c.B = 30 And c.A = 40)
    Exit Function
Fail:
    test_rgba_from_comp = False
End Function

' Verifies LerpRGBA at the two boundary factors:
'   - factor=0 should return the first color (A) unchanged
'   - factor=1 should return the second color (B) unchanged
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

' Verifies ModulateRGBA multiplies two colors channel-by-channel
' (each product divided by 255). For example:
'   R: 255*128/255 = 128,  G: 128*255/255 = 128,
'   B: 0*0/255 = 0,        A: 255*255/255 = 255
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

' Verifies AddRGBA adds two colors channel-by-channel and clamps
' each result to 255. For example:
'   R: 200+100=300 -> 255,  G: 100+200=300 -> 255,
'   B: 0+0=0,               A: 255+10=265  -> 255
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

' Verifies RGBA_From_Comp uses A=255 when the alpha argument is omitted.
Private Function test_rgba_from_comp_default_alpha() As Boolean
    On Error GoTo Fail
    ' Call without the optional alpha argument
    Dim c As RGBA
    c = RGBA_From_Comp(10, 20, 30)
    ' RGB should match, and alpha should default to 255 (fully opaque)
    test_rgba_from_comp_default_alpha = (c.R = 10 And c.G = 20 And c.B = 30 And c.A = 255)
    Exit Function
Fail:
    test_rgba_from_comp_default_alpha = False
End Function

' Verifies RGBA_From_Long converts a Long back to an RGBA struct
' (function-return variant of Long_2_RGBA).
Private Function test_rgba_from_long() As Boolean
    On Error GoTo Fail
    ' Create a known color and pack it to Long
    Dim original As RGBA
    original.R = 50: original.G = 100: original.B = 150: original.A = 200
    Dim lng As Long
    lng = RGBA_2_Long(original)
    ' Use the function-return variant to unpack
    Dim restored As RGBA
    restored = RGBA_From_Long(lng)
    ' Should produce the same channels as the original
    test_rgba_from_long = (restored.R = 50 And restored.G = 100 _
                           And restored.B = 150 And restored.A = 200)
    Exit Function
Fail:
    test_rgba_from_long = False
End Function

' Verifies RGBA_From_vbColor swaps the R and B channels
' (VB6 stores colors as BGR) and forces alpha to 255.
' Input Long has bytes laid out as [B=0, G=0, R=200, A=0] in memory,
' so after the swap we expect R=200 to land in .R and .B=0, .A=255.
Private Function test_rgba_from_vbcolor() As Boolean
    On Error GoTo Fail
    ' Build a known Long: R=30, G=60, B=90 in VB BGR order
    Dim src As RGBA
    src.R = 30: src.G = 60: src.B = 90: src.A = 0
    Dim lng As Long
    lng = RGBA_2_Long(src)
    Dim result As RGBA
    result = RGBA_From_vbColor(lng)
    ' The function swaps R<->B and sets A=255
    test_rgba_from_vbcolor = (result.R = 90 And result.G = 60 _
                              And result.B = 30 And result.A = 255)
    Exit Function
Fail:
    test_rgba_from_vbcolor = False
End Function

' Verifies SetRGBA writes all four channels into an existing struct by reference.
Private Function test_set_rgba() As Boolean
    On Error GoTo Fail
    ' Start with an uninitialized struct
    Dim c As RGBA
    ' Set all four channels via the ByRef helper
    Call SetRGBA(c, 11, 22, 33, 44)
    ' Verify each channel was written correctly
    test_set_rgba = (c.R = 11 And c.G = 22 And c.B = 33 And c.A = 44)
    Exit Function
Fail:
    test_set_rgba = False
End Function

' Verifies LerpRGBA at factor=0.5 returns the midpoint of two colors.
' (100,100,100,100) lerped with (200,200,200,200) at 0.5 = (150,150,150,150)
Private Function test_lerp_midpoint() As Boolean
    On Error GoTo Fail
    ' Set up two colors 100 apart on every channel
    Dim a As RGBA, b As RGBA, dest As RGBA
    a.R = 100: a.G = 100: a.B = 100: a.A = 100
    b.R = 200: b.G = 200: b.B = 200: b.A = 200
    ' Lerp at 0.5 should land exactly in the middle: 150 on each channel
    Call LerpRGBA(dest, a, b, 0.5!)
    test_lerp_midpoint = (dest.R = 150 And dest.G = 150 _
                          And dest.B = 150 And dest.A = 150)
    Exit Function
Fail:
    test_lerp_midpoint = False
End Function

' Verifies LerpRGB only interpolates R, G, B and leaves the
' destination alpha channel untouched.
Private Function test_lerp_rgb_preserves_alpha() As Boolean
    On Error GoTo Fail
    Dim a As RGBA, b As RGBA, dest As RGBA
    a.R = 0: a.G = 0: a.B = 0: a.A = 100
    b.R = 200: b.G = 200: b.B = 200: b.A = 200
    ' Pre-set dest alpha to a known value; LerpRGB should not change it
    dest.A = 42
    Call LerpRGB(dest, a, b, 1!)
    ' RGB should be fully color B, alpha should remain 42
    test_lerp_rgb_preserves_alpha = (dest.R = 200 And dest.G = 200 _
                                     And dest.B = 200 And dest.A = 42)
    Exit Function
Fail:
    test_lerp_rgb_preserves_alpha = False
End Function

' Verifies Long_2_RGBAList fills all 4 elements of an RGBA array
' with the same color decoded from a Long.
Private Function test_long_2_rgba_list() As Boolean
    On Error GoTo Fail
    ' Prepare a 4-element array and a known color packed as Long
    Dim arr(0 To 3) As RGBA
    Dim src As RGBA
    src.R = 10: src.G = 20: src.B = 30: src.A = 40
    Dim lng As Long
    lng = RGBA_2_Long(src)
    ' Fill all 4 slots from the same Long
    Call Long_2_RGBAList(arr, lng)
    ' Every slot should have the same color
    Dim i As Long, ok As Boolean
    ok = True
    For i = 0 To 3
        If arr(i).R <> 10 Or arr(i).G <> 20 Or arr(i).B <> 30 Or arr(i).A <> 40 Then
            ok = False
        End If
    Next
    test_long_2_rgba_list = ok
    Exit Function
Fail:
    test_long_2_rgba_list = False
End Function

' Verifies RGBAList fills all 4 elements of an RGBA array
' from individual R, G, B, A component values.
Private Function test_rgba_list() As Boolean
    On Error GoTo Fail
    ' Prepare a 4-element array
    Dim arr(0 To 3) As RGBA
    ' Fill all 4 slots from individual component values
    Call RGBAList(arr, 5, 10, 15, 20)
    ' Every slot should have R=5, G=10, B=15, A=20
    Dim i As Long, ok As Boolean
    ok = True
    For i = 0 To 3
        If arr(i).R <> 5 Or arr(i).G <> 10 Or arr(i).B <> 15 Or arr(i).A <> 20 Then
            ok = False
        End If
    Next
    test_rgba_list = ok
    Exit Function
Fail:
    test_rgba_list = False
End Function

' Verifies RGBA_ToList copies a single RGBA color into all 4 slots
' of a destination array.
Private Function test_rgba_to_list() As Boolean
    On Error GoTo Fail
    ' Prepare a 4-element array and a source color
    Dim arr(0 To 3) As RGBA
    Dim src As RGBA
    src.R = 77: src.G = 88: src.B = 99: src.A = 111
    ' Broadcast the single color into all 4 array slots
    Call RGBA_ToList(arr, src)
    ' Every slot should match the source color
    Dim i As Long, ok As Boolean
    ok = True
    For i = 0 To 3
        If arr(i).R <> 77 Or arr(i).G <> 88 Or arr(i).B <> 99 Or arr(i).A <> 111 Then
            ok = False
        End If
    Next
    test_rgba_to_list = ok
    Exit Function
Fail:
    test_rgba_to_list = False
End Function

' Verifies Copy_RGBAList copies all 4 elements from one array to another.
Private Function test_copy_rgba_list() As Boolean
    On Error GoTo Fail
    ' Fill source array with 4 different colors (increasing channel values)
    Dim src(0 To 3) As RGBA, dest(0 To 3) As RGBA
    Dim i As Long
    For i = 0 To 3
        Call SetRGBA(src(i), CByte(i * 10), CByte(i * 20), CByte(i * 30), 255)
    Next
    ' Copy the entire source array into dest
    Call Copy_RGBAList(dest, src)
    ' Every slot in dest should match the corresponding source slot
    Dim ok As Boolean
    ok = True
    For i = 0 To 3
        If dest(i).R <> src(i).R Or dest(i).G <> src(i).G _
           Or dest(i).B <> src(i).B Or dest(i).A <> src(i).A Then
            ok = False
        End If
    Next
    test_copy_rgba_list = ok
    Exit Function
Fail:
    test_copy_rgba_list = False
End Function

' Verifies Copy_RGBAList_WithAlpha copies all 4 elements but overrides
' the alpha channel with the given value.
Private Function test_copy_rgba_list_with_alpha() As Boolean
    On Error GoTo Fail
    ' Fill source array with a uniform color (alpha=255)
    Dim src(0 To 3) As RGBA, dest(0 To 3) As RGBA
    Dim i As Long
    For i = 0 To 3
        Call SetRGBA(src(i), 100, 150, 200, 255)
    Next
    ' Copy but override alpha to 128 on every slot
    Call Copy_RGBAList_WithAlpha(dest, src, 128)
    Dim ok As Boolean
    ok = True
    For i = 0 To 3
        ' RGB should match source, but alpha should be overridden to 128
        If dest(i).R <> 100 Or dest(i).G <> 150 Or dest(i).B <> 200 Or dest(i).A <> 128 Then
            ok = False
        End If
    Next
    test_copy_rgba_list_with_alpha = ok
    Exit Function
Fail:
    test_copy_rgba_list_with_alpha = False
End Function

' Verifies vbColor_2_Long converts a VB BGR Long to a D3D RGBA Long
' by swapping R and B channels and setting alpha to 255.
Private Function test_vbcolor_2_long() As Boolean
    On Error GoTo Fail
    ' Build a VB-style color Long with known R/B layout
    Dim vbCol As RGBA
    vbCol.R = 30: vbCol.G = 60: vbCol.B = 90: vbCol.A = 0
    Dim vbLng As Long
    vbLng = RGBA_2_Long(vbCol)
    ' Convert and decode the result
    Dim resultLng As Long
    resultLng = vbColor_2_Long(vbLng)
    Dim result As RGBA
    Call Long_2_RGBA(result, resultLng)
    ' R and B should be swapped, alpha forced to 255
    test_vbcolor_2_long = (result.R = 90 And result.G = 60 _
                           And result.B = 30 And result.A = 255)
    Exit Function
Fail:
    test_vbcolor_2_long = False
End Function

' Verifies RGBA_ToString produces the expected human-readable string
' format "RGBA(R, G, B, A)".
Private Function test_rgba_to_string() As Boolean
    On Error GoTo Fail
    ' Build a color with distinct channel values
    Dim c As RGBA
    c.R = 255: c.G = 128: c.B = 0: c.A = 200
    ' Should produce the formatted string "RGBA(255, 128, 0, 200)"
    test_rgba_to_string = (RGBA_ToString(c) = "RGBA(255, 128, 0, 200)")
    Exit Function
Fail:
    test_rgba_to_string = False
End Function

#End If
