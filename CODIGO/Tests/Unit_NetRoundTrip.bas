Attribute VB_Name = "Unit_NetRoundTrip"
Option Explicit

' ==========================================================================
' Network Serialization Round-Trip Test Suite
' Tests the DirectPlay buffer functions used by clsNetWriter/clsNetReader:
' AddDataToBuffer/GetDataFromBuffer for Int8, Int16, Int32, Real32, Bool,
' and AddStringToBuffer/GetStringFromBuffer for String8.
' Verifies that write-then-read produces the original value and that
' sequential multi-type writes preserve order.
'
' Requires both UNIT_TEST and DIRECT_PLAY conditional compilation flags
' since the buffer functions and SIZE_* constants come from DxVBLibA.
' ==========================================================================

#If UNIT_TEST = 1 Then
#If DIRECT_PLAY = 1 Then

Public Sub test_suite_net_round_trip()
    Call UnitTesting.RunTest("net_sequential_multi_type", test_net_sequential_multi_type())
    Call UnitTesting.RunTest("net_pbt_int16_round_trip", test_net_pbt_int16_round_trip())
    Call UnitTesting.RunTest("net_pbt_int32_round_trip", test_net_pbt_int32_round_trip())
    Call UnitTesting.RunTest("net_pbt_int8_round_trip", test_net_pbt_int8_round_trip())
    Call UnitTesting.RunTest("net_pbt_real32_round_trip", test_net_pbt_real32_round_trip())
    Call UnitTesting.RunTest("net_pbt_bool_round_trip", test_net_pbt_bool_round_trip())
    Call UnitTesting.RunTest("net_pbt_string8_round_trip", test_net_pbt_string8_round_trip())
    Call UnitTesting.RunTest("net_pbt_sequential_ordering", test_net_pbt_sequential_ordering())
End Sub

' Requirement 2.7: Sequential multi-type write/read preserving order
' Also validates 2.1 (Int16), 2.2 (Int32), 2.3 (Int8), 2.4 (Real32),
' 2.5 (Bool), 2.6 (String8) via the sequential round-trip
Private Function test_net_sequential_multi_type() As Boolean
    On Error GoTo Fail

    Dim buf() As Byte
    Dim wOffset As Long
    Dim rOffset As Long

    ' --- Write phase: write multiple typed values sequentially ---
    wOffset = NewBuffer(buf)

    Dim wInt8 As Byte:      wInt8 = 200
    Dim wInt16 As Integer:  wInt16 = -12345
    Dim wInt32 As Long:     wInt32 = 1234567890
    Dim wReal32 As Single:  wReal32 = 3.14!
    Dim wBool As Boolean:   wBool = True
    Dim wStr As String:     wStr = "Hello"
    Dim wInt8b As Byte:     wInt8b = 0
    Dim wBoolF As Boolean:  wBoolF = False
    Dim wInt16b As Integer: wInt16b = 32767
    Dim wInt32b As Long:    wInt32b = -2147483647

    Call AddDataToBuffer(buf, wInt8, SIZE_BYTE, wOffset)
    Call AddDataToBuffer(buf, wInt16, SIZE_INTEGER, wOffset)
    Call AddDataToBuffer(buf, wInt32, SIZE_LONG, wOffset)
    Call AddDataToBuffer(buf, wReal32, SIZE_SINGLE, wOffset)
    Call AddDataToBuffer(buf, wBool, SIZE_BOOLEAN, wOffset)
    Call AddStringToBuffer(buf, wStr, wOffset)
    Call AddDataToBuffer(buf, wInt8b, SIZE_BYTE, wOffset)
    Call AddDataToBuffer(buf, wBoolF, SIZE_BOOLEAN, wOffset)
    Call AddDataToBuffer(buf, wInt16b, SIZE_INTEGER, wOffset)
    Call AddDataToBuffer(buf, wInt32b, SIZE_LONG, wOffset)

    ' --- Read phase: read back in the same order ---
    rOffset = 0

    Dim rInt8 As Byte
    Dim rInt16 As Integer
    Dim rInt32 As Long
    Dim rReal32 As Single
    Dim rBool As Boolean
    Dim rStr As String
    Dim rInt8b As Byte
    Dim rBoolF As Boolean
    Dim rInt16b As Integer
    Dim rInt32b As Long

    Call GetDataFromBuffer(buf, rInt8, SIZE_BYTE, rOffset)
    Call GetDataFromBuffer(buf, rInt16, SIZE_INTEGER, rOffset)
    Call GetDataFromBuffer(buf, rInt32, SIZE_LONG, rOffset)
    Call GetDataFromBuffer(buf, rReal32, SIZE_SINGLE, rOffset)
    Call GetDataFromBuffer(buf, rBool, SIZE_BOOLEAN, rOffset)
    rStr = GetStringFromBuffer(buf, rOffset)
    Call GetDataFromBuffer(buf, rInt8b, SIZE_BYTE, rOffset)
    Call GetDataFromBuffer(buf, rBoolF, SIZE_BOOLEAN, rOffset)
    Call GetDataFromBuffer(buf, rInt16b, SIZE_INTEGER, rOffset)
    Call GetDataFromBuffer(buf, rInt32b, SIZE_LONG, rOffset)

    ' --- Verify each value matches and order is preserved ---
    If rInt8 <> wInt8 Then GoTo Fail
    If rInt16 <> wInt16 Then GoTo Fail
    If rInt32 <> wInt32 Then GoTo Fail
    If Abs(rReal32 - wReal32) > 0.001 Then GoTo Fail
    If rBool <> wBool Then GoTo Fail
    If rStr <> wStr Then GoTo Fail
    If rInt8b <> wInt8b Then GoTo Fail
    If rBoolF <> wBoolF Then GoTo Fail
    If rInt16b <> wInt16b Then GoTo Fail
    If rInt32b <> wInt32b Then GoTo Fail

    test_net_sequential_multi_type = True
    Exit Function
Fail:
    test_net_sequential_multi_type = False
End Function

' Feature: unit-test-coverage, Property 1: Network serialization typed round-trip (Int16)
' **Validates: Requirements 2.1**
Private Function test_net_pbt_int16_round_trip() As Boolean
    On Error GoTo Fail

    Dim buf() As Byte
    Dim wOffset As Long
    Dim rOffset As Long
    Dim i As Long
    Dim testVal As Integer
    Dim readVal As Integer

    ' Loop over 100+ values across the Integer range (-32768 to 32767)
    ' Step 655 yields ~100 iterations covering the full range
    For i = -32768 To 32767 Step 655
        testVal = CInt(i)

        ' Fresh buffer for each value
        wOffset = NewBuffer(buf)

        ' Write Int16
        Call AddDataToBuffer(buf, testVal, SIZE_INTEGER, wOffset)

        ' Read Int16 back
        rOffset = 0
        Call GetDataFromBuffer(buf, readVal, SIZE_INTEGER, rOffset)

        ' Verify round-trip equality
        If readVal <> testVal Then
            test_net_pbt_int16_round_trip = False
            Exit Function
        End If
    Next i

    test_net_pbt_int16_round_trip = True
    Exit Function
Fail:
    test_net_pbt_int16_round_trip = False
End Function

' Feature: unit-test-coverage, Property 1: Network serialization typed round-trip (Int32)
' **Validates: Requirements 2.2**
Private Function test_net_pbt_int32_round_trip() As Boolean
    On Error GoTo Fail

    Dim buf() As Byte
    Dim wOffset As Long
    Dim rOffset As Long
    Dim i As Long
    Dim testVal As Long
    Dim readVal As Long

    ' Loop over 100+ values across the Long range (-2147483648 to 2147483647)
    ' Step 42949672 yields ~100 iterations covering the full range
    For i = -2147483647 To 2147483647 - 42949672 Step 42949672
        testVal = i

        ' Fresh buffer for each value
        wOffset = NewBuffer(buf)

        ' Write Int32
        Call AddDataToBuffer(buf, testVal, SIZE_LONG, wOffset)

        ' Read Int32 back
        rOffset = 0
        Call GetDataFromBuffer(buf, readVal, SIZE_LONG, rOffset)

        ' Verify round-trip equality
        If readVal <> testVal Then
            test_net_pbt_int32_round_trip = False
            Exit Function
        End If
    Next i

    ' Also test boundary values explicitly
    Dim boundaries(3) As Long
    boundaries(0) = -2147483648#
    boundaries(1) = -2147483647
    boundaries(2) = 2147483647
    boundaries(3) = 0

    Dim b As Long
    For b = 0 To 3
        testVal = boundaries(b)

        wOffset = NewBuffer(buf)
        Call AddDataToBuffer(buf, testVal, SIZE_LONG, wOffset)

        rOffset = 0
        Call GetDataFromBuffer(buf, readVal, SIZE_LONG, rOffset)

        If readVal <> testVal Then
            test_net_pbt_int32_round_trip = False
            Exit Function
        End If
    Next b

    test_net_pbt_int32_round_trip = True
    Exit Function
Fail:
    test_net_pbt_int32_round_trip = False
End Function

' Feature: unit-test-coverage, Property 1: Network serialization typed round-trip (Int8)
' **Validates: Requirements 2.3**
Private Function test_net_pbt_int8_round_trip() As Boolean
    On Error GoTo Fail

    Dim buf() As Byte
    Dim wOffset As Long
    Dim rOffset As Long
    Dim i As Long
    Dim testVal As Byte
    Dim readVal As Byte

    ' Loop over all 256 Byte values (0 to 255)
    For i = 0 To 255
        testVal = CByte(i)

        ' Fresh buffer for each value
        wOffset = NewBuffer(buf)

        ' Write Int8
        Call AddDataToBuffer(buf, testVal, SIZE_BYTE, wOffset)

        ' Read Int8 back
        rOffset = 0
        Call GetDataFromBuffer(buf, readVal, SIZE_BYTE, rOffset)

        ' Verify round-trip equality
        If readVal <> testVal Then
            test_net_pbt_int8_round_trip = False
            Exit Function
        End If
    Next i

    test_net_pbt_int8_round_trip = True
    Exit Function
Fail:
    test_net_pbt_int8_round_trip = False
End Function

' Feature: unit-test-coverage, Property 1: Network serialization typed round-trip (Real32)
' **Validates: Requirements 2.4**
Private Function test_net_pbt_real32_round_trip() As Boolean
    On Error GoTo Fail

    Dim buf() As Byte
    Dim wOffset As Long
    Dim rOffset As Long
    Dim i As Long
    Dim testVal As Single
    Dim readVal As Single

    ' Loop over 100+ values from -500 to 500 step 10 (101 iterations)
    For i = -500 To 500 Step 10
        testVal = CSng(i)

        ' Fresh buffer for each value
        wOffset = NewBuffer(buf)

        ' Write Real32
        Call AddDataToBuffer(buf, testVal, SIZE_SINGLE, wOffset)

        ' Read Real32 back
        rOffset = 0
        Call GetDataFromBuffer(buf, readVal, SIZE_SINGLE, rOffset)

        ' Verify round-trip within floating-point tolerance
        If Abs(readVal - testVal) > 0.001 Then
            test_net_pbt_real32_round_trip = False
            Exit Function
        End If
    Next i

    ' Also test special floating-point values
    Dim specials(4) As Single
    specials(0) = 0!
    specials(1) = -0.5!
    specials(2) = 3.14!
    specials(3) = -9999.99!
    specials(4) = 0.001!

    Dim s As Long
    For s = 0 To 4
        testVal = specials(s)

        wOffset = NewBuffer(buf)
        Call AddDataToBuffer(buf, testVal, SIZE_SINGLE, wOffset)

        rOffset = 0
        Call GetDataFromBuffer(buf, readVal, SIZE_SINGLE, rOffset)

        If Abs(readVal - testVal) > 0.001 Then
            test_net_pbt_real32_round_trip = False
            Exit Function
        End If
    Next s

    test_net_pbt_real32_round_trip = True
    Exit Function
Fail:
    test_net_pbt_real32_round_trip = False
End Function

' Feature: unit-test-coverage, Property 1: Network serialization typed round-trip (Bool)
' **Validates: Requirements 2.5**
Private Function test_net_pbt_bool_round_trip() As Boolean
    On Error GoTo Fail

    Dim buf() As Byte
    Dim wOffset As Long
    Dim rOffset As Long
    Dim testVal As Boolean
    Dim readVal As Boolean

    ' Test True
    testVal = True
    wOffset = NewBuffer(buf)
    Call AddDataToBuffer(buf, testVal, SIZE_BOOLEAN, wOffset)

    rOffset = 0
    Call GetDataFromBuffer(buf, readVal, SIZE_BOOLEAN, rOffset)

    If readVal <> testVal Then
        test_net_pbt_bool_round_trip = False
        Exit Function
    End If

    ' Test False
    testVal = False
    wOffset = NewBuffer(buf)
    Call AddDataToBuffer(buf, testVal, SIZE_BOOLEAN, wOffset)

    rOffset = 0
    Call GetDataFromBuffer(buf, readVal, SIZE_BOOLEAN, rOffset)

    If readVal <> testVal Then
        test_net_pbt_bool_round_trip = False
        Exit Function
    End If

    test_net_pbt_bool_round_trip = True
    Exit Function
Fail:
    test_net_pbt_bool_round_trip = False
End Function

' Feature: unit-test-coverage, Property 1: Network serialization typed round-trip (String8)
' **Validates: Requirements 2.6**
Private Function test_net_pbt_string8_round_trip() As Boolean
    On Error GoTo Fail

    Dim buf() As Byte
    Dim wOffset As Long
    Dim rOffset As Long
    Dim i As Long
    Dim testStr As String
    Dim readStr As String

    ' Loop over 100+ strings of varying lengths (1 to 120 chars)
    For i = 1 To 120
        ' Build a string of length i using repeating printable ASCII chars
        testStr = String$(i, Chr$(65 + (i Mod 26)))  ' A-Z cycling

        ' Fresh buffer for each string
        wOffset = NewBuffer(buf)

        ' Write String8
        Call AddStringToBuffer(buf, testStr, wOffset)

        ' Read String8 back
        rOffset = 0
        readStr = GetStringFromBuffer(buf, rOffset)

        ' Verify round-trip equality
        If readStr <> testStr Then
            test_net_pbt_string8_round_trip = False
            Exit Function
        End If
    Next i

    test_net_pbt_string8_round_trip = True
    Exit Function
Fail:
    test_net_pbt_string8_round_trip = False
End Function

' Feature: unit-test-coverage, Property 2: Network serialization sequential ordering
' **Validates: Requirements 2.7**
Private Function test_net_pbt_sequential_ordering() As Boolean
    On Error GoTo Fail

    Dim buf() As Byte
    Dim wOffset As Long
    Dim rOffset As Long
    Dim i As Long

    ' Write values for each iteration
    Dim wByte As Byte
    Dim wInt As Integer
    Dim wLng As Long
    Dim wStr As String

    ' Read values for verification
    Dim rByte As Byte
    Dim rInt As Integer
    Dim rLng As Long
    Dim rStr As String

    ' Loop over 110 sequences (i = 0 to 109), each writing a mixed-type
    ' sequence of Byte, Integer, Long, String then reading back in order
    For i = 0 To 109
        ' Compute test values for this iteration
        wByte = CByte(i Mod 256)
        wInt = CInt(i * 100 - 5000)
        wLng = CLng(i * 10000)
        wStr = "test_" & CStr(i)

        ' Fresh buffer for each iteration
        wOffset = NewBuffer(buf)

        ' Write phase: Byte, Integer, Long, String sequentially
        Call AddDataToBuffer(buf, wByte, SIZE_BYTE, wOffset)
        Call AddDataToBuffer(buf, wInt, SIZE_INTEGER, wOffset)
        Call AddDataToBuffer(buf, wLng, SIZE_LONG, wOffset)
        Call AddStringToBuffer(buf, wStr, wOffset)

        ' Read phase: read back in the same order
        rOffset = 0
        Call GetDataFromBuffer(buf, rByte, SIZE_BYTE, rOffset)
        Call GetDataFromBuffer(buf, rInt, SIZE_INTEGER, rOffset)
        Call GetDataFromBuffer(buf, rLng, SIZE_LONG, rOffset)
        rStr = GetStringFromBuffer(buf, rOffset)

        ' Verify each value and position matches
        If rByte <> wByte Then
            test_net_pbt_sequential_ordering = False
            Exit Function
        End If

        If rInt <> wInt Then
            test_net_pbt_sequential_ordering = False
            Exit Function
        End If

        If rLng <> wLng Then
            test_net_pbt_sequential_ordering = False
            Exit Function
        End If

        If rStr <> wStr Then
            test_net_pbt_sequential_ordering = False
            Exit Function
        End If
    Next i

    test_net_pbt_sequential_ordering = True
    Exit Function
Fail:
    test_net_pbt_sequential_ordering = False
End Function

#End If
#End If
