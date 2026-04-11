Attribute VB_Name = "Unit_MD5"
Option Explicit

' ==========================================================================
' MD5 Helpers Test Suite
' Tests crypto utility functions: hex-to-decimal conversion, hex MD5 string
' to ASCII conversion, and character offset cipher (identity and round-trip).
' ==========================================================================

#If UNIT_TEST = 1 Then

' Runs all MD5 helper unit tests.
Public Sub test_suite_md5()
    Call UnitTesting.RunTest("md5_hex2dec_ff", test_hex2dec_ff())
    Call UnitTesting.RunTest("md5_hex2dec_zero", test_hex2dec_zero())
    Call UnitTesting.RunTest("md5_to_ascii", test_md5_to_ascii())
    Call UnitTesting.RunTest("md5_offset_identity", test_offset_identity())
    Call UnitTesting.RunTest("md5_offset_round_trip", test_offset_round_trip())
End Sub

' Verifies hexHex2Dec converts "FF" to 255.
Private Function test_hex2dec_ff() As Boolean
    On Error GoTo Fail
    test_hex2dec_ff = (hexHex2Dec("FF") = 255)
    Exit Function
Fail:
    test_hex2dec_ff = False
End Function

' Verifies hexHex2Dec converts "00" to 0.
Private Function test_hex2dec_zero() As Boolean
    On Error GoTo Fail
    test_hex2dec_zero = (hexHex2Dec("00") = 0)
    Exit Function
Fail:
    test_hex2dec_zero = False
End Function

' Verifies hexMd52Asc converts "4142" to "AB" (0x41=A, 0x42=B).
Private Function test_md5_to_ascii() As Boolean
    On Error GoTo Fail
    test_md5_to_ascii = (hexMd52Asc("4142") = "AB")
    Exit Function
Fail:
    test_md5_to_ascii = False
End Function

' Verifies txtOffset with offset 0 returns the original string unchanged.
Private Function test_offset_identity() As Boolean
    On Error GoTo Fail
    test_offset_identity = (txtOffset("Hello", 0) = "Hello")
    Exit Function
Fail:
    test_offset_identity = False
End Function

' Verifies txtOffset with a positive offset followed by the negated
' offset returns the original string (round-trip for printable ASCII).
Private Function test_offset_round_trip() As Boolean
    On Error GoTo Fail
    Dim shifted As String
    shifted = txtOffset("ABC", 3)
    Dim restored As String
    restored = txtOffset(shifted, -3)
    test_offset_round_trip = (restored = "ABC")
    Exit Function
Fail:
    test_offset_round_trip = False
End Function

#End If
