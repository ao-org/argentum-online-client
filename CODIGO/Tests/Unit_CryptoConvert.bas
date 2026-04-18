Attribute VB_Name = "Unit_CryptoConvert"
Option Explicit

' ==========================================================================
' CryptoConvert Test Suite
' Tests AO20CryptoSysWrapper.bas: HiByte, LoByte, MakeInt, Str2ByteArr,
' ByteArr2String, CopyBytes, ByteArrayToHex, ByteArrayToDecimalString.
'
' Requirements: 2.1, 2.2, 2.3, 2.4, 2.5, 2.6, 2.7, 2.8, 2.9
' ==========================================================================

#If UNIT_TEST = 1 Then

Public Sub test_suite_crypto_convert()
    ' Example-based tests
    Call UnitTesting.RunTest("cc_hibyte_256", test_hibyte_256())
    Call UnitTesting.RunTest("cc_hibyte_0", test_hibyte_0())
    Call UnitTesting.RunTest("cc_lobyte_258", test_lobyte_258())
    Call UnitTesting.RunTest("cc_lobyte_255", test_lobyte_255())
    Call UnitTesting.RunTest("cc_str2bytearr", test_str2bytearr())
    Call UnitTesting.RunTest("cc_bytearr2string", test_bytearr2string())
    Call UnitTesting.RunTest("cc_copybytes", test_copybytes())
    Call UnitTesting.RunTest("cc_bytearraytohex", test_bytearraytohex())
    Call UnitTesting.RunTest("cc_bytearraytodecimalstring", test_bytearraytodecimalstring())
    
    ' Property-based tests
    Call UnitTesting.RunTest("cc_pbt_makeint_roundtrip", test_pbt_makeint_roundtrip())
    Call UnitTesting.RunTest("cc_pbt_str2bytearr_roundtrip", test_pbt_str2bytearr_roundtrip())
    Call UnitTesting.RunTest("cc_pbt_copybytes_correctness", test_pbt_copybytes_correctness())
    Call UnitTesting.RunTest("cc_pbt_byte_to_string_repr", test_pbt_byte_to_string_repr())
End Sub

' --------------------------------------------------------------------------
' Example-based tests
' --------------------------------------------------------------------------

' Requirement 2.1: HiByte(256) = 1
Private Function test_hibyte_256() As Boolean
    On Error GoTo Fail
    test_hibyte_256 = (AO20CryptoSysWrapper.hiByte(256) = 1)
    Exit Function
Fail:
    test_hibyte_256 = False
End Function

' Requirement 2.1: HiByte(0) = 0
Private Function test_hibyte_0() As Boolean
    On Error GoTo Fail
    test_hibyte_0 = (AO20CryptoSysWrapper.hiByte(0) = 0)
    Exit Function
Fail:
    test_hibyte_0 = False
End Function

' Requirement 2.2: LoByte(258) = 2
Private Function test_lobyte_258() As Boolean
    On Error GoTo Fail
    test_lobyte_258 = (AO20CryptoSysWrapper.LoByte(258) = 2)
    Exit Function
Fail:
    test_lobyte_258 = False
End Function

' Requirement 2.2: LoByte(255) = 255
Private Function test_lobyte_255() As Boolean
    On Error GoTo Fail
    test_lobyte_255 = (AO20CryptoSysWrapper.LoByte(255) = 255)
    Exit Function
Fail:
    test_lobyte_255 = False
End Function

' Requirement 2.4: Str2ByteArr produces correct length and ASCII byte values
Private Function test_str2bytearr() As Boolean
    On Error GoTo Fail
    Dim arr() As Byte
    Call AO20CryptoSysWrapper.Str2ByteArr("ABC", arr)
    ' Array should be 0-based with 3 elements
    If UBound(arr) - LBound(arr) + 1 <> 3 Then
        test_str2bytearr = False
        Exit Function
    End If
    ' A=65, B=66, C=67
    test_str2bytearr = (arr(0) = 65) And (arr(1) = 66) And (arr(2) = 67)
    Exit Function
Fail:
    test_str2bytearr = False
End Function

' Requirement 2.5: ByteArr2String returns correct ASCII string
Private Function test_bytearr2string() As Boolean
    On Error GoTo Fail
    Dim arr(0 To 2) As Byte
    arr(0) = 72   ' H
    arr(1) = 105  ' i
    arr(2) = 33   ' !
    test_bytearr2string = (AO20CryptoSysWrapper.ByteArr2String(arr) = "Hi!")
    Exit Function
Fail:
    test_bytearr2string = False
End Function

' Requirement 2.7: CopyBytes copies at offset, other bytes unchanged
Private Function test_copybytes() As Boolean
    On Error GoTo Fail
    Dim src(0 To 1) As Byte
    src(0) = 10
    src(1) = 20
    
    Dim dst(0 To 4) As Byte
    dst(0) = 99
    dst(1) = 99
    dst(2) = 99
    dst(3) = 99
    dst(4) = 99
    
    Call AO20CryptoSysWrapper.CopyBytes(src, dst, 2, 2)
    
    ' Bytes at offset 2 and 3 should be copied from src
    ' Bytes at 0, 1, 4 should remain 99
    test_copybytes = (dst(0) = 99) And (dst(1) = 99) And _
                     (dst(2) = 10) And (dst(3) = 20) And _
                     (dst(4) = 99)
    Exit Function
Fail:
    test_copybytes = False
End Function

' Requirement 2.8: ByteArrayToHex produces hex representation
' VB6 Hex$() does not zero-pad, so &H0A becomes "A" not "0A"
Private Function test_bytearraytohex() As Boolean
    On Error GoTo Fail
    Dim arr(0 To 1) As Byte
    arr(0) = &HFF
    arr(1) = &HA
    test_bytearraytohex = (AO20CryptoSysWrapper.ByteArrayToHex(arr) = "FF A")
    Exit Function
Fail:
    test_bytearraytohex = False
End Function

' Requirement 2.9: ByteArrayToDecimalString produces decimal representation
Private Function test_bytearraytodecimalstring() As Boolean
    On Error GoTo Fail
    Dim arr(0 To 1) As Byte
    arr(0) = 255
    arr(1) = 10
    test_bytearraytodecimalstring = (AO20CryptoSysWrapper.ByteArrayToDecimalString(arr) = "255 10")
    Exit Function
Fail:
    test_bytearraytodecimalstring = False
End Function

' --------------------------------------------------------------------------
' Property-based tests
' --------------------------------------------------------------------------

' Feature: unit-test-coverage-tier4, Property 4: HiByte/LoByte/MakeInt round-trip
' Validates: Requirements 2.3
Private Function test_pbt_makeint_roundtrip() As Boolean
    On Error GoTo Fail
    
    Dim n As Long
    
    For n = 0 To 32767
        If AO20CryptoSysWrapper.MakeInt(AO20CryptoSysWrapper.LoByte(CInt(n)), AO20CryptoSysWrapper.hiByte(CInt(n))) <> CInt(n) Then
            test_pbt_makeint_roundtrip = False
            Exit Function
        End If
    Next n
    
    test_pbt_makeint_roundtrip = True
    Exit Function
Fail:
    test_pbt_makeint_roundtrip = False
End Function

' Feature: unit-test-coverage-tier4, Property 5: Str2ByteArr/ByteArr2String round-trip
' Validates: Requirements 2.6
Private Function test_pbt_str2bytearr_roundtrip() As Boolean
    On Error GoTo Fail
    
    Dim i As Long
    Dim j As Long
    Dim s As String
    Dim arr() As Byte
    Dim result As String
    Dim strLen As Long
    Dim charCode As Long
    
    For i = 1 To 110
        ' Generate an ASCII string of length i with chars in range 1-127
        s = vbNullString
        strLen = ((i - 1) Mod 20) + 1
        For j = 1 To strLen
            charCode = ((i * 7 + j * 13) Mod 127) + 1  ' range 1-127
            s = s & Chr$(charCode)
        Next j
        
        Call AO20CryptoSysWrapper.Str2ByteArr(s, arr)
        result = AO20CryptoSysWrapper.ByteArr2String(arr)
        
        If result <> s Then
            test_pbt_str2bytearr_roundtrip = False
            Exit Function
        End If
    Next i
    
    test_pbt_str2bytearr_roundtrip = True
    Exit Function
Fail:
    test_pbt_str2bytearr_roundtrip = False
End Function

' Feature: unit-test-coverage-tier4, Property 6: CopyBytes correctness
' Validates: Requirements 2.7
Private Function test_pbt_copybytes_correctness() As Boolean
    On Error GoTo Fail
    
    Dim i As Long
    Dim j As Long
    Dim srcSize As Long
    Dim dstSize As Long
    Dim offset As Long
    Dim src() As Byte
    Dim dst() As Byte
    Dim origDst() As Byte
    
    For i = 1 To 110
        ' Generate varying source sizes (1..10) and offsets
        srcSize = ((i - 1) Mod 10) + 1
        offset = (i - 1) Mod 5
        dstSize = srcSize + offset + 2  ' ensure dst is large enough with room to spare
        
        ReDim src(0 To srcSize - 1)
        ReDim dst(0 To dstSize - 1)
        ReDim origDst(0 To dstSize - 1)
        
        ' Fill source with deterministic values
        For j = 0 To srcSize - 1
            src(j) = CByte((i * 3 + j * 7) Mod 256)
        Next j
        
        ' Fill dest with a sentinel value
        For j = 0 To dstSize - 1
            dst(j) = CByte(200)
            origDst(j) = CByte(200)
        Next j
        
        Call AO20CryptoSysWrapper.CopyBytes(src, dst, srcSize, offset)
        
        ' Verify copied bytes match source
        For j = 0 To srcSize - 1
            If dst(j + offset) <> src(j) Then
                test_pbt_copybytes_correctness = False
                Exit Function
            End If
        Next j
        
        ' Verify non-copied bytes are unchanged
        For j = 0 To dstSize - 1
            If j < offset Or j >= offset + srcSize Then
                If dst(j) <> origDst(j) Then
                    test_pbt_copybytes_correctness = False
                    Exit Function
                End If
            End If
        Next j
    Next i
    
    test_pbt_copybytes_correctness = True
    Exit Function
Fail:
    test_pbt_copybytes_correctness = False
End Function

' Feature: unit-test-coverage-tier4, Property 7: Byte array to string representation
' Validates: Requirements 2.8, 2.9
Private Function test_pbt_byte_to_string_repr() As Boolean
    On Error GoTo Fail
    
    Dim i As Long
    Dim j As Long
    Dim arrSize As Long
    Dim arr() As Byte
    Dim hexResult As String
    Dim decResult As String
    Dim hexTokens() As String
    Dim decTokens() As String
    Dim expectedHex As String
    Dim expectedDec As String
    
    For i = 1 To 110
        ' Generate arrays of size 1..10
        arrSize = ((i - 1) Mod 10) + 1
        ReDim arr(0 To arrSize - 1)
        
        For j = 0 To arrSize - 1
            arr(j) = CByte((i * 11 + j * 17) Mod 256)
        Next j
        
        hexResult = AO20CryptoSysWrapper.ByteArrayToHex(arr)
        decResult = AO20CryptoSysWrapper.ByteArrayToDecimalString(arr)
        
        ' Split results into tokens
        hexTokens = Split(hexResult, " ")
        decTokens = Split(decResult, " ")
        
        ' Verify token count matches array size
        If UBound(hexTokens) - LBound(hexTokens) + 1 <> arrSize Then
            test_pbt_byte_to_string_repr = False
            Exit Function
        End If
        
        If UBound(decTokens) - LBound(decTokens) + 1 <> arrSize Then
            test_pbt_byte_to_string_repr = False
            Exit Function
        End If
        
        ' Verify each hex token is the hex representation of the byte
        For j = 0 To arrSize - 1
            expectedHex = Hex$(arr(j))
            If hexTokens(j) <> expectedHex Then
                test_pbt_byte_to_string_repr = False
                Exit Function
            End If
        Next j
        
        ' Verify each decimal token is the decimal representation of the byte
        For j = 0 To arrSize - 1
            expectedDec = CStr(CLng(arr(j)))
            If decTokens(j) <> expectedDec Then
                test_pbt_byte_to_string_repr = False
                Exit Function
            End If
        Next j
    Next i
    
    test_pbt_byte_to_string_repr = True
    Exit Function
Fail:
    test_pbt_byte_to_string_repr = False
End Function

#End If
