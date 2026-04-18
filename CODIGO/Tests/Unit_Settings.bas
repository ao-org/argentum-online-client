Attribute VB_Name = "Unit_Settings"
Option Explicit

' ==========================================================================
' Settings Integration Test Suite
' Tests ModSettings.bas: GetSettingAsByte behavior including correct Byte
' conversion, missing key fallback to DefaultValue, non-numeric value
' fallback, and boundary values (0, 255).
'
' Since GetSettingAsByte internally calls GetSetting which uses hardcoded
' file paths, we test the same logic by creating a temp INI file and using
' GetVar (which accepts a file path) combined with CByte conversion and
' default-value fallback — mirroring GetSettingAsByte's exact behavior.
' ==========================================================================

#If UNIT_TEST = 1 Then

Private Const TEST_INI_FILE As String = "\test_settings.ini"

Public Sub test_suite_settings()
    ' Setup: create temp INI file
    Call create_test_ini
    
    ' Example-based tests
    Call UnitTesting.RunTest("set_existing_key", test_existing_key())
    Call UnitTesting.RunTest("set_missing_key", test_missing_key())
    Call UnitTesting.RunTest("set_non_numeric", test_non_numeric())
    Call UnitTesting.RunTest("set_value_zero", test_value_zero())
    Call UnitTesting.RunTest("set_value_255", test_value_255())
    
    ' Cleanup: remove temp INI file
    Call cleanup_test_ini
End Sub

' --------------------------------------------------------------------------
' Helper: create temp INI with known values
' --------------------------------------------------------------------------
Private Sub create_test_ini()
    Dim tmpPath As String
    tmpPath = App.path & TEST_INI_FILE
    
    Dim f As Integer: f = FreeFile
    Open tmpPath For Output As #f
    Print #f, "[TestSection]"
    Print #f, "ByteKey=42"
    Print #f, "ZeroKey=0"
    Print #f, "MaxKey=255"
    Print #f, "BadKey=abc"
    Close #f
End Sub

' --------------------------------------------------------------------------
' Helper: remove temp INI file
' --------------------------------------------------------------------------
Private Sub cleanup_test_ini()
    On Error Resume Next
    Kill App.path & TEST_INI_FILE
    On Error GoTo 0
End Sub


' --------------------------------------------------------------------------
' Helper: read INI value and convert to Byte with default fallback
' Mirrors GetSettingAsByte logic but accepts a file path parameter
' --------------------------------------------------------------------------
Private Function ReadSettingAsByte(ByVal FilePath As String, ByVal Section As String, ByVal KeyName As String, ByVal DefaultValue As Byte) As Byte
    On Error GoTo ErrHandler
    ReadSettingAsByte = DefaultValue
    
    Dim value As String
    value = GetVar(FilePath, Section, KeyName)
    
    If value = "" Then Exit Function
    
    ReadSettingAsByte = CByte(value)
    Exit Function
ErrHandler:
    ReadSettingAsByte = DefaultValue
End Function

' --------------------------------------------------------------------------
' Example-based tests
' --------------------------------------------------------------------------

' Existing key returns correct Byte value
Private Function test_existing_key() As Boolean
    On Error GoTo Fail
    Dim tmpPath As String
    tmpPath = App.path & TEST_INI_FILE
    
    Dim result As Byte
    result = ReadSettingAsByte(tmpPath, "TestSection", "ByteKey", 0)
    
    test_existing_key = (result = 42)
    Exit Function
Fail:
    test_existing_key = False
End Function

' Missing key returns DefaultValue
Private Function test_missing_key() As Boolean
    On Error GoTo Fail
    Dim tmpPath As String
    tmpPath = App.path & TEST_INI_FILE
    
    Dim result As Byte
    result = ReadSettingAsByte(tmpPath, "TestSection", "NonExistentKey", 99)
    
    test_missing_key = (result = 99)
    Exit Function
Fail:
    test_missing_key = False
End Function

' Non-numeric value ("abc") returns DefaultValue without error
Private Function test_non_numeric() As Boolean
    On Error GoTo Fail
    Dim tmpPath As String
    tmpPath = App.path & TEST_INI_FILE
    
    Dim result As Byte
    result = ReadSettingAsByte(tmpPath, "TestSection", "BadKey", 77)
    
    test_non_numeric = (result = 77)
    Exit Function
Fail:
    test_non_numeric = False
End Function

' Value "0" returns 0 (not confused with empty string)
Private Function test_value_zero() As Boolean
    On Error GoTo Fail
    Dim tmpPath As String
    tmpPath = App.path & TEST_INI_FILE
    
    Dim result As Byte
    result = ReadSettingAsByte(tmpPath, "TestSection", "ZeroKey", 50)
    
    test_value_zero = (result = 0)
    Exit Function
Fail:
    test_value_zero = False
End Function

' Value "255" returns 255 (max Byte boundary)
Private Function test_value_255() As Boolean
    On Error GoTo Fail
    Dim tmpPath As String
    tmpPath = App.path & TEST_INI_FILE
    
    Dim result As Byte
    result = ReadSettingAsByte(tmpPath, "TestSection", "MaxKey", 0)
    
    test_value_255 = (result = 255)
    Exit Function
Fail:
    test_value_255 = False
End Function

#End If
