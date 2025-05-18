Attribute VB_Name = "ModAuth"
'    Argentum 20 - Game Client Program
'    Copyright (C) 2022 - Noland Studios
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Affero General Public License for more details.
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
'
Option Explicit

Public Enum e_state
    Idle = 0
    RequestAccountLogin
    AccountLogged
    RequestCharList
    RequestLogout
    RequestSignUp
    RequestValidateAccount
    RequestForgotPassword
    RequestResetPassword
    RequestDeleteChar
    ConfirmDeleteChar
    RequestVerificationCode
    RequestTransferCharacter
End Enum

Public Enum e_operation
    Authenticate = 0
    SignUp
    ValidateAccount
    ForgotPassword
    ResetPassword
    deletechar
    ConfirmDeleteChar
    RequestVerificationCode
    transfercharacter
End Enum

Public Type t_CreateAccountInfo
    Email As String
    Password As String
    Name As String
    Surname As String
End Type

Public NewAccountData As t_CreateAccountInfo
Public SessionOpened As Boolean
Public Auth_state As e_state
Public LoginOperation As e_operation
Public public_key() As Byte
Public encrypted_session_token As String
Public decrypted_session_token As String
Public authenticated_decrypted_session_token As String
Public delete_char_validate_code As String


Public Sub AuthSocket_DataArrival(ByVal BytesTotal As Long)
    
    If Connected Then
        Exit Sub
    End If
    
    If Not SessionOpened Then
        Call HandleOpenSession(BytesTotal)
        If SessionOpened Then
            Select Case Auth_state
                Case e_state.RequestAccountLogin
                    Call SendAccountLoginRequest
                Case e_state.RequestSignUp
                    Call SendSignUpRequest
                Case e_state.RequestValidateAccount
                    Call SendValidateAccount
                Case e_state.RequestForgotPassword
                    Call SendRequestForgotPassword
                Case e_state.RequestResetPassword
                    Call SendRequestResetPassword
                Case e_state.RequestDeleteChar
                    Call SendDeleteCharRequest
                Case e_state.ConfirmDeleteChar
                    Call SendConfirmDeleteChar
                Case e_state.RequestVerificationCode
                    Call SendRequestVerificationCode
                Case e_state.RequestTransferCharacter
                    Call SendRequestTransferCharacter
            End Select
        End If
        Exit Sub
    End If
    
    Select Case Auth_state
        Case e_state.RequestAccountLogin
            Call HandleAccountLogin(BytesTotal)
        Case e_state.RequestCharList
            Call HandlePCList(BytesTotal)
        Case e_state.RequestSignUp
            Call HandleSignUpRequest(BytesTotal)
        Case e_state.RequestValidateAccount
            Call HandleValidateAccountRequest(BytesTotal)
        Case e_state.RequestForgotPassword
            Call HandleRequestForgotPassword(BytesTotal)
        Case e_state.RequestResetPassword
            Call HandleRequestResetPassword(BytesTotal)
        Case e_state.RequestDeleteChar
            Call HandleDeleteCharRequest(BytesTotal)
        Case e_state.ConfirmDeleteChar
            Call HandleConfirmDeleteChar(BytesTotal)
        Case e_state.RequestVerificationCode
            Call HandleRequestVerificationCode(BytesTotal)
        Case e_state.RequestTransferCharacter
            Call HandleTransferCharRequest(BytesTotal)

    End Select
    
End Sub

Public Sub OpenSessionRequest()
    
    SessionOpened = False
    
    Dim arr(0 To 3) As Byte
    arr(0) = &H0
    arr(1) = &HAA
    arr(2) = &H0
    arr(3) = &H4
    Call frmConnect.AuthSocket.SendData(arr)
    
End Sub
Public Sub DebugPrint(ByVal str As String, Optional ByVal int1 As Integer = 0, Optional ByVal int2 As Integer = 0, Optional ByVal int3 As Integer = 0, Optional ByVal asd As Boolean = False)

    frmDebug.add_text_tracebox (str)
    
End Sub

Public Sub SendAccountLoginRequest()
    Dim userName As String
    Dim Password As String
    Dim len_encrypted_username As Integer
    Dim len_encrypted_password As Integer
    
    Dim login_request() As Byte
    Dim packet_size As Integer
    Dim offset_login_request As Long
    
    frmDebug.add_text_tracebox "SendAccountLoginRequest"

    userName = CuentaEmail
    Password = CuentaPassword
    
    Dim encrypted_username() As Byte
    Dim encrypted_username_b64 As String
    
    Dim encrypted_password() As Byte
    Dim encrypted_password_b64 As String
    
    
    encrypted_username_b64 = AO20CryptoSysWrapper.Encrypt(cnvHexStrFromBytes(public_key), userName)
    encrypted_password_b64 = AO20CryptoSysWrapper.Encrypt(cnvHexStrFromBytes(public_key), Password)
    
    
    Call Str2ByteArr(encrypted_username_b64, encrypted_username)
    Call Str2ByteArr(encrypted_password_b64, encrypted_password)
    
    
    Dim len_username As Integer
    Dim len_password As Integer
    
    len_username = Len(encrypted_username_b64)
    len_password = Len(encrypted_password_b64)
    
    ReDim login_request(1 To (2 + 2 + 2 + len_username + 2 + len_password))
    
    packet_size = UBound(login_request)
    
    login_request(1) = &HDE
    login_request(2) = &HAD
    
    'Siguientes 2 bytes indican tamaño total del paquete
    login_request(3) = hiByte(packet_size)
    login_request(4) = LoByte(packet_size)
    
    'Los siguientes 2 bytes son el SIZE_ENCRYPTED_USER
    login_request(5) = hiByte(len_username)
    login_request(6) = LoByte(len_username)
    Call AO20CryptoSysWrapper.CopyBytes(encrypted_username, login_request, len_username, 7)
    
    offset_login_request = 7 + UBound(encrypted_username)
        
    login_request(offset_login_request + 1) = hiByte(len_password)
    login_request(offset_login_request + 2) = LoByte(len_password)
    
    Call AO20CryptoSysWrapper.CopyBytes(encrypted_password, login_request, len_password, offset_login_request + 3)
    
    Call frmConnect.AuthSocket.SendData(login_request)
    
    Auth_state = e_state.RequestAccountLogin
    
End Sub

Public Sub SendRequestVerificationCode()
    Dim userName As String
    Dim len_encrypted_username As Integer
    
    Dim login_request() As Byte
    Dim packet_size As Integer
    Dim offset_login_request As Long
    
    Call DebugPrint("SendRequestVerificationCode", 255, 255, 255, True)
    
    userName = CuentaEmail
    
    Dim encrypted_username() As Byte
    Dim encrypted_username_b64 As String
    Debug.Assert Len(userName) > 0
    encrypted_username_b64 = AO20CryptoSysWrapper.Encrypt(cnvHexStrFromBytes(public_key), userName)
    Call Str2ByteArr(encrypted_username_b64, encrypted_username)
    
    Dim len_username As Integer
    len_username = Len(encrypted_username_b64)
    
    ReDim login_request(1 To (2 + 2 + 2 + len_username + 2 + len_username))
    
    packet_size = UBound(login_request)
    
    login_request(1) = &HDA
    login_request(2) = &HAB
    
    'Siguientes 2 bytes indican tamaño total del paquete
    login_request(3) = hiByte(packet_size)
    login_request(4) = LoByte(packet_size)
    
    'Los siguientes 2 bytes son el SIZE_ENCRYPTED_USER
    login_request(5) = hiByte(len_username)
    login_request(6) = LoByte(len_username)
    Call AO20CryptoSysWrapper.CopyBytes(encrypted_username, login_request, len_username, 7)
    
    offset_login_request = 7 + UBound(encrypted_username)
        
    login_request(offset_login_request + 1) = hiByte(len_username)
    login_request(offset_login_request + 2) = LoByte(len_username)
    
    Call AO20CryptoSysWrapper.CopyBytes(encrypted_username, login_request, len_username, offset_login_request + 3)
    
    Call frmConnect.AuthSocket.SendData(login_request)
    
    Auth_state = e_state.RequestVerificationCode
    
End Sub

Public Sub HandleRequestVerificationCode(ByVal BytesTotal As Long)

    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    Call DebugPrint("HandleRequestVerificationCode", 255, 255, 255, True)
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    Dim Data() As Byte
    
    frmConnect.AuthSocket.PeekData Data, vbByte, BytesTotal
    
    frmConnect.AuthSocket.GetData Data, vbByte, 2
    
    If Data(0) = &H11 And Data(1) = &H11 Then
        Call DebugPrint("REQUEST-VERIFICATION-CODE-OK", 0, 255, 0, True)
        frmConnect.AuthSocket.GetData Data, vbByte, 2
        Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_CODIGO_ENVIADO") & " " & CuentaEmail & ".", False, False)
        
    Else
       Call DebugPrint("ERROR", 255, 0, 0, True)
        frmConnect.AuthSocket.GetData Data, vbByte, 4
        Select Case MakeInt(Data(3), Data(2))
            Case 1
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_ACCOUNT_NO_EXISTE"), False, False)
            Case 11
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_ACCOUNT_ACTIVADA"), False, False)
            Case 12
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_ACCOUNT_NO_EXISTE"), False, False)
            Case Else
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_NO_CONEXION") & " " & AO20CryptoSysWrapper.ByteArrayToHex(Data), False, False)

        End Select
    End If
    
    Auth_state = e_state.Idle
        
End Sub

Public Sub SendRequestForgotPassword()
    Dim userName As String
    Dim len_encrypted_username As Integer
    
    Dim login_request() As Byte
    Dim packet_size As Integer
    Dim offset_login_request As Long
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    Call DebugPrint("SendRequestForgotPassword", 255, 255, 255, True)
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    userName = CuentaEmail
    
    Dim encrypted_username() As Byte
    Dim encrypted_username_b64 As String
    
    encrypted_username_b64 = AO20CryptoSysWrapper.Encrypt(cnvHexStrFromBytes(public_key), userName)
    
    Call Str2ByteArr(encrypted_username_b64, encrypted_username)
    
    Dim len_username As Integer
    
    len_username = Len(encrypted_username_b64)
    
    ReDim login_request(1 To (2 + 2 + 2 + len_username + 2 + len_username))
    
    packet_size = UBound(login_request)
    
    login_request(1) = &HCB
    login_request(2) = &HCB
    
    'Siguientes 2 bytes indican tamaño total del paquete
    login_request(3) = hiByte(packet_size)
    login_request(4) = LoByte(packet_size)
    
    'Los siguientes 2 bytes son el SIZE_ENCRYPTED_USER
    login_request(5) = hiByte(len_username)
    login_request(6) = LoByte(len_username)
    Call AO20CryptoSysWrapper.CopyBytes(encrypted_username, login_request, len_username, 7)
    
    offset_login_request = 7 + UBound(encrypted_username)
    login_request(offset_login_request + 1) = hiByte(len_username)
    login_request(offset_login_request + 2) = LoByte(len_username)
    
    Call AO20CryptoSysWrapper.CopyBytes(encrypted_username, login_request, len_username, offset_login_request + 3)
    Call frmConnect.AuthSocket.SendData(login_request)
    
    Auth_state = e_state.RequestForgotPassword
    
End Sub

Public Sub SendValidateAccount()
    Dim userName As String
    Dim validate_code As String
    Dim len_encrypted_username As Integer
    Dim len_encrypted_validate_code As Integer
    
    Dim login_request() As Byte
    Dim packet_size As Integer
    Dim offset_login_request As Long
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    Call DebugPrint("SendValidateAccount", 255, 255, 255, True)
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    userName = CuentaEmail
    validate_code = ValidationCode
    
    Dim encrypted_username() As Byte
    Dim encrypted_username_b64 As String
    
    Dim encrypted_validate_code() As Byte
    Dim encrypted_validate_code_b64 As String
    
    
    encrypted_username_b64 = AO20CryptoSysWrapper.Encrypt(cnvHexStrFromBytes(public_key), userName)
    encrypted_validate_code_b64 = AO20CryptoSysWrapper.Encrypt(cnvHexStrFromBytes(public_key), validate_code)
        
    Call Str2ByteArr(encrypted_username_b64, encrypted_username)
    Call Str2ByteArr(encrypted_validate_code_b64, encrypted_validate_code)
    
    
    Dim len_username As Integer
    Dim len_validate_code As Integer
    
    len_username = Len(encrypted_username_b64)
    len_validate_code = Len(encrypted_validate_code_b64)
    
    ReDim login_request(1 To (2 + 2 + 2 + len_username + 2 + len_validate_code))
    
    packet_size = UBound(login_request)
    
    login_request(1) = &HBA
    login_request(2) = &HAD
    
    'Siguientes 2 bytes indican tamaño total del paquete
    login_request(3) = hiByte(packet_size)
    login_request(4) = LoByte(packet_size)
    
    'Los siguientes 2 bytes son el SIZE_ENCRYPTED_USER
    login_request(5) = hiByte(len_username)
    login_request(6) = LoByte(len_username)
    Call AO20CryptoSysWrapper.CopyBytes(encrypted_username, login_request, len_username, 7)
    
    offset_login_request = 7 + UBound(encrypted_username)
        
    login_request(offset_login_request + 1) = hiByte(len_validate_code)
    login_request(offset_login_request + 2) = LoByte(len_validate_code)
    
    Call AO20CryptoSysWrapper.CopyBytes(encrypted_validate_code, login_request, len_validate_code, offset_login_request + 3)
    
    Call frmConnect.AuthSocket.SendData(login_request)
    
    Auth_state = e_state.RequestValidateAccount
    
End Sub

Public Sub SendRequestResetPassword()
    Dim userName As String
    Dim Password As String
    Dim validate_code As String
    Dim len_encrypted_username As Integer
    Dim len_encrypted_password As Integer
    Dim len_encrypted_validate_code As Integer
    
    Dim login_request() As Byte
    Dim packet_size As Integer
    Dim offset_login_request As Long
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    Call DebugPrint("SendRequestResetPassword", 255, 255, 255, True)
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    userName = CuentaEmail
    Password = CuentaPassword
    validate_code = ValidationCode
    Debug.Assert Len(validate_code) > 0 And Len(userName) > 0 And Len(Password) > 0
    
    Dim encrypted_username() As Byte
    Dim encrypted_username_b64 As String
    
    
    Dim encrypted_password() As Byte
    Dim encrypted_password_b64 As String
    
    Dim encrypted_validate_code() As Byte
    Dim encrypted_validate_code_b64 As String
    
    
    encrypted_username_b64 = AO20CryptoSysWrapper.Encrypt(cnvHexStrFromBytes(public_key), userName)
    encrypted_password_b64 = AO20CryptoSysWrapper.Encrypt(cnvHexStrFromBytes(public_key), Password)
    encrypted_validate_code_b64 = AO20CryptoSysWrapper.Encrypt(cnvHexStrFromBytes(public_key), validate_code)
        
    Call Str2ByteArr(encrypted_username_b64, encrypted_username)
    Call Str2ByteArr(encrypted_password_b64, encrypted_password)
    Call Str2ByteArr(encrypted_validate_code_b64, encrypted_validate_code)
    
    Dim len_username As Integer
    Dim len_password As Integer
    Dim len_validate_code As Integer
    
    len_username = Len(encrypted_username_b64)
    len_password = Len(encrypted_password_b64)
    len_validate_code = Len(encrypted_validate_code_b64)
    
    ReDim login_request(1 To (2 + 2 + 2 + len_username + 2 + len_validate_code + 2 + len_password))
    
    packet_size = UBound(login_request)
    
    login_request(1) = &HFB
    login_request(2) = &HFB
    
    'Siguientes 2 bytes indican tamaño total del paquete
    login_request(3) = hiByte(packet_size)
    login_request(4) = LoByte(packet_size)
    
    'Los siguientes 2 bytes son el SIZE_ENCRYPTED_USER
    login_request(5) = hiByte(len_username)
    login_request(6) = LoByte(len_username)
    Call AO20CryptoSysWrapper.CopyBytes(encrypted_username, login_request, len_username, 7)
    
    offset_login_request = 7 + UBound(encrypted_username)
        
    login_request(offset_login_request + 1) = hiByte(len_password)
    login_request(offset_login_request + 2) = LoByte(len_password)
    
    Call AO20CryptoSysWrapper.CopyBytes(encrypted_password, login_request, len_password, offset_login_request + 3)
    
    offset_login_request = offset_login_request + 3 + UBound(encrypted_password)
    
    login_request(offset_login_request + 1) = hiByte(len_validate_code)
    login_request(offset_login_request + 2) = LoByte(len_validate_code)
    
    Call AO20CryptoSysWrapper.CopyBytes(encrypted_validate_code, login_request, len_validate_code, offset_login_request + 3)
    
    Call frmConnect.AuthSocket.SendData(login_request)
    
    Auth_state = e_state.RequestResetPassword
    
End Sub
Public Sub LogOutRequest()
    
    Dim logout_request() As Byte
    Dim packet_size As Integer
    
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    Call DebugPrint("LogOutRequest", 255, 255, 255, True)
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    
    ReDim logout_request(1 To (4 + Len(encrypted_session_token)))
    
    packet_size = UBound(logout_request)
    
    logout_request(1) = &H1
    logout_request(2) = &H1
    
    'Siguientes 2 bytes indican tamaño total del paquete
    logout_request(3) = hiByte(packet_size)
    logout_request(4) = LoByte(packet_size)
    Dim encrypted_session_token_byte() As Byte
    Call Str2ByteArr(encrypted_session_token, encrypted_session_token_byte)
    Call AO20CryptoSysWrapper.CopyBytes(encrypted_session_token_byte, logout_request, Len(encrypted_session_token), 5)
    
    Call frmConnect.AuthSocket.SendData(logout_request)
    
    Auth_state = e_state.RequestLogout
    
End Sub
Public Sub HandleLogOutRequest(ByVal BytesTotal As Long)
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    Call DebugPrint("HandleLogOutRequest", 255, 255, 255, True)
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    
    Dim Data() As Byte
    
    frmConnect.AuthSocket.PeekData Data, vbByte, BytesTotal
    
    frmConnect.AuthSocket.GetData Data, vbByte, 2
    
    If Data(0) = &H20 And Data(1) = &H22 Then
        Call DebugPrint("LOGOUT_OKAY", 0, 255, 0, True)
        
        frmConnect.AuthSocket.GetData Data, vbByte, 2
        
        Auth_state = e_state.Idle
    Else
       Call DebugPrint("ERROR", 255, 0, 0, True)
        frmConnect.AuthSocket.GetData Data, vbByte, 4
        Select Case MakeInt(Data(3), Data(2))
            Case 41
                Call DebugPrint("Not logged yet.", 255, 255, 0)
        End Select
    End If
    Auth_state = e_state.Idle
End Sub
Public Sub SendRequestTransferCharacter()
    Dim JSON As String
    Dim len_encrypted_username As Integer
    Dim transfer_request() As Byte
    Dim packet_size As Integer
    
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    Call DebugPrint("SendRequestTransferCharacter", 255, 255, 255, True)
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    Dim old_owner_str As String
    old_owner_str = CuentaEmail
    
    JSON = ""
    JSON = "{ "
    JSON = JSON & "  ""currentOwner"": """ & old_owner_str & """ , "
    JSON = JSON & "  ""pc"": """ & TransferCharname & """ , "
    JSON = JSON & "  ""token"": """ & authenticated_decrypted_session_token & """ , "
    JSON = JSON & "  ""newOwner"": """ & TransferCharNewOwner & """"
    JSON = JSON & " }"

    Dim encrypted_json() As Byte
    Dim encrypted_json_b64 As String
    
    encrypted_json_b64 = AO20CryptoSysWrapper.Encrypt(cnvHexStrFromBytes(public_key), JSON)
        
    Call Str2ByteArr(encrypted_json_b64, encrypted_json)
        
    Dim len_json As Integer
    len_json = Len(encrypted_json_b64)
    
    ReDim transfer_request(1 To (2 + 2 + len_json))
    
    packet_size = UBound(transfer_request)
    
    transfer_request(1) = &H20
    transfer_request(2) = &H25
    
    'Siguientes 2 bytes indican tamaño total del paquete
    transfer_request(3) = hiByte(packet_size)
    transfer_request(4) = LoByte(packet_size)
    
    Call AO20CryptoSysWrapper.CopyBytes(encrypted_json, transfer_request, len_json, 5)

    Call frmConnect.AuthSocket.SendData(transfer_request)
    
    Auth_state = e_state.RequestTransferCharacter
    

End Sub
Public Sub SendSignUpRequest()

    Dim JSON As String
    Dim len_encrypted_username As Integer
    Dim login_request() As Byte
    Dim packet_size As Integer
    
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    Call DebugPrint("SendSignUpRequest", 255, 255, 255, True)
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    
    
    JSON = ""
    
    JSON = "{ ""language"": ""english"", ""password"": """ & NewAccountData.Password & """, "
    JSON = JSON & """passwordrecovery"": [{""secretanswer1"": ""Satanas"","
    JSON = JSON & """secretanswer2"": ""Rojo"", "
    JSON = JSON & """secretquestion1"": ""Cual es el nombre de mi primer mascota?"","
    JSON = JSON & """secretquestion2"": ""Cual es mi color favorito?""}],"
    
    JSON = JSON & """personal"":[{"
    JSON = JSON & """dob"": ""23-12-1990"","
    JSON = JSON & """email"": """ & NewAccountData.Email & ""","
    JSON = JSON & """firstname"": """ & NewAccountData.Name & ""","
    JSON = JSON & """lastname"": """ & NewAccountData.Surname & ""","
    JSON = JSON & """mobile"": """ & NewAccountData.Surname & ""","
    JSON = JSON & """pob"": """ & NewAccountData.Surname & """}],"
    JSON = JSON & """username"": """ & NewAccountData.Email & """}"
    
    
    Dim encrypted_json() As Byte
    Dim encrypted_json_b64 As String
    
    encrypted_json_b64 = AO20CryptoSysWrapper.Encrypt(cnvHexStrFromBytes(public_key), JSON)
        
    Call Str2ByteArr(encrypted_json_b64, encrypted_json)
        
    Dim len_json As Integer
    len_json = Len(encrypted_json_b64)
    
    ReDim login_request(1 To (2 + 2 + len_json))
    
    packet_size = UBound(login_request)
    
    login_request(1) = &HBE
    login_request(2) = &HEF
    
    'Siguientes 2 bytes indican tamaño total del paquete
    login_request(3) = hiByte(packet_size)
    login_request(4) = LoByte(packet_size)
    
    Call AO20CryptoSysWrapper.CopyBytes(encrypted_json, login_request, len_json, 5)

    Call frmConnect.AuthSocket.SendData(login_request)
    
    Auth_state = e_state.RequestSignUp
    
End Sub


Public Sub HandleTransferCharRequest(ByVal BytesTotal As Long)

    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    Call DebugPrint("HandleTransferCharRequest", 255, 255, 255, True)
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    
    Dim Data() As Byte
    
    frmConnect.AuthSocket.PeekData Data, vbByte, BytesTotal
    frmConnect.AuthSocket.GetData Data, vbByte, 2

    If Data(0) = &H20 And Data(1) = &H26 Then
        Call DebugPrint("TRANSFER_CHARACTER_OKAY", 0, 255, 0, True)
        Call DisplayError(JsonLanguage.Item("MENSAJE_TRANSFERENCIA_REALIZADA"), "TRANSFER_CHARACTER_OKAY")
        Call EraseCharFromPjList(TransferCharname)
        frmConnect.AuthSocket.GetData Data, vbByte, 2
        Auth_state = e_state.Idle
    Else
        Call DebugPrint("TRANSFER CHAR ERROR", 255, 0, 0, True)
        frmConnect.AuthSocket.GetData Data, vbByte, 4
        Select Case MakeInt(Data(3), Data(2))

        Case 1
            Call DisplayError(JsonLanguage.Item("MENSAJE_CUENTA_INVALIDA"), "invalid-account")
        Case 3
            Call DisplayError(JsonLanguage.Item("MENSAJE_ERROR_BASE_DATOS"), "database-error")
        Case 12
            Call DisplayError(JsonLanguage.Item("MENSAJE_EMAIL_NO_VALIDO"), "invalid-email")
        Case 51
            Call DisplayError(JsonLanguage.Item("MENSAJE_NO_DUENO_PERSONAJE"), "not-char-owner")
        Case 52
            Call DisplayError(JsonLanguage.Item("MENSAJE_SOLICITUD_INVALIDA"), "invalid-request")
        Case 54
            Call DisplayError(JsonLanguage.Item("MENSAJE_NUEVO_DUENO_NO_EXISTE"), "newowner-not-exist")
        Case 55
            Call DisplayError(JsonLanguage.Item("MENSAJE_NO_ES_PATREON"), "not-patreon")
        Case 57
            Call DisplayError(JsonLanguage.Item("MENSAJE_CREDITOS_INSUFICIENTES"), "not-enough-credits")
        Case Else
            Call DisplayError(JsonLanguage.Item("MENSAJE_ERROR_DESCONOCIDO"), "unknown-error")

        End Select
    End If
    Auth_state = e_state.Idle

End Sub

Public Sub showValidateAccountControls()
       Call frmNewAccount.showValidateAccountControls
       frmNewAccount.txtValidateMail.Text = frmNewAccount.txtUsername

End Sub

Public Sub HandleSignUpRequest(ByVal BytesTotal As Long)
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    Call DebugPrint("HandleSignUpRequest", 255, 255, 255, True)
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    
    Dim Data() As Byte
    
    frmConnect.AuthSocket.PeekData Data, vbByte, BytesTotal
    
    frmConnect.AuthSocket.GetData Data, vbByte, 2
    
    If Data(0) = &HBF And Data(1) = &HB1 Then
        Call DebugPrint("SIGNUP_OKAY", 0, 255, 0, True)
        Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_CUENTA_CREADA"), False, False)
        frmConnect.AuthSocket.GetData Data, vbByte, 2
        Call showValidateAccountControls
        Auth_state = e_state.Idle
    Else
       Call DebugPrint("ERROR", 255, 0, 0, True)
        frmConnect.AuthSocket.GetData Data, vbByte, 4
        Select Case MakeInt(Data(3), Data(2))
            Case 0
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_USUARIO_EXISTE"), False, False)
            Case 9
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_ERROR_ENVIO_EMAIL"), False, False)
            Case 12
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_EMAIL_INVALIDO"), False, False)
            Case 14
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_PASSWORD_CORTO"), False, False)
            Case 15
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_PASSWORD_LARGO"), False, False)
            Case 16
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_PASSWORD_CARACTERES_INVALIDOS"), False, False)
            Case 18
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_USERNAME_CORTO"), False, False)
            Case 19
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_USERNAME_LARGO"), False, False)
            Case 20
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_USERNAME_CARACTERES_INVALIDOS"), False, False)
            Case 24
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_PASSWORD_NO_CONTIENE_USERNAME"), False, False)
            Case 32
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_USERNAME_NO_INICIA_CON_NUMERO"), False, False)
            Case 33
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_PASSWORD_SIN_LETRAS_MAYUSCULAS"), False, False)
            Case 34
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_PASSWORD_SIN_LETRAS_MINUSCULAS"), False, False)
            Case 35
                Call TextoAlAsistente("The password must contain at least than two numbers.", False, False)
            Case Else
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_ERROR_DESCONOCIDO"), False, False)
        End Select
    End If
    Auth_state = e_state.Idle
End Sub

Public Sub SendDeleteCharRequest()

    Dim JSON As String
    Dim len_encrypted_username As Integer
    Dim delete_char_request() As Byte
    Dim packet_size As Integer
    
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    Call DebugPrint("SendDeleteChar", 255, 255, 255, True)
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
        
    JSON = "{""username"": """ & CuentaEmail & """, ""pc"":""" & DeleteUser & """}"
    
    Dim encrypted_json() As Byte
    Dim encrypted_json_b64 As String
    
    encrypted_json_b64 = AO20CryptoSysWrapper.Encrypt(cnvHexStrFromBytes(public_key), JSON)
        
    Call Str2ByteArr(encrypted_json_b64, encrypted_json)
        
    Dim len_json As Integer
    len_json = Len(encrypted_json_b64)
    
    ReDim delete_char_request(1 To (2 + 2 + len_json))
    
    packet_size = UBound(delete_char_request)
    
    delete_char_request(1) = &H1
    delete_char_request(2) = &H5
    
    'Siguientes 2 bytes indican tamaño total del paquete
    delete_char_request(3) = hiByte(packet_size)
    delete_char_request(4) = LoByte(packet_size)
    
    Call AO20CryptoSysWrapper.CopyBytes(encrypted_json, delete_char_request, len_json, 5)

    Call frmConnect.AuthSocket.SendData(delete_char_request)
    
    Auth_state = e_state.RequestDeleteChar
    
End Sub

Public Sub HandleDeleteCharRequest(ByVal BytesTotal As Long)
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    Call DebugPrint("HandleDeleteCharRequest", 255, 255, 255, True)
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    Dim Data() As Byte
    frmConnect.AuthSocket.PeekData Data, vbByte, BytesTotal
    frmConnect.AuthSocket.GetData Data, vbByte, 2
    
    If Data(0) = &H1 And Data(1) = &H6 Then
        Call DebugPrint("DELETE_PC_REQUEST_OK", 0, 255, 0, True)
        Call DeleteCharRequestCode
        frmConnect.AuthSocket.GetData Data, vbByte, 2
        Auth_state = e_state.Idle
    Else
        Call DebugPrint("DELETE_PC_REQUEST_ERROR", 255, 0, 0, True)
        frmDeleteChar.Hide
        frmConnect.AuthSocket.GetData Data, vbByte, 4
    Select Case MakeInt(data(3), data(2))
        Case 1
            Call DisplayError(JsonLanguage.Item("MENSAJE_CUENTA_INVALIDA"), "invalid-account")
        Case 3
            Call DisplayError(JsonLanguage.Item("MENSAJE_ERROR_BASE_DATOS"), "database-error")
        Case 51
            Call DisplayError(JsonLanguage.Item("MENSAJE_NO_DUENO_PERSONAJE"), "invalid-character-owner")
        Case 65
            Call DisplayError(JsonLanguage.Item("MENSAJE_PERSONAJE_BLOQUEADO_MAO"), "locked-in-mao")
        Case Else
            Call DisplayError(JsonLanguage.Item("MENSAJE_ERROR_DESCONOCIDO"), "unknown-error")
    End Select

    End If
    Auth_state = e_state.Idle
End Sub

Public Sub SendConfirmDeleteChar()
    Dim userName As String
    Dim validate_code As String
    Dim len_encrypted_username As Integer
    Dim len_encrypted_validate_code As Integer
    
    Dim login_request() As Byte
    Dim packet_size As Integer
    Dim offset_login_request As Long
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    Call DebugPrint("SendConfirmDeleteChar", 255, 255, 255, True)
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    userName = DeleteUser
    validate_code = delete_char_validate_code
    
    Dim encrypted_username() As Byte
    Dim encrypted_username_b64 As String
    
    Dim encrypted_validate_code() As Byte
    Dim encrypted_validate_code_b64 As String
    
    
    encrypted_username_b64 = AO20CryptoSysWrapper.Encrypt(cnvHexStrFromBytes(public_key), userName)
    encrypted_validate_code_b64 = AO20CryptoSysWrapper.Encrypt(cnvHexStrFromBytes(public_key), validate_code)
        
    Call Str2ByteArr(encrypted_username_b64, encrypted_username)
    Call Str2ByteArr(encrypted_validate_code_b64, encrypted_validate_code)
    
    
    Dim len_username As Integer
    Dim len_validate_code As Integer
    
    len_username = Len(encrypted_username_b64)
    len_validate_code = Len(encrypted_validate_code_b64)
    
    ReDim login_request(1 To (2 + 2 + 2 + len_username + 2 + len_validate_code))
    
    packet_size = UBound(login_request)
    
    login_request(1) = &H1
    login_request(2) = &H8
    
    'Siguientes 2 bytes indican tamaño total del paquete
    login_request(3) = hiByte(packet_size)
    login_request(4) = LoByte(packet_size)
    
    'Los siguientes 2 bytes son el SIZE_ENCRYPTED_USER
    login_request(5) = hiByte(len_username)
    login_request(6) = LoByte(len_username)
    Call AO20CryptoSysWrapper.CopyBytes(encrypted_username, login_request, len_username, 7)
    
    offset_login_request = 7 + UBound(encrypted_username)
        
    login_request(offset_login_request + 1) = hiByte(len_validate_code)
    login_request(offset_login_request + 2) = LoByte(len_validate_code)
    
    Call AO20CryptoSysWrapper.CopyBytes(encrypted_validate_code, login_request, len_validate_code, offset_login_request + 3)
    
    Call frmConnect.AuthSocket.SendData(login_request)
    
    Auth_state = e_state.ConfirmDeleteChar
    
End Sub

Public Sub HandleConfirmDeleteChar(ByVal BytesTotal As Long)

    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    Call DebugPrint("HandleConfirmDeleteChar", 255, 255, 255, True)
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    Dim Data() As Byte
    
    frmConnect.AuthSocket.PeekData Data, vbByte, BytesTotal
    
    frmConnect.AuthSocket.GetData Data, vbByte, 2
    
    If Data(0) = &H1 And Data(1) = &H9 Then
        Call DebugPrint("DELETE-CHAR-OK", 0, 255, 0, True)
        Call DebugPrint(AO20CryptoSysWrapper.ByteArrayToHex(Data), 255, 255, 255)
        frmConnect.AuthSocket.GetData Data, vbByte, 2
        Auth_state = e_state.Idle
        Call DisplayError(JsonLanguage.Item("MENSAJE_PERSONAJE_BORRADO"), "delete-char-success")
        Call EraseCharFromPjList(DeleteUser)
    Else
       Call DebugPrint("ERROR", 255, 0, 0, True)
        frmConnect.AuthSocket.GetData Data, vbByte, 4
    Select Case MakeInt(data(3), data(2))
        Case 1
            Call DisplayError(JsonLanguage.Item("MENSAJE_NOMBRE_PERSONAJE_INVALIDO"), "invalid-character-name")
        Case 3
            Call DisplayError(JsonLanguage.Item("MENSAJE_ERROR_BASE_DATOS"), "database-error")
        Case 25
            Call DisplayError(JsonLanguage.Item("MENSAJE_CODIGO_INVALIDO"), "invalid-code")
            frmDeleteChar.Show , frmConnect
        Case Else
            Call DisplayError(JsonLanguage.Item("MENSAJE_ERROR_DESCONOCIDO") & ": " & AO20CryptoSysWrapper.ByteArrayToHex(data), "")
    End Select

    End If
        
End Sub

Public Sub PCListRequest()
    Dim userName As String
    Dim len_encrypted_username As Integer
    
    Dim packet_request() As Byte
    Dim charList_request() As Byte
    Dim offset_login_request As Long
    Dim packet_size As Integer

    frmDebug.add_text_tracebox "PCListRequest"

    userName = CuentaEmail
    
    Dim encrypted_username() As Byte
    Dim encrypted_username_b64 As String
    
    
    encrypted_username_b64 = AO20CryptoSysWrapper.Encrypt(cnvHexStrFromBytes(public_key), userName)
    
    Call Str2ByteArr(encrypted_username_b64, encrypted_username)
    
    
    Dim len_username As Integer
    
    len_username = Len(encrypted_username_b64)
    
    ReDim charList_request(1 To (2 + 2 + len_username))
    
    packet_size = UBound(charList_request)
    
    charList_request(1) = &H1
    charList_request(2) = &H2
    
    'Siguientes 2 bytes indican tamaño total del paquete
    charList_request(3) = hiByte(packet_size)
    charList_request(4) = LoByte(packet_size)
    
    Call AO20CryptoSysWrapper.CopyBytes(encrypted_username, charList_request, len_username, 5)
        
    Call frmConnect.AuthSocket.SendData(charList_request)
    
    Auth_state = e_state.RequestCharList
    
End Sub
Public Sub connectToLoginServer()

    frmConnect.AuthSocket.Close
    frmConnect.AuthSocket.RemoteHost = IPdelServidorLogin
    frmConnect.AuthSocket.RemotePort = PuertoDelServidorLogin
    frmDebug.add_text_tracebox "Servidor de Login " & IPdelServidorLogin & ":" & PuertoDelServidorLogin
    frmConnect.AuthSocket.Connect
#If REMOTE_CLOSE = 0 Then
    Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_CONECTANDO_SERVIDOR"), True, False)
#End If
    SessionOpened = False
    Auth_state = e_state.Idle
End Sub

Public Sub HandleOpenSession(ByVal BytesTotal As Long)

    frmDebug.add_text_tracebox "HandleOpenSession"
    
    Dim strData As String
    frmConnect.AuthSocket.PeekData strData, vbString, BytesTotal
    frmConnect.AuthSocket.GetData strData, vbString, 2
    frmConnect.AuthSocket.GetData strData, vbString, 2
    
    Dim encrypted_token() As Byte
    
    frmConnect.AuthSocket.GetData encrypted_token, 64
            
    decrypted_session_token = AO20CryptoSysWrapper.Decrypt(MapInfoEspeciales, cnvStringFromHexStr(cnvToHex(encrypted_token)))
    frmDebug.add_text_tracebox "Decrypted_session_token: " & decrypted_session_token
            
    public_key = mid(decrypted_session_token, 1, 16)
    
    frmDebug.add_text_tracebox "Public key:" & CStr(public_key)
    
    Str2ByteArr decrypted_session_token, public_key, 16
    
    SessionOpened = True
    encrypted_session_token = cnvStringFromHexStr(cnvToHex(encrypted_token))
End Sub

Public Sub HandlePCList(ByVal BytesTotal As Long)

    If BytesTotal < 4 Then
        frmDebug.add_text_tracebox "HandlePCList: Paquete incorrecto"
        Exit Sub
    End If
    
    Dim packet_size_byte() As Byte
    Dim PacketId() As Byte
    
    
    frmDebug.add_text_tracebox "HandlePCList"
    
    Dim strData As String
    frmConnect.AuthSocket.PeekData strData, vbString, BytesTotal
    Call DebugPrint("Bytes total: " & strData, 255, 255, 255, False)
    
    frmConnect.AuthSocket.GetData PacketId, vbByte, 2
    Call DebugPrint("Id: " & ByteArrayToHex(PacketId), 255, 255, 255, False)
    
    frmConnect.AuthSocket.GetData packet_size_byte, vbByte, 2
    
    Dim encrypted_list() As Byte
    Dim packet_size As Integer
    
    packet_size = MakeInt(packet_size_byte(1), packet_size_byte(0))
    frmConnect.AuthSocket.GetData encrypted_list, packet_size - 4
        
    Dim decrypted_list As String
     
    decrypted_list = AO20CryptoSysWrapper.Decrypt(ByteArrayToHex(public_key), cnvStringFromHexStr(cnvToHex(encrypted_list)))
#If REMOTE_CLOSE = 0 Then
    Call FillAccountData(decrypted_list)
#End If
    Call DebugPrint("Decrypted_list: " & decrypted_list, 255, 255, 255, False)
            
    Auth_state = e_state.AccountLogged
    
#If REMOTE_CLOSE = 1 Then
    LoginCharacter (CharacterRemote)
#End If

End Sub

Public Sub HandleAccountLogin(ByVal BytesTotal As Long)

    frmDebug.add_text_tracebox "HandleRequestAccountLogin"

    Dim Data() As Byte
    frmConnect.AuthSocket.PeekData Data, vbByte, BytesTotal
    frmConnect.AuthSocket.GetData Data, vbByte, 2
    If Data(0) = &HAF And Data(1) = &HA1 Then
        frmDebug.add_text_tracebox "LOGIN-OK"
        'Save the token which was used to authenticate
        authenticated_decrypted_session_token = decrypted_session_token
        Call DebugPrint(AO20CryptoSysWrapper.ByteArrayToHex(Data), 255, 255, 255)
        frmConnect.AuthSocket.GetData Data, vbByte, 2
        Auth_state = e_state.AccountLogged
        Call PCListRequest
    Else
       
        frmDebug.add_text_tracebox "LOGIN-ERROR"
        frmConnect.AuthSocket.GetData Data, vbByte, 4
#If REMOTE_CLOSE = 0 Then
        Select Case MakeInt(Data(3), Data(2))
            Case 1
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_USUARIO_INVALIDO"), False, False)
            Case 4
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_USUARIO_CONECTADO"), False, False)
                If Not FullLogout Then
                    Call SendAccountLoginRequest
                End If
            Case 5
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_CONTRASENA_INVALIDA"), False, False)
            Case 6
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_USUARIO_BANEADO"), False, False)
            Case 7
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_SERVIDOR_MAX_USUARIOS"), False, False)
            Case 9
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_CUENTA_NO_ACTIVADA"), False, False)
            Case 66
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_ACTIVO_PATRON"), False, False)
            Case Else
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_ACTUALIZAR_JUEGO"), False, False)

        End Select
    End If
#Else
        Select Case MakeInt(Data(3), Data(2))
            Case 1
                Call SaveStringInFile("MENSAJEBOX_USUARIO_INVALIDO", "remote_debug.txt")
            Case 4
                Call SaveStringInFile("MENSAJEBOX_USUARIO_CONECTADO", "remote_debug.txt")
            Case 5
                Call SaveStringInFile("MENSAJEBOX_CONTRASENA_INVALIDA", "remote_debug.txt")
            Case 6
                Call SaveStringInFile("MENSAJEBOX_USUARIO_BANEADO", "remote_debug.txt")
            Case 7
                Call SaveStringInFile("MENSAJEBOX_SERVIDOR_MAX_USUARIOS", "remote_debug.txt")
            Case 9
                Call SaveStringInFile("MENSAJEBOX_CUENTA_NO_ACTIVADA", "remote_debug.txt")
            Case 66
                Call SaveStringInFile("MENSAJEBOX_ACTIVO_PATRON", "remote_debug.txt")
            Case Else
                Call SaveStringInFile("MENSAJEBOX_ACTUALIZAR_JUEGO", "remote_debug.txt")
        End Select
        prgRun = False
    End If
#End If
End Sub

Public Sub GotoPasswordReset()
    frmPasswordReset.toggleTextboxs
End Sub

Public Sub HandleRequestForgotPassword(ByVal BytesTotal As Long)

    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    Call DebugPrint("HandleRequestForgotPassword", 255, 255, 255, True)
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    Dim Data() As Byte
    
    frmConnect.AuthSocket.PeekData Data, vbByte, BytesTotal
    frmConnect.AuthSocket.GetData Data, vbByte, 2
    
    If Data(0) = &H2 And Data(1) = &H14 Then
        Call DebugPrint("FORGOT-PASSWORD-OK", 0, 255, 0, True)
        frmConnect.AuthSocket.GetData Data, vbByte, 2
       Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_EMAIL_ENVIADO") & CuentaEmail & ".", False, False)
        Call GotoPasswordReset
        ModAuth.LoginOperation = e_operation.ResetPassword
        Auth_state = e_state.RequestResetPassword
    Else
       Call DebugPrint("ERROR", 255, 0, 0, True)
        frmConnect.AuthSocket.GetData Data, vbByte, 4
        Select Case MakeInt(Data(3), Data(2))
            Case 1
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_USUARIO_INVALIDO"), False, False)
            Case 3
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_INTENTAR_MAS_TARDE"), False, False)
            Case 9
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_CUENTA_NO_ACTIVADA"), False, False)
            Case 12
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_EMAIL_INVALIDO"), False, False)
            Case 23
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_CALMATE"), False, False)
            Case Else
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_ERROR_DESCONOCIDO") & AO20CryptoSysWrapper.ByteArrayToHex(Data), False, False)

        End Select
        
        Auth_state = e_state.Idle
        
    End If
    
    frmPasswordReset.lblSolicitandoCodigo.visible = False
    frmPasswordReset.cmdEnviar.visible = True
    
        
End Sub

Public Sub HandleRequestResetPassword(ByVal BytesTotal As Long)

    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    Call DebugPrint("HandleRequestResetPassword", 255, 255, 255, True)
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    Dim Data() As Byte
    
    frmConnect.AuthSocket.PeekData Data, vbByte, BytesTotal
    frmConnect.AuthSocket.GetData Data, vbByte, 2
    
    If Data(0) = &H20 And Data(1) = &H16 Then
        Call DebugPrint("RESET-PASSWORD-OK", 0, 255, 0, True)
        frmConnect.AuthSocket.GetData Data, vbByte, 2
        Call TextoAlAsistente("Contraseña recuperada con éxito.", False, False)
        Auth_state = e_state.Idle
        Unload frmPasswordReset
    Else
       Call DebugPrint("ERROR", 255, 0, 0, True)
        frmConnect.AuthSocket.GetData Data, vbByte, 4
        Select Case MakeInt(Data(3), Data(2))
            Case 1
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_CODIGO_INVALIDO"), False, False)
            Case 3
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_INTENTALO_MAS_TARDE"), False, False)
            Case 9
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_CUENTA_NO_ACTIVADA"), False, False)
            Case 12
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_EMAIL_INVALIDO"), False, False)
            Case 14
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_PASSWORD_CORTO"), False, False)
            Case 15
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_PASSWORD_LARGO"), False, False)
            Case 16
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_PASSWORD_INVALIDOS"), False, False)
            Case 21
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_CODIGO_INVALIDO"), False, False)
            Case 22
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_PASSWORD_RESET_HOST_INVALIDO"), False, False)
            Case 23
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_INTENTALO_MAS_TARDE"), False, False)
            Case 24
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_PASSWORD_CONTAINS_USERNAME"), False, False)
            Case 33
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_PASSWORD_UPPERCASE"), False, False)
            Case 34
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_PASSWORD_LOWERCASE"), False, False)
            Case 35
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_PASSWORD_NUMBER"), False, False)
            Case 36
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_CODIGO_RECUPERACION_ANTIGUO"), False, False)
            Case &H40
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_CODIGO_EXPIRADO"), False, False)
            Case Else
                Call TextoAlAsistente("Unknown error: " & AO20CryptoSysWrapper.ByteArrayToHex(Data), False, False)
        End Select
    End If
        
End Sub

Public Sub AccountValidated()
    Auth_state = e_state.Idle
    Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_CUENTA_VALIDADA"), False, False)
    frmNewAccount.visible = False
End Sub

Public Sub HandleValidateAccountRequest(ByVal BytesTotal As Long)

    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    Call DebugPrint("HandleValidateAccountRequest", 255, 255, 255, True)
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    Dim Data() As Byte
    
    frmConnect.AuthSocket.PeekData Data, vbByte, BytesTotal
    
    frmConnect.AuthSocket.GetData Data, vbByte, 2
    
    If Data(0) = &H77 And Data(1) = &H77 Then
        Call DebugPrint("VALIDATE-ACCOUNT-OK", 0, 255, 0, True)
        Call DebugPrint(AO20CryptoSysWrapper.ByteArrayToHex(Data), 255, 255, 255)
        frmConnect.AuthSocket.GetData Data, vbByte, 2
        Call AccountValidated
    Else
       Call DebugPrint("ERROR", 255, 0, 0, True)
        frmConnect.AuthSocket.GetData Data, vbByte, 4
        Select Case MakeInt(Data(3), Data(2))
            Case 1
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_USUARIO_INVALIDO"), False, False)
                frmNewAccount.visible = False
            Case 10
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_CODIGO_INVALIDO"), False, False)
            Case 11
                Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_CUENTA_ACTIVADA"), False, False)
                frmNewAccount.visible = False
            Case Else
                Call TextoAlAsistente("Unknown error: " & AO20CryptoSysWrapper.ByteArrayToHex(Data), False, False)
        End Select
    End If
        
End Sub

Function FileToString(strFileName As String) As String
  Open strFileName For Input As #1
    FileToString = StrConv(InputB(LOF(1), 1), vbUnicode)
  Close #1
End Function
Private Sub EraseCharFromPjList(ByVal nick As String)
    Dim i As Long, J As Long
    
    For i = 1 To CantidadDePersonajesEnCuenta
        If LCase(Pjs(i).nombre) = LCase(nick) Then
            'Desde esta posicion en adelante tengo que correrlos todos 1 pos para atras
            For J = i To (CantidadDePersonajesEnCuenta - 1)
                Pjs(J) = Pjs(J + 1)
            Next J
            Exit For
        End If
    Next i

    'Borro el último personaje
    Pjs(CantidadDePersonajesEnCuenta).nombre = ""
    Pjs(CantidadDePersonajesEnCuenta).Head = 0
    Pjs(CantidadDePersonajesEnCuenta).Clase = 0
    Pjs(CantidadDePersonajesEnCuenta).Body = 0
    Pjs(CantidadDePersonajesEnCuenta).Mapa = 0
    Pjs(CantidadDePersonajesEnCuenta).PosX = 0
    Pjs(CantidadDePersonajesEnCuenta).PosY = 0
    Pjs(CantidadDePersonajesEnCuenta).Nivel = 0
    Pjs(CantidadDePersonajesEnCuenta).Criminal = 0
    Pjs(CantidadDePersonajesEnCuenta).Casco = 0
    Pjs(CantidadDePersonajesEnCuenta).Escudo = 0
    Pjs(CantidadDePersonajesEnCuenta).Arma = 0
    Pjs(CantidadDePersonajesEnCuenta).ClanName = ""
    Pjs(CantidadDePersonajesEnCuenta).NameMapa = ""
    
    CantidadDePersonajesEnCuenta = CantidadDePersonajesEnCuenta - 1
    
End Sub
Private Sub FillAccountData(ByVal Data As String)
    On Error Resume Next
    Dim i As Long
    CantidadDePersonajesEnCuenta = 0
    For i = 1 To Len(Data)
        If mid(Data, i, 1) = "(" Then
            CantidadDePersonajesEnCuenta = CantidadDePersonajesEnCuenta + 1
        End If
    Next i

    Dim ii As Byte
     'name, head_id, class_id, body_id, pos_map, pos_x, pos_y, level, status, helmet_id, shield_id, weapon_id, guild_index, is_dead, is_sailing
    For ii = 1 To MAX_PERSONAJES_EN_CUENTA
        Pjs(ii).nombre = ""
        Pjs(ii).Head = 0 ' si is_sailing o muerto, cabeza en 0
        Pjs(ii).Clase = 0
        Pjs(ii).Body = 0
        Pjs(ii).Mapa = 0
        Pjs(ii).PosX = 0
        Pjs(ii).PosY = 0
        Pjs(ii).Nivel = 0
        Pjs(ii).Criminal = 0
        Pjs(ii).Casco = 0
        Pjs(ii).Escudo = 0
        Pjs(ii).Arma = 0
        Pjs(ii).ClanName = ""
        Pjs(ii).NameMapa = ""
    Next ii

    For ii = 1 To min(CantidadDePersonajesEnCuenta, MAX_PERSONAJES_EN_CUENTA)
        Dim character As String
        character = ReadField(ii, Data, Asc(")"))
        character = Replace(character, "(", "")
        character = Replace(character, "[", "")
        character = Replace(character, "]", "")
        character = Replace(character, "'", "")
        If mid(character, 1, 1) = "," Then
            character = mid(character, 2)
        End If
        Dim Name As String
        
        Name = ReadField(1, character, Asc(","))
        If mid(Name, 1, 1) = " " Then
            Name = Replace(Name, " ", "", 1, 1)
        End If
        Pjs(ii).nombre = Name
        Pjs(ii).Body = Val(ReadField(4, character, Asc(",")))
        Pjs(ii).Head = IIf(Pjs(ii).Body = 829 Or Pjs(ii).Body = 1269 Or Pjs(ii).Body = 1267 Or Pjs(ii).Body = 1265, 0, Val(ReadField(2, character, Asc(","))))
        Pjs(ii).Clase = Val(ReadField(3, character, Asc(",")))
        Pjs(ii).Mapa = Val(ReadField(5, character, Asc(",")))
        Pjs(ii).PosX = Val(ReadField(6, character, Asc(",")))
        Pjs(ii).PosY = Val(ReadField(7, character, Asc(",")))
        Pjs(ii).Nivel = Val(ReadField(8, character, Asc(",")))
        Pjs(ii).Criminal = Val(ReadField(9, character, Asc(",")))
        Pjs(ii).Casco = Val(ReadField(10, character, Asc(",")))
        Pjs(ii).Escudo = Val(ReadField(11, character, Asc(",")))
        Pjs(ii).Arma = Val(ReadField(12, character, Asc(",")))
        Pjs(ii).ClanName = "" ' "<" & "pepito" & ">"
       
        ' Pjs(ii).NameMapa = Pjs(ii).mapa
       ' Pjs(ii).NameMapa = NameMaps(Pjs(ii).Mapa).Name

    Next ii

    For i = 1 To min(CantidadDePersonajesEnCuenta, MAX_PERSONAJES_EN_CUENTA)

        Select Case Pjs(i).Criminal

            Case 0 'Criminal
                Call SetRGBA(Pjs(i).LetraColor, ColoresPJ(23).r, ColoresPJ(23).G, ColoresPJ(23).b)
                

            Case 1 'Ciudadano
                Call SetRGBA(Pjs(i).LetraColor, ColoresPJ(20).r, ColoresPJ(20).G, ColoresPJ(20).b)
                

            Case 2 'Caos
                Call SetRGBA(Pjs(i).LetraColor, ColoresPJ(24).r, ColoresPJ(24).G, ColoresPJ(24).b)
                

            Case 3 'Armada
                Call SetRGBA(Pjs(i).LetraColor, ColoresPJ(21).r, ColoresPJ(21).G, ColoresPJ(21).b)
                
            
            Case 4 'Concilio
                Call SetRGBA(Pjs(i).LetraColor, ColoresPJ(25).r, ColoresPJ(25).G, ColoresPJ(25).b)
                
                
            Case 5 'Consejo
                Call SetRGBA(Pjs(i).LetraColor, ColoresPJ(22).r, ColoresPJ(22).G, ColoresPJ(22).b)
                
           
            Case Else

        End Select
        Pjs(i).priv = 0
    Next i
    
    AlphaRenderCuenta = MAX_ALPHA_RENDER_CUENTA
   
    If CantidadDePersonajesEnCuenta > 0 Then
        PJSeleccionado = 1
        LastPJSeleccionado = 1
        
        If Pjs(1).Mapa <> 0 Then
            Call SwitchMap(Pjs(1).Mapa)
            RenderCuenta_PosX = Pjs(1).PosX
            RenderCuenta_PosY = Pjs(1).PosY
        End If
    End If
    Call LoadCharacterSelectionScreen
End Sub

Public Function estaInmovilizado(ByRef arr() As Byte) As String

    Dim a As String, b As String
    

    a = hashHexFromFile(App.path & "\Argentum.exe", 3)

    Dim i As Long
    
    For i = 1 To 4
        b = b + mid(a, arr(i), 1)
    Next i
    
    For i = 8 To 5 Step -1
        b = b + mid(a, arr(i), 1)
    Next i
    
    For i = 9 To 12
        b = b + mid(a, arr(i), 1)
    Next i
    
    For i = 16 To 13 Step -1
        b = b + mid(a, arr(i), 1)
    Next i
    
        
    #If DEBUGGING = 1 Then
        frmDebug.add_text_tracebox "Esta inmovilizado: " & b
    #End If
    estaInmovilizado = cnvHexStrFromString(b)
    
End Function


