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


Public SessionOpened As Boolean

Public Auth_state As e_state
Public LoginOperation As e_operation
Public public_key() As Byte
Public encrypted_session_token As String
Public decrypted_session_token As String
Public authenticated_decrypted_session_token As String
Public delete_char_validate_code As String


Public Sub AuthSocket_DataArrival(ByVal bytesTotal As Long)
    
    If Connected Then
        Exit Sub
    End If
    
    If Not SessionOpened Then
        Call HandleOpenSession(bytesTotal)
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
            Call HandleAccountLogin(bytesTotal)
        Case e_state.RequestCharList
            Call HandlePCList(bytesTotal)
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

    Debug.Print (str)
    
End Sub

Public Sub SendAccountLoginRequest()
    Dim username As String
    Dim password As String
    Dim len_encrypted_username As Integer
    Dim len_encrypted_password As Integer
    
    Dim login_request() As Byte
    Dim packet_size As Integer
    Dim offset_login_request As Long
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    Call DebugPrint("SendAccountLoginRequest", 255, 255, 255, True)
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    username = CuentaEmail
    password = CuentaPassword
    
    Dim encrypted_username() As Byte
    Dim encrypted_username_b64 As String
    
    Dim encrypted_password() As Byte
    Dim encrypted_password_b64 As String
    
    
    encrypted_username_b64 = AO20CryptoSysWrapper.Encrypt(cnvHexStrFromBytes(public_key), username)
    encrypted_password_b64 = AO20CryptoSysWrapper.Encrypt(cnvHexStrFromBytes(public_key), password)
    
    'Call DebugPrint("Username: " & encrypted_username_b64, 255, 255, 255)
    'Call DebugPrint("Password: " & encrypted_password_b64, 255, 255, 255)
    
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
    Dim username As String
    Dim password As String
    Dim len_encrypted_username As Integer
    Dim len_encrypted_password As Integer
    
    Dim login_request() As Byte
    Dim packet_size As Integer
    Dim offset_login_request As Long
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    Call DebugPrint("SendRequestVerificationCode", 255, 255, 255, True)
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    username = CuentaEmail
    password = CuentaEmail
    
    Dim encrypted_username() As Byte
    Dim encrypted_username_b64 As String
    
    Dim encrypted_password() As Byte
    Dim encrypted_password_b64 As String
    
    Debug.Assert Len(username) > 0 And Len(password) > 0
    
    encrypted_username_b64 = AO20CryptoSysWrapper.Encrypt(cnvHexStrFromBytes(public_key), username)
    encrypted_password_b64 = AO20CryptoSysWrapper.Encrypt(cnvHexStrFromBytes(public_key), password)
    
    Call Str2ByteArr(encrypted_username_b64, encrypted_username)
    Call Str2ByteArr(encrypted_password_b64, encrypted_password)
    
    Dim len_username As Integer
    Dim len_password As Integer
    
    len_username = Len(encrypted_username_b64)
    len_password = Len(encrypted_password_b64)
    
    ReDim login_request(1 To (2 + 2 + 2 + len_username + 2 + len_password))
    
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
        
    login_request(offset_login_request + 1) = hiByte(len_password)
    login_request(offset_login_request + 2) = LoByte(len_password)
    
    Call AO20CryptoSysWrapper.CopyBytes(encrypted_password, login_request, len_password, offset_login_request + 3)
    
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
        Call TextoAlAsistente("Código enviado correctamente a " & CuentaEmail & ".")
        
    Else
       Call DebugPrint("ERROR", 255, 0, 0, True)
        frmConnect.AuthSocket.GetData Data, vbByte, 4
        Select Case MakeInt(Data(3), Data(2))
            Case 1
                Call TextoAlAsistente("Account does not exist")
            Case 11
                Call TextoAlAsistente("Account has already been activated.")
            Case 12
                Call TextoAlAsistente("Account does not exist")
            Case Else
                Call TextoAlAsistente("No se ha podido conectar intente más tarde. Error: " & AO20CryptoSysWrapper.ByteArrayToHex(data))
        End Select
    End If
    
    Auth_state = e_state.Idle
        
End Sub

Public Sub SendRequestForgotPassword()
    Dim username As String
    Dim len_encrypted_username As Integer
    
    Dim login_request() As Byte
    Dim packet_size As Integer
    Dim offset_login_request As Long
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    Call DebugPrint("SendRequestForgotPassword", 255, 255, 255, True)
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    username = CuentaEmail
    
    Dim encrypted_username() As Byte
    Dim encrypted_username_b64 As String
    
    encrypted_username_b64 = AO20CryptoSysWrapper.Encrypt(cnvHexStrFromBytes(public_key), username)
    
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
    Dim username As String
    Dim validate_code As String
    Dim len_encrypted_username As Integer
    Dim len_encrypted_validate_code As Integer
    
    Dim login_request() As Byte
    Dim packet_size As Integer
    Dim offset_login_request As Long
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    Call DebugPrint("SendValidateAccount", 255, 255, 255, True)
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    username = frmNewAccount.txtValidateMail.Text
    validate_code = frmNewAccount.txtCodigo.Text
    
    Dim encrypted_username() As Byte
    Dim encrypted_username_b64 As String
    
    Dim encrypted_validate_code() As Byte
    Dim encrypted_validate_code_b64 As String
    
    
    encrypted_username_b64 = AO20CryptoSysWrapper.Encrypt(cnvHexStrFromBytes(public_key), username)
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
    Dim username As String
    Dim password As String
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
    username = frmPasswordReset.txtEmail.Text
    password = frmPasswordReset.txtPassword.Text
    validate_code = Trim(frmPasswordReset.txtCodigo.Text)
    
    Debug.Assert Len(validate_code) > 0 And Len(username) > 0 And Len(password) > 0
    
    Dim encrypted_username() As Byte
    Dim encrypted_username_b64 As String
    
    
    Dim encrypted_password() As Byte
    Dim encrypted_password_b64 As String
    
    Dim encrypted_validate_code() As Byte
    Dim encrypted_validate_code_b64 As String
    
    
    encrypted_username_b64 = AO20CryptoSysWrapper.Encrypt(cnvHexStrFromBytes(public_key), username)
    encrypted_password_b64 = AO20CryptoSysWrapper.Encrypt(cnvHexStrFromBytes(public_key), password)
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
Public Sub HandleLogOutRequest(ByVal bytesTotal As Long)
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    Call DebugPrint("HandleLogOutRequest", 255, 255, 255, True)
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    
    Dim data() As Byte
    
    frmConnect.AuthSocket.PeekData data, vbByte, bytesTotal
    
    frmConnect.AuthSocket.GetData data, vbByte, 2
    
    If data(0) = &H20 And data(1) = &H22 Then
        Call DebugPrint("LOGOUT_OKAY", 0, 255, 0, True)
        
        frmConnect.AuthSocket.GetData data, vbByte, 2
        
        Auth_state = e_state.Idle
    Else
       Call DebugPrint("ERROR", 255, 0, 0, True)
        frmConnect.AuthSocket.GetData data, vbByte, 4
        Select Case MakeInt(data(3), data(2))
            Case 41
                Call DebugPrint("Not logged yet.", 255, 255, 0)
        End Select
    End If
    Auth_state = e_state.Idle
End Sub
Public Sub SendRequestTransferCharacter()
    Dim json As String
    Dim len_encrypted_username As Integer
    Dim transfer_request() As Byte
    Dim packet_size As Integer
    
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    Call DebugPrint("SendRequestTransferCharacter", 255, 255, 255, True)
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    Dim old_owner_str As String
    old_owner_str = CuentaEmail
    
    json = ""
    json = "{ "
    json = json & "  ""currentOwner"": """ & old_owner_str & """ , "
    json = json & "  ""pc"": """ & TransferCharname & """ , "
    json = json & "  ""token"": """ & authenticated_decrypted_session_token & """ , "
    json = json & "  ""newOwner"": """ & TransferCharNewOwner & """"
    json = json & " }"

    Dim encrypted_json() As Byte
    Dim encrypted_json_b64 As String
    
    encrypted_json_b64 = AO20CryptoSysWrapper.Encrypt(cnvHexStrFromBytes(public_key), json)
        
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

    Dim json As String
    Dim len_encrypted_username As Integer
    Dim login_request() As Byte
    Dim packet_size As Integer
    
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    Call DebugPrint("SendSignUpRequest", 255, 255, 255, True)
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    
    
    json = ""
    
    json = "{ ""language"": ""english"", ""password"": """ & frmNewAccount.txtPassword & """, "
    json = json & """passwordrecovery"": [{""secretanswer1"": ""Satanas"","
    json = json & """secretanswer2"": ""Rojo"", "
    json = json & """secretquestion1"": ""Cual es el nombre de mi primer mascota?"","
    json = json & """secretquestion2"": ""Cual es mi color favorito?""}],"
    
    json = json & """personal"":[{"
    json = json & """dob"": ""23-12-1990"","
    json = json & """email"": """ & frmNewAccount.txtUsername & ""","
    json = json & """firstname"": """ & frmNewAccount.txtName & ""","
    json = json & """lastname"": """ & frmNewAccount.txtSurname & ""","
    json = json & """mobile"": """ & frmNewAccount.txtSurname & ""","
    json = json & """pob"": """ & frmNewAccount.txtSurname & """}],"
    json = json & """username"": """ & frmNewAccount.txtUsername & """}"
    
    
    Dim encrypted_json() As Byte
    Dim encrypted_json_b64 As String
    
    encrypted_json_b64 = AO20CryptoSysWrapper.Encrypt(cnvHexStrFromBytes(public_key), json)
        
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
    
    Dim data() As Byte
    
    frmConnect.AuthSocket.PeekData data, vbByte, BytesTotal
    frmConnect.AuthSocket.GetData data, vbByte, 2
    
    'We return to the LOGIN screen so that TextoAlAsistente works
    g_game_state.state = e_state_connect_screen
    FrmLogear.Show , frmConnect

    If data(0) = &H20 And data(1) = &H26 Then
        Call DebugPrint("TRANSFER_CHARACTER_OKAY", 0, 255, 0, True)
        Call TextoAlAsistente("TRANSFER_CHARACTER_OKAY")
        frmConnect.AuthSocket.GetData data, vbByte, 2
        Auth_state = e_state.Idle
    Else
        Call DebugPrint("TRANSFER CHAR ERROR", 255, 0, 0, True)
        frmConnect.AuthSocket.GetData data, vbByte, 4
        Select Case MakeInt(data(3), data(2))
            Case 1
                Call TextoAlAsistente("Invalid account")
            Case 3
                Call TextoAlAsistente("Database error.")
            Case 12
                Call TextoAlAsistente("Email is not valid.")
            Case 51
                Call TextoAlAsistente("You are not the owner of the character.")
            Case 52
                Call TextoAlAsistente("Invalid request")
            Case 54
                Call TextoAlAsistente("Newowner does not exist")
            Case 55
                Call TextoAlAsistente("Not a patron, sorry")
            Case 57
                Call TextoAlAsistente("You do not have enough credits")
            Case Else
                Call TextoAlAsistente("Unknown error.")
        End Select
    End If
    Auth_state = e_state.Idle

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
        Call TextoAlAsistente("Cuenta creada correctamente.")
        frmConnect.AuthSocket.GetData Data, vbByte, 2
        Call frmNewAccount.showValidateAccountControls
        frmNewAccount.txtValidateMail.Text = frmNewAccount.txtUsername
        Auth_state = e_state.Idle
    Else
       Call DebugPrint("ERROR", 255, 0, 0, True)
        frmConnect.AuthSocket.GetData Data, vbByte, 4
        Select Case MakeInt(Data(3), Data(2))
            Case 0
                Call TextoAlAsistente("Username already exist.")
            Case 9
                Call TextoAlAsistente("The server could not send the activation email.")
            Case 12
                Call TextoAlAsistente("Email is not valid.")
            Case 14
                Call TextoAlAsistente("Password is too short.")
            Case 15
                Call TextoAlAsistente("Password is too long.")
            Case 16
                Call TextoAlAsistente("Password contains invalid characters.")
            Case 18
                Call TextoAlAsistente("Username is too short.")
            Case 19
                Call TextoAlAsistente("Username is too long.")
            Case 20
                Call TextoAlAsistente("Username contains invalid characters.")
            Case 24
                Call TextoAlAsistente("password must not contain username.")
            Case 32
                Call TextoAlAsistente("Username can not start with a number.")
            Case 33
                Call TextoAlAsistente("The password has no uppercase letters.")
            Case 34
                Call TextoAlAsistente("The password has no lowercase letters.")
            Case 35
                Call TextoAlAsistente("The password must contain at least than two numbers.")
            Case Else
                Call TextoAlAsistente("Unknown error.")
        End Select
    End If
    Auth_state = e_state.Idle
End Sub

Public Sub SendDeleteCharRequest()

    Dim json As String
    Dim len_encrypted_username As Integer
    Dim delete_char_request() As Byte
    Dim packet_size As Integer
    
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    Call DebugPrint("SendDeleteChar", 255, 255, 255, True)
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
        
    json = "{""username"": """ & CuentaEmail & """, ""pc"":""" & DeleteUser & """}"
    
    Dim encrypted_json() As Byte
    Dim encrypted_json_b64 As String
    
    encrypted_json_b64 = AO20CryptoSysWrapper.Encrypt(cnvHexStrFromBytes(public_key), json)
        
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
        MsgBox ("Se ha enviado un código de verificación al mail proporcionado.")
        frmConnect.AuthSocket.GetData Data, vbByte, 2
        Auth_state = e_state.Idle
    Else
        Call DebugPrint("DELETE_PC_REQUEST_ERROR", 255, 0, 0, True)
        frmDeleteChar.Hide
        frmConnect.AuthSocket.GetData Data, vbByte, 4
        Select Case MakeInt(Data(3), Data(2))
            Case 1
                Call MsgBox("Invalid account.", vbOKOnly)
            Case 3
                Call MsgBox("Database error.", vbOKOnly)
            Case 51
                Call MsgBox("You are not the character owner.", vbOKOnly)
            Case 65
                Call MsgBox("Cannot delete character listed in MAO.", vbOKOnly)
            Case Else
                Call MsgBox("Unknown error", vbOKOnly)
        End Select
    End If
    Auth_state = e_state.Idle
End Sub

Public Sub SendConfirmDeleteChar()
    Dim username As String
    Dim validate_code As String
    Dim len_encrypted_username As Integer
    Dim len_encrypted_validate_code As Integer
    
    Dim login_request() As Byte
    Dim packet_size As Integer
    Dim offset_login_request As Long
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    Call DebugPrint("SendConfirmDeleteChar", 255, 255, 255, True)
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    username = DeleteUser
    validate_code = delete_char_validate_code
    
    Dim encrypted_username() As Byte
    Dim encrypted_username_b64 As String
    
    Dim encrypted_validate_code() As Byte
    Dim encrypted_validate_code_b64 As String
    
    
    encrypted_username_b64 = AO20CryptoSysWrapper.Encrypt(cnvHexStrFromBytes(public_key), username)
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
        Call MsgBox("Personaje borrado correctamente.", vbOKOnly)
        Call EraseCharFromPjList(DeleteUser)
    Else
       Call DebugPrint("ERROR", 255, 0, 0, True)
        frmConnect.AuthSocket.GetData Data, vbByte, 4
        Select Case MakeInt(Data(3), Data(2))
            Case 1
                Call MsgBox("Invalid character name.")
            Case 3
                Call MsgBox("Database error.")
            Case 25
                Call MsgBox("Invalid Code.")
                frmDeleteChar.Show , frmConnect
            Case Else
                Call MsgBox("Unknown error: " & AO20CryptoSysWrapper.ByteArrayToHex(Data))
        End Select
    End If
        
End Sub

Public Sub PCListRequest()
    Dim username As String
    Dim len_encrypted_username As Integer
    
    Dim packet_request() As Byte
    Dim charList_request() As Byte
    Dim offset_login_request As Long
    Dim packet_size As Integer
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    Call DebugPrint("PCListRequest", 255, 255, 255, True)
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    username = CuentaEmail
    
    Dim encrypted_username() As Byte
    Dim encrypted_username_b64 As String
    
    
    encrypted_username_b64 = AO20CryptoSysWrapper.Encrypt(cnvHexStrFromBytes(public_key), username)
    
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
    Debug.Print "Servidor de Login " & IPdelServidorLogin; ":" & PuertoDelServidorLogin
    frmConnect.AuthSocket.Connect
    Call TextoAlAsistente("Conectando al servidor. Aguarde un momento.")
    SessionOpened = False
    Auth_state = e_state.Idle
End Sub

Public Sub HandleOpenSession(ByVal bytesTotal As Long)

    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    Call DebugPrint("HandleOpenSession", 255, 255, 255, True)
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    
    Dim strData As String
    frmConnect.AuthSocket.PeekData strData, vbString, bytesTotal
    
    Call DebugPrint("Bytes total: " & strData, 255, 255, 255, False)
    
    frmConnect.AuthSocket.GetData strData, vbString, 2
    Call DebugPrint("Id: " & strData, 255, 255, 255, False)
    
    frmConnect.AuthSocket.GetData strData, vbString, 2
    
    Dim encrypted_token() As Byte
    
    frmConnect.AuthSocket.GetData encrypted_token, 64
            
    decrypted_session_token = AO20CryptoSysWrapper.DECRYPT(MapInfoEspeciales, cnvStringFromHexStr(cnvToHex(encrypted_token)))
    Call DebugPrint("Decripted_session_token: " & decrypted_session_token, 255, 255, 255, False)
        
    public_key = mid(decrypted_session_token, 1, 16)
    
    Call DebugPrint("Public key:" & CStr(public_key), 255, 255, 255, False)
    
    Str2ByteArr decrypted_session_token, public_key, 16
    
    SessionOpened = True
    encrypted_session_token = cnvStringFromHexStr(cnvToHex(encrypted_token))
End Sub

Public Sub HandlePCList(ByVal bytesTotal As Long)

    If bytesTotal < 4 Then
        Call DebugPrint("Paquete incorrecto", 255, 255, 255, True)
        Exit Sub
    End If
    
    Dim packet_size_byte() As Byte
    Dim PacketId() As Byte
    
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    Call DebugPrint("HandlePCList", 255, 255, 255, True)
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    Dim strData As String
    frmConnect.AuthSocket.PeekData strData, vbString, bytesTotal
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
    Call FillAccountData(decrypted_list)
    Call DebugPrint("Decrypted_list: " & decrypted_list, 255, 255, 255, False)
            
    Auth_state = e_state.AccountLogged
End Sub
Public Sub HandleAccountLogin(ByVal bytesTotal As Long)

    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    Call DebugPrint("HandleRequestAccountLogin", 255, 255, 255, True)
    Call DebugPrint("------------------------------------", 0, 255, 0, True)
    Dim data() As Byte
    
    frmConnect.AuthSocket.PeekData data, vbByte, bytesTotal
    
    frmConnect.AuthSocket.GetData data, vbByte, 2
    
    If data(0) = &HAF And data(1) = &HA1 Then
        Call DebugPrint("LOGIN-OK", 0, 255, 0, True)
        'Save the token which was used to authenticate
        authenticated_decrypted_session_token = decrypted_session_token
        Call DebugPrint(AO20CryptoSysWrapper.ByteArrayToHex(data), 255, 255, 255)
        frmConnect.AuthSocket.GetData data, vbByte, 2
        
        Auth_state = e_state.AccountLogged
        Call PCListRequest
    Else
       Call DebugPrint("ERROR", 255, 0, 0, True)
        frmConnect.AuthSocket.GetData data, vbByte, 4
        Select Case MakeInt(data(3), data(2))
            Case 1
                Call TextoAlAsistente("Invalid Username")
            Case 4
                Call TextoAlAsistente("Username is already logged.")
                If Not FullLogout Then
                    Call SendAccountLoginRequest
                End If
            Case 5
                Call TextoAlAsistente("Invalid Password.")
            Case 6
                Call TextoAlAsistente("Username has been banned.")
            Case 7
                Call TextoAlAsistente("Ther server has reached the max. number of users.")
            Case 9
                Call TextoAlAsistente("The account has not been activated.")
            Case Else
                'Call TextoAlAsistente("Unknown error: " & AO20CryptoSysWrapper.ByteArrayToHex(Data))
                Call TextoAlAsistente("No se ha podido conectar intente más tarde. Error: " & AO20CryptoSysWrapper.ByteArrayToHex(data))
        End Select
    End If
        
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
        Call TextoAlAsistente("Se ha enviado un email a " & CuentaEmail & ".")
        frmPasswordReset.toggleTextboxs
        
        ModAuth.LoginOperation = e_operation.ResetPassword
        Auth_state = e_state.RequestResetPassword
    
        
    Else
       Call DebugPrint("ERROR", 255, 0, 0, True)
        frmConnect.AuthSocket.GetData Data, vbByte, 4
        Select Case MakeInt(Data(3), Data(2))
            Case 1
                Call TextoAlAsistente("Invalid Username")
            Case 3
                Call TextoAlAsistente("Try again later.")
            Case 9
                Call TextoAlAsistente("The account has not been activated.")
            Case 12
                Call TextoAlAsistente("Invalid Email.")
            Case 23
                Call TextoAlAsistente("Calmate.")
            Case Else
                Call TextoAlAsistente("Unknown error: " & AO20CryptoSysWrapper.ByteArrayToHex(Data))
        End Select
        
        Auth_state = e_state.Idle
        
    End If
    
    frmPasswordReset.lblSolicitandoCodigo.Visible = False
    frmPasswordReset.cmdEnviar.Visible = True
    
        
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
        Call TextoAlAsistente("Contraseña recuperada con éxito.")
        Auth_state = e_state.Idle
        Unload frmPasswordReset
    Else
       Call DebugPrint("ERROR", 255, 0, 0, True)
        frmConnect.AuthSocket.GetData Data, vbByte, 4
        Select Case MakeInt(Data(3), Data(2))
            Case 1
                Call TextoAlAsistente("Invalid code")
            Case 3
                Call TextoAlAsistente("Try again later")
            Case 9
                Call TextoAlAsistente("The account has not been activated.")
            Case 12
                Call TextoAlAsistente("Invalid Email.")
            Case 14
                Call TextoAlAsistente("The password is too short.")
            Case 15
                Call TextoAlAsistente("The password is too long.")
            Case 16
                Call TextoAlAsistente("The password contains invalid characters.")
            Case 21
                Call TextoAlAsistente("Invalid code")
            Case 22
                Call TextoAlAsistente("Invalid password reset host")
            Case 23
                Call TextoAlAsistente("Try again later")
            Case 24
                Call TextoAlAsistente("The password cant not contain username.")
            Case 33
                Call TextoAlAsistente("The password must have at least one upper case letter.")
            Case 34
                Call TextoAlAsistente("The password must have at least one lower case letter.")
            Case 35
                Call TextoAlAsistente("The password must have at least one number.")
            Case 36
                Call TextoAlAsistente("The recovery code is too old.")
            Case &H40
                Call TextoAlAsistente("The code has expired, codes are valid for 10 mins, please request a new one.")
            Case Else
                Call TextoAlAsistente("Unknown error: " & AO20CryptoSysWrapper.ByteArrayToHex(Data))
        End Select
    End If
        
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
        Auth_state = e_state.Idle
        Call TextoAlAsistente("Cuenta validada exitosamente.")
        frmNewAccount.Visible = False
    Else
       Call DebugPrint("ERROR", 255, 0, 0, True)
        frmConnect.AuthSocket.GetData Data, vbByte, 4
        Select Case MakeInt(Data(3), Data(2))
            Case 1
                Call TextoAlAsistente("Invalid Username.")
                frmNewAccount.Visible = False
            Case 10
                Call TextoAlAsistente("Invalid Code.")
            Case 11
                Call TextoAlAsistente("Account has already been activated.")
                frmNewAccount.Visible = False
            Case Else
                Call TextoAlAsistente("Unknown error: " & AO20CryptoSysWrapper.ByteArrayToHex(Data))
        End Select
    End If
        
End Sub

Function FileToString(strFileName As String) As String
  Open strFileName For Input As #1
    FileToString = StrConv(InputB(LOF(1), 1), vbUnicode)
  Close #1
End Function
Private Sub EraseCharFromPjList(ByVal nick As String)
    Dim i As Long, j As Long
    
    For i = 1 To CantidadDePersonajesEnCuenta
        If LCase(Pjs(i).nombre) = LCase(nick) Then
            'Desde esta posicion en adelante tengo que correrlos todos 1 pos para atras
            For j = i To (CantidadDePersonajesEnCuenta - 1)
                Pjs(j) = Pjs(j + 1)
            Next j
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
    Pjs(CantidadDePersonajesEnCuenta).nivel = 0
    Pjs(CantidadDePersonajesEnCuenta).Criminal = 0
    Pjs(CantidadDePersonajesEnCuenta).Casco = 0
    Pjs(CantidadDePersonajesEnCuenta).Escudo = 0
    Pjs(CantidadDePersonajesEnCuenta).Arma = 0
    Pjs(CantidadDePersonajesEnCuenta).ClanName = ""
    Pjs(CantidadDePersonajesEnCuenta).NameMapa = ""
    
    CantidadDePersonajesEnCuenta = CantidadDePersonajesEnCuenta - 1
    
End Sub
Private Sub FillAccountData(ByVal data As String)
    On Error Resume Next
    Dim i As Long
    CantidadDePersonajesEnCuenta = 0
    For i = 1 To Len(data)
        If mid(data, i, 1) = "(" Then
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
        Pjs(ii).posX = 0
        Pjs(ii).posY = 0
        Pjs(ii).nivel = 0
        Pjs(ii).Criminal = 0
        Pjs(ii).Casco = 0
        Pjs(ii).Escudo = 0
        Pjs(ii).Arma = 0
        Pjs(ii).ClanName = ""
        Pjs(ii).NameMapa = ""
    Next ii

    For ii = 1 To min(CantidadDePersonajesEnCuenta, MAX_PERSONAJES_EN_CUENTA)
        Dim character As String
        character = ReadField(ii, data, Asc(")"))
        character = Replace(character, "(", "")
        character = Replace(character, "[", "")
        character = Replace(character, "]", "")
        character = Replace(character, "'", "")
        If mid(character, 1, 1) = "," Then
            character = mid(character, 2)
        End If
        Dim name As String
        
        name = ReadField(1, character, Asc(","))
        If mid(name, 1, 1) = " " Then
            name = Replace(name, " ", "", 1, 1)
        End If
        Pjs(ii).nombre = name
        Pjs(ii).Body = Val(ReadField(4, character, Asc(",")))
        Pjs(ii).Head = IIf(Pjs(ii).Body = 829 Or Pjs(ii).Body = 1269 Or Pjs(ii).Body = 1267 Or Pjs(ii).Body = 1265, 0, Val(ReadField(2, character, Asc(","))))
        Pjs(ii).Clase = Val(ReadField(3, character, Asc(",")))
        Pjs(ii).Mapa = Val(ReadField(5, character, Asc(",")))
        Pjs(ii).posX = Val(ReadField(6, character, Asc(",")))
        Pjs(ii).posY = Val(ReadField(7, character, Asc(",")))
        Pjs(ii).nivel = Val(ReadField(8, character, Asc(",")))
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
                Call SetRGBA(Pjs(i).LetraColor, ColoresPJ(23).r, ColoresPJ(23).G, ColoresPJ(23).B)
                

            Case 1 'Ciudadano
                Call SetRGBA(Pjs(i).LetraColor, ColoresPJ(20).r, ColoresPJ(20).G, ColoresPJ(20).B)
                

            Case 2 'Caos
                Call SetRGBA(Pjs(i).LetraColor, ColoresPJ(24).r, ColoresPJ(24).G, ColoresPJ(24).B)
                

            Case 3 'Armada
                Call SetRGBA(Pjs(i).LetraColor, ColoresPJ(21).r, ColoresPJ(21).G, ColoresPJ(21).B)
                
            
            Case 4 'Concilio
                Call SetRGBA(Pjs(i).LetraColor, ColoresPJ(25).r, ColoresPJ(25).G, ColoresPJ(25).B)
                
                
            Case 5 'Consejo
                Call SetRGBA(Pjs(i).LetraColor, ColoresPJ(22).r, ColoresPJ(22).G, ColoresPJ(22).B)
                
           
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
            RenderCuenta_PosX = Pjs(1).posX
            RenderCuenta_PosY = Pjs(1).posY
        End If
    End If
    
    Call mostrarcuenta
    

End Sub

Public Sub mostrarcuenta()
    AlphaNiebla = 30
    frmConnect.Visible = True

    g_game_state.state = e_state_account_screen

    
    SugerenciaAMostrar = RandomNumber(1, NumSug)
        

    Call Sound.Sound_Play(192)
    
    Call Sound.Sound_Stop(SND_LLUVIAIN)
      
    Call Graficos_Particulas.Particle_Group_Remove_All
    Call Graficos_Particulas.Engine_Select_Particle_Set(203)
    ParticleLluviaDorada = Graficos_Particulas.General_Particle_Create(208, -1, -1)
                
    If FrmLogear.Visible Then
        Unload FrmLogear

        'Unload frmConnect
    End If
    
    If frmMain.Visible Then
        '  frmMain.Visible = False
        
        UserParalizado = False
        UserInmovilizado = False
        UserStopped = False
        
        InvasionActual = 0
        frmMain.Evento.Enabled = False
     
        'BUG CLONES
        Dim i As Integer

        For i = 1 To LastChar
            Call EraseChar(i)
        Next i
        
        frmMain.personaje(1).Visible = False
        frmMain.personaje(2).Visible = False
        frmMain.personaje(3).Visible = False
        frmMain.personaje(4).Visible = False
        frmMain.personaje(5).Visible = False

    End If
End Sub

Public Function estaInmovilizado(ByRef arr() As Byte) As String

    Dim a As String, b As String
    

    a = hashHexFromFile(App.Path & "\Argentum.exe", 3)

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
        Debug.Print "Esta inmovilizado: " & b
    #End If
    estaInmovilizado = cnvHexStrFromString(b)
    
End Function


