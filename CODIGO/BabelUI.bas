Attribute VB_Name = "BabelUI"
Option Explicit

Const CreateCharMap = 782
Const CreateCharMapX = 25
Const CreateCharMapY = 35
Private Type LOGINDATA
    User As Long
    UserLen As Long
    password As Long
    PasswordLen As Long
    storeCredentials As Long
End Type

Private Type NewAccountData
    User As Long
    UserLen As Long
    password As Long
    PasswordLen As Long
    Name As Long
    NameLen As Long
    Surname As Long
    SurnameLen As Long
End Type

Private Type NEWCHARACTERDATA
    name As Long
    NameLen As Long
    Gender As Long
    Race As Long
    Class As Long
    Head As Long
    City As Long
End Type

Private Type SINGLESTRINGPARAM
    Ptr As Long
    Len As Long
End Type

Private Type DOUBLESTRINGPARAM
    FirstPtr As Long
    FirstLen As Long
    SecondPtr As Long
    SecondLen As Long
End Type

Private Type TRIPLESTRINGPARAM
    FirstPtr As Long
    FirstLen As Long
    SecondPtr As Long
    SecondLen As Long
    ThirdPtr As Long
    ThirdLen As Long
End Type

Private Type BABELSETTINGS
    Width As Long
    Height As Long
    Compresed As Long
    EnableDebug As Long
End Type

Private ServerEnvironment As String

Public Declare Function InitializeBabel Lib "BabelUI.dll" (ByRef Settings As BABELSETTINGS) As Boolean
Public Declare Function GetBebelImageBuffer Lib "BabelUI.dll" Alias "GetImageBuffer" (ByRef Buffer As Byte, ByVal size As Long) As Boolean
Public Declare Sub BabelSendMouseEvent Lib "BabelUI.dll" Alias "SendMouseEvent" (ByVal PosX As Long, ByVal PosY As Long, ByVal EvtType As Long, ByVal button As Long)
Public Declare Sub BabelSendKeyEvent Lib "BabelUI.dll" Alias "SendKeyEvent" (ByVal KeyCode As Integer, ByVal Shift As Boolean, ByVal EvtType As Long, ByVal CapsState As Boolean, ByVal Inspector As Boolean)
Public Declare Function NextPowerOf2 Lib "BabelUI.dll" (ByVal original As Long) As Long
Public Declare Sub RegisterCallbacks Lib "BabelUI.dll" (ByVal Login As Long, ByVal CloseClient As Long, ByVal CreateAccount As Long, ByVal SetHost As Long, ByVal ValidateAccountr As Long, ByVal ResendCode As Long, ByVal RequestPasswordReset As Long, _
                                                        ByVal RequestNewPassord As Long, ByVal SelectCharacter As Long, ByVal LoginCharacter As Long, ByVal ReturnToLogin As Long, ByVal CreateCharacter As Long, ByVal RequestCharDelete As Long, _
                                                        ByVal ConfirmCharDelete As Long, ByVal TransferCharacter As Long)
Public Declare Sub SendErrorMessage Lib "BabelUI.dll" (ByVal message As String, ByVal localize As Long, ByVal Action As Long)
Public Declare Sub SetActiveScreen Lib "BabelUI.dll" (ByVal screenName As String)
Public Declare Sub SetLoadingMessage Lib "BabelUI.dll" (ByVal message As String, ByVal localize As Long)
Public Declare Sub LoginCharacterListPrepare Lib "BabelUI.dll" (ByVal CharacterCount As Long)
Public Declare Sub LoginAddCharacter Lib "BabelUI.dll" (ByVal name As String, ByVal Head As Long, ByVal Body As Long, ByVal helm As Long, ByVal shield As Long, ByVal weapon As Long, ByVal level As Long, ByVal status As Long, ByVal Index As Long)
Public Declare Sub LoginSendCharacters Lib "BabelUI.dll" ()
Public Declare Sub RequestDeleteCode Lib "BabelUI.dll" ()
Public Declare Sub RemoveCharacterFromList Lib "BabelUI.dll" (ByVal Index As Long)
Public Declare Function GetTelemetry Lib "BabelUI.dll" (ByRef code As Byte, ByRef DataBuff As Byte, ByVal BuffSize As Long) As Long

'debug info
Public Declare Function CreateDebugWindow Lib "BabelUI.dll" (ByVal Width As Long, ByVal Height As Long) As Boolean
Public Declare Function GetDebugImageBuffer Lib "BabelUI.dll" (ByRef Buffer As Byte, ByVal size As Long) As Boolean
Public Declare Sub SendDebugMouseEvent Lib "BabelUI.dll" (ByVal PosX As Long, ByVal PosY As Long, ByVal EvtType As Long, ByVal button As Long)


Public Enum MouseEvent
    kType_MouseMoved = 0
    kType_MouseDown = 1
    kType_MouseUp = 2
End Enum

Public Enum MouseButton
    kButton_None = 0
    kButton_Left = 1
    kButton_Middle = 2
    kButton_Right = 3
End Enum

Public Enum KeyEventType
    kType_KeyDown = 0
    kType_KeyUp = 1
    kType_RawKeyDown = 2
    kType_Char = 3
End Enum

Public Type t_UITexture
    Texture As Direct3DTexture8
    ImageBuffer() As Byte
    Width As Long
    Height As Long
    TextureWidth As Long
    TextureHeight As Long
    pixelSize As Long
End Type

Public UITexture As t_UITexture
Public DebugUITexture As t_UITexture
Public BabelInitialized As Boolean
Public DebugInitialized As Boolean
Public GetRemoteError As Boolean
Public UseBabelUI As Boolean

Public Function ConvertMouseButton(ByVal button As Integer) As MouseButton
    Select Case button
        Case vbLeftButton
            ConvertMouseButton = kButton_Left
        Case vbRightButton
            ConvertMouseButton = kButton_Right
        Case vbMiddleButton
            ConvertMouseButton = kButton_Middle
        Case Else
            ConvertMouseButton = kButton_None
    End Select
End Function

Public Function CheckAndSetBabelUIUsage() As Boolean
On Error GoTo CheckAndSetBabelUIUsage_Err
100    UseBabelUI = Val(GetSetting("OPCIONES", "UseExperimentalUI"))
102    LogError ("Initilalize UI: " & UseBabelUI)
104    CheckAndSetBabelUIUsage = UseBabelUI
    Exit Function
CheckAndSetBabelUIUsage_Err:
    CheckAndSetBabelUIUsage = False
    Call RegistrarError(Err.Number, Err.Description, "BabelUI.CheckAndSetBabelUIUsage", Erl)
End Function

Public Function GetMainHwdn() As String
    If UseBabelUI Then
        GetMainHwdn = frmBabelLogin.hwnd
    Else
        GetMainHwdn = frmConnect.hwnd
    End If
End Function

Public Sub InitializeUI(ByVal Width As Long, ByVal Height As Long, ByVal pixelSize As Long)
On Error GoTo InitializeUI_Err
    Debug.Assert Width > 0 And Height > 0 And pixelSize > 0
    If BabelInitialized Then Exit Sub
102 UITexture.Height = Height
104 UITexture.Width = Width
    Dim Settings As BABELSETTINGS
    Settings.Height = Height
    Settings.Width = Width
#If Compresion = 1 Then
    Settings.Compresed = 1
#End If
#If DEBUGGING = 1 Or Developer = 1 Then
    Settings.EnableDebug = 1
#End If
106  If InitializeBabel(Settings) Then
108     UITexture.TextureHeight = NextPowerOf2(Height)
110     UITexture.TextureWidth = NextPowerOf2(Width)
112     ReDim UITexture.ImageBuffer(UITexture.Height * UITexture.Width * pixelSize)
114     UITexture.pixelSize = pixelSize
116     Set UITexture.Texture = SurfaceDB.CreateTexture(UITexture.TextureWidth, UITexture.TextureHeight)
118     BabelInitialized = True
        Call RegisterCallbacks(AddressOf LoginCB, AddressOf CloseClientCB, AddressOf BabelUI.CreateAccount, AddressOf SetHostCB, AddressOf ValidateCodeCB, AddressOf ResendValidationCodeCB, AddressOf RequestPasswordResetCB, AddressOf RequestNewPasswordCB, AddressOf SelectCharacterPreviewCB, AddressOf LoginCharacterCB, AddressOf ReturnToLoginCB, AddressOf CreateCharacterCB, AddressOf RequestDeleteCharCB, AddressOf ConfirmDeleteCharCB, AddressOf TransferCharacterCB)
    Else
        Call RegistrarError(0, "", "Failed to initialize babel UI with w:" & Width & " h:" & Height & " pixelSizee: " & pixelSize, 106)
    End If
    Exit Sub
InitializeUI_Err:
    Call RegistrarError(Err.Number, Err.Description, "BabelUI.InitializeUI", Erl)
End Sub

Public Sub InitializeInspectorUI(ByVal Width As Long, ByVal Height As Long)
On Error GoTo InitializeInspectorUI_Err
    Debug.Assert Width > 0 And Height > 0
102 Call CreateDebugWindow(Width, Height)
104 DebugUITexture.Height = Height
106 DebugUITexture.Width = Width
108 DebugUITexture.TextureHeight = NextPowerOf2(Height)
110 DebugUITexture.TextureWidth = NextPowerOf2(Width)
112 ReDim DebugUITexture.ImageBuffer(DebugUITexture.Height * DebugUITexture.Width * 4)
114 DebugUITexture.pixelSize = 4
116 Set DebugUITexture.Texture = SurfaceDB.CreateTexture(DebugUITexture.TextureWidth, DebugUITexture.TextureHeight)
    If DebugUITexture.Texture Is Nothing Then
118    Call RegistrarError(102, "texture undefined ", "BabelUI.InitializeInspectorUI", 102)
       Exit Sub
    End If
    DebugInitialized = True
    Exit Sub
InitializeInspectorUI_Err:
    Call RegistrarError(Err.Number, Err.Description, "BabelUI.InitializeInspectorUI", Erl)
End Sub

Public Sub DrawUITexture(ByRef TextureInfo As t_UITexture)
On Error GoTo DrawTexture_Err
    With TextureInfo
100     If Not .Texture Is Nothing Then
116         Call SpriteBatch.SetTexture(.Texture)
118         Call SpriteBatch.SetAlpha(False)
120         Call SpriteBatch.Draw(0, 0, .Width, .Height, COLOR_WHITE, , , .Width / .TextureWidth, .Height / .TextureHeight, 0)
        Else
            Call RegistrarError(102, "Undefined Texture", "BabelUI.DrawTexture", 202)
            Exit Sub
        End If
    End With
    Exit Sub
DrawTexture_Err:
    Call RegistrarError(Err.Number, Err.Description, "BabelUI.DrawTexture", Erl)
End Sub

Public Sub UpdateTexture(ByRef TextureInfo As t_UITexture)
On Error GoTo UpdateTexture_Err
    With TextureInfo
        If Not .Texture Is Nothing Then
            Call SurfaceDB.SetTextureData(.Texture, .ImageBuffer, UBound(.ImageBuffer), .TextureWidth, .Width, 0, .Height)
        End If
    End With
    Exit Sub
UpdateTexture_Err:
    Call RegistrarError(Err.Number, Err.Description, "BabelUI.UpdateTexture", Erl)
End Sub

Public Sub UpdateUI()
On Error GoTo UpdateInspectorUI_Err
100 Dim updateGui As Boolean
102 updateGui = GetBebelImageBuffer(UITexture.ImageBuffer(LBound(UITexture.ImageBuffer)), UBound(UITexture.ImageBuffer))
104 If updateGui Then
106     Call UpdateTexture(UITexture)
    End If
    Exit Sub
UpdateInspectorUI_Err:
    Call RegistrarError(Err.Number, Err.Description, "BabelUI.UpdateInspectorUI", Erl)
End Sub

Public Sub UpdateInspectorUI()
On Error GoTo UpdateInspectorUI_Err
100 Dim updateGui As Boolean
102 updateGui = GetDebugImageBuffer(DebugUITexture.ImageBuffer(LBound(DebugUITexture.ImageBuffer)), UBound(DebugUITexture.ImageBuffer))
108 If updateGui Then
110     Call UpdateTexture(DebugUITexture)
    End If
    Exit Sub
UpdateInspectorUI_Err:
    Call RegistrarError(Err.Number, Err.Description, "BabelUI.UpdateInspectorUI", Erl)
End Sub

Private Function GetStringFromPtr(ByVal Ptr As Long, ByVal size As Long) As String
    Dim Buffer() As Byte
    ReDim Buffer(0 To (size - 1)) As Byte
    CopyMemory Buffer(0), ByVal Ptr, size
    GetStringFromPtr = StrConv(Buffer, vbUnicode)
End Function

Public Sub LoginCB(ByRef LoginValue As LOGINDATA)
    Dim User, password As String
    If LoginValue.UserLen > 0 Then
        User = GetStringFromPtr(LoginValue.User, LoginValue.UserLen)
    End If
    If LoginValue.PasswordLen > 0 Then
        password = GetStringFromPtr(LoginValue.password, LoginValue.PasswordLen)
    End If
    Call SetActiveEnvironment(ServerEnvironment)
    Call DoLogin(User, password, LoginValue.storeCredentials > 0)
End Sub

Public Sub CreateAccount(ByRef NewAccount As NewAccountData)
    Dim User, password, Name, Surname As String
    If NewAccount.UserLen > 0 Then
        User = GetStringFromPtr(NewAccount.User, NewAccount.UserLen)
    End If
    If NewAccount.PasswordLen > 0 Then
        password = GetStringFromPtr(NewAccount.password, NewAccount.PasswordLen)
    End If
    If NewAccount.NameLen > 0 Then
        Name = GetStringFromPtr(NewAccount.Name, NewAccount.NameLen)
    End If
    If NewAccount.SurnameLen > 0 Then
        Surname = GetStringFromPtr(NewAccount.Surname, NewAccount.SurnameLen)
    End If
    Call SetActiveEnvironment(ServerEnvironment)
    Call ModLogin.CreateAccount(Name, Surname, User, password)
End Sub

Public Sub CloseClientCB()
    Call CloseClient
End Sub

Public Sub ResendValidationCodeCB(ByRef code As SINGLESTRINGPARAM)
    If code.Len > 0 Then
        CuentaEmail = GetStringFromPtr(code.Ptr, code.Len)
    End If
    Call SetActiveEnvironment(ServerEnvironment)
    Call ResendValidationCode(CuentaEmail)
End Sub

Public Sub ValidateCodeCB(ByRef Params As DOUBLESTRINGPARAM)
    If Params.FirstLen > 0 Then
        CuentaEmail = GetStringFromPtr(Params.FirstPtr, Params.FirstLen)
    End If
    If Params.SecondLen > 0 Then
        ValidationCode = GetStringFromPtr(Params.SecondPtr, Params.SecondLen)
    End If
    Call SetActiveEnvironment(ServerEnvironment)
    Call ValidateCode(CuentaEmail, ValidationCode)
End Sub

Public Sub SetHostCB(ByRef Params As SINGLESTRINGPARAM)
    If Params.Len > 0 Then
        ServerEnvironment = GetStringFromPtr(Params.Ptr, Params.Len)
    End If
End Sub

Public Sub RequestPasswordResetCB(ByRef Params As SINGLESTRINGPARAM)
    If Params.Len > 0 Then
        CuentaEmail = GetStringFromPtr(Params.Ptr, Params.Len)
    End If
    Call SetActiveEnvironment(ServerEnvironment)
    Call RequestPasswordReset(CuentaEmail)
End Sub

Public Sub RequestNewPasswordCB(ByRef Params As TRIPLESTRINGPARAM)
    If Params.FirstLen > 0 Then
        CuentaEmail = GetStringFromPtr(Params.FirstPtr, Params.FirstLen)
    End If
    If Params.SecondLen > 0 Then
        ValidationCode = GetStringFromPtr(Params.SecondPtr, Params.SecondLen)
    End If
    If Params.ThirdLen > 0 Then
        CuentaPassword = GetStringFromPtr(Params.ThirdPtr, Params.ThirdLen)
    End If
    Call SetActiveEnvironment(ServerEnvironment)
    Call RequestNewPassword(CuentaEmail, CuentaPassword, ValidationCode)
End Sub

Public Sub SelectCharacterPreviewCB(ByVal charindex As Long)
    charindex = charindex + 1
    If charindex < LBound(Pjs) Or charindex > CantidadDePersonajesEnCuenta Then
        Call SwitchMap(CreateCharMap)
        RenderCuenta_PosX = CreateCharMapX
        RenderCuenta_PosY = CreateCharMapY
        g_game_state.state = e_state_createchar_screen
    Else
        Call SwitchMap(Pjs(charindex).Mapa)
        RenderCuenta_PosX = Pjs(charindex).PosX
        RenderCuenta_PosY = Pjs(charindex).PosY
    End If
End Sub

Public Sub LoginCharacterCB(ByVal charindex As Long)
    charindex = charindex + 1
    If charindex > 0 Or charindex <= CantidadDePersonajesEnCuenta Then
        Call LoginCharacter(Pjs(charindex).nombre)
    End If
End Sub

Public Sub ReturnToLoginCB()
    Call LogOut
End Sub

Public Sub RequestDeleteCharCB(ByVal CharIndex As Long)
    DeleteUser = Pjs(CharIndex + 1).nombre
    Call RequestDeleteCharacter
End Sub

Public Sub ConfirmDeleteCharCB(ByVal CharIndex As Long, ByRef Code As SINGLESTRINGPARAM)
    DeleteUser = Pjs(CharIndex + 1).nombre
    delete_char_validate_code = GetStringFromPtr(Code.Ptr, Code.Len)
    ModAuth.LoginOperation = e_operation.ConfirmDeleteChar
    Call connectToLoginServer
End Sub

Public Sub CreateCharacterCB(ByRef charinfo As NEWCHARACTERDATA)
    Dim name As String
    name = GetStringFromPtr(charinfo.name, charinfo.NameLen)
    Call ModLogin.CreateCharacter(name, charinfo.Race + 1, charinfo.Gender + 1, charinfo.Class + 1, charinfo.Head, charinfo.City + 1)
End Sub

Public Sub TransferCharacterCB(ByVal CharIndex As Integer, ByRef Email As SINGLESTRINGPARAM)
    Dim Account As String
    Account = GetStringFromPtr(Email.Ptr, Email.Len)
    Call TransferChar(Pjs(CharIndex + 1).nombre, Account)
End Sub

Public Sub DisplayError(ByVal message As String, ByVal LocalizationStr As String)
    If BabelInitialized Then
        If LocalizationStr = "" Then
            Call SendErrorMessage(message, 0, 0)
        Else
            Call SendErrorMessage(LocalizationStr, 1, 0)
        End If
    Else
        Call MsgBox(message)
    End If
End Sub

Public Sub SendLoginCharacters(ByRef charlist() As UserCuentaPJS, ByVal charCount As Long)
    Call LoginCharacterListPrepare(charCount)
    Dim i As Integer
    For i = LBound(charlist) To charCount
        Call LoginAddCharacter(charlist(i).nombre, charlist(i).Head, charlist(i).Body, charlist(i).Casco, charlist(i).Escudo, charlist(i).Arma, charlist(i).nivel, charlist(i).Criminal, i - 1)
    Next i
    Call LoginSendCharacters
End Sub
