Attribute VB_Name = "BabelUI"
Option Explicit

Const CreateCharMap = 782
Const CreateCharMapX = 25
Const CreateCharMapY = 35

'windows messages
Private Const WM_DESTROY = &H2
Private Const WM_MOUSEWHEEL = &H20A

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

Public Type t_Color
    r As Byte
    G As Byte
    b As Byte
End Type

Public Type t_ChatMessage
    Sender As String
    SenderColor As t_Color
    Text As String
    TextColor As t_Color
    BoldText As Byte
    ItalicText As Byte
End Type

Public Enum e_CDTypeMask
    eBasicAttack = 1
    eRangedAttack = 2
    eMagic = 4
    eUsable = 8
    eCustom = 16
End Enum

Public Type t_InvItem
    Slot As Byte
    ObjIndex As Integer
    GrhIndex As Long
    ObjType As Byte
    Equiped As Byte
    CanUse As Byte
    Amount As Integer
    MinHit As Integer
    MaxHit As Integer
    MinDef As Integer
    MaxDef As Integer
    Value As Single
    Cooldown As Long
    CDType As Integer
    CDMask As Long
    Amunition As Integer
    IsBindable As Byte
    Name As String
    Desc As String
End Type

Public Type t_GamePlayCallbacks
    HandleConsoleMsg As Long
    ShowDialog As Long
    SelectInvSlot As Long
    UseInvSlot As Long
    SelectSpellSlot As Long
    UseSpellSlot As Long
    UpdateFocus As Long
    UpdateOpenDialog As Long
    OpenLink As Long
    ClickGold As Long
    MoveInvSlot As Long
    RequestAction As Long
    UseKey As Long
    MoveSpellSlot As Long
    RequestDeleteItem As Long
    UpdateScrollPos As Long
    TeleportToMiniMapPos As Long
    UpdateCombatAndGlobalChat As Long
    UpdateHotKeySlot As Long
    UpdateHideHotkeyState As Long
    HandleQuestionResponse As Long
    MoveMerchanSlot As Long
    CloseMerchant As Long
    BuyItem As Long
    SellItem As Long
    BuyAOShop As Long
    UpdateIntSetting As Long
End Type

Public Type t_SpellSlot
    Slot As Byte
    SpellIndex As Integer
    icon As Long
    Cooldown As Long
    IsBindable As Byte
    SpellName As String
End Type

Public Type t_ShopItem
    ObjIndex As Long
    Price As Long
End Type

Public Enum e_ChatMode
    NormalChat = 0
    ClanChat = 1
End Enum

Public Enum e_ActionRequest
    eMinimize = 1
    eClose = 2
    eOpenMinimap = 3
    eOpenClanDialog = 4
    eOpenChallenge = 5
    eOpenKeys = 6
    eOpenActiveQuest = 7
    eGoHome = 8
    eShowStats = 9
    eUpdateGroupLock = 10
    eUpdateClanSafeLock = 11
    eUpdateAttackSafeLock = 12
    eUpdateResurrectionLock = 13
    eReportBug = 14
    eRequestSkill = 15
    eOpenGroupDialog = 16
    eOpenGmPannel = 17
    eOpenCreateObjMenu = 18
    eOpenSpawnMenu = 19
    eSetGmInvisible = 20
    eDisplayHPInfo = 21
    eOpenSettings = 22
    eDisplayInventory = 23
    eDisplaySpells = 24
    eSetMeditate = 25
    eOpenKeySettings = 26
    eSaveSettings = 27
End Enum

Public Enum e_UpdateSetting
    eCopyDialogsEnabled = 1
    eWriteAndMove = 2
    eBlockSpellListScroll = 3
    eThrowSpellLockBehavior = 4
    eMouseSens = 5
    eUserGraphicCursor = 6
    eLanguage = 7
    eRenderNpcText = 8
    eTutorialEnabled = 9
    eShowFps = 10
    eMoveGameWindow = 11
    eCharacterBreathing = 12
    eFullScreen = 13
    eDisplayFloorItemInfo = 14
    eDisplayFullNumbersInventory = 15
    eEnableBabelUI = 16
    eEnableMusic = 17
    eEnableFx = 18
    eEnableAmbient = 19
    eSailFx = 20
    eInvertChannels = 21
    eMusicVolume = 22
    eFxVolume = 23
    eAmbientVolume = 24
    eLightSettings = 25
End Enum

Public Enum e_SafeType
    eGroup = 1
    eClan = 2
    eAttack = 3
    eResurrecion = 4
End Enum

Private ServerEnvironment As String

Public Declare Function InitializeBabel Lib "BabelUI.dll" (ByRef Settings As BABELSETTINGS) As Boolean
Public Declare Function GetBebelImageBuffer Lib "BabelUI.dll" Alias "GetImageBuffer" (ByRef Buffer As Byte, ByVal size As Long) As Boolean
Public Declare Sub BabelSendMouseEvent Lib "BabelUI.dll" Alias "SendMouseEvent" (ByVal PosX As Long, ByVal PosY As Long, ByVal EvtType As Long, ByVal button As Long)
Public Declare Sub SendScrollEvent Lib "BabelUI.dll" (ByVal distance As Long)
Public Declare Sub BabelSendKeyEvent Lib "BabelUI.dll" Alias "SendKeyEvent" (ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal EvtType As Long, ByVal CapsState As Boolean, ByVal Inspector As Boolean)
Public Declare Function NextPowerOf2 Lib "BabelUI.dll" (ByVal original As Long) As Long
Public Declare Sub RegisterCallbacks Lib "BabelUI.dll" (ByVal Login As Long, ByVal CloseClient As Long, ByVal CreateAccount As Long, ByVal SetHost As Long, ByVal ValidateAccountr As Long, ByVal ResendCode As Long, ByVal RequestPasswordReset As Long, _
                                                        ByVal RequestNewPassord As Long, ByVal SelectCharacter As Long, ByVal LoginCharacter As Long, ByVal ReturnToLogin As Long, ByVal CreateCharacter As Long, ByVal RequestCharDelete As Long, _
                                                        ByVal ConfirmCharDelete As Long, ByVal TransferCharacter As Long)
Public Declare Sub SendErrorMessage Lib "BabelUI.dll" (ByVal message As String, ByVal localize As Long, ByVal Action As Long)
Public Declare Sub SetActiveScreen Lib "BabelUI.dll" (ByVal screenName As String)
Public Declare Sub SetLoadingMessage Lib "BabelUI.dll" (ByVal message As String, ByVal localize As Long)
Public Declare Sub LoginCharacterListPrepare Lib "BabelUI.dll" (ByVal CharacterCount As Long)
Public Declare Sub LoginAddCharacter Lib "BabelUI.dll" (ByVal Name As String, ByVal Head As Long, ByVal Body As Long, ByVal helm As Long, ByVal Shield As Long, ByVal weapon As Long, ByVal level As Long, ByVal status As Long, ByVal Index As Long, ByVal Class As Integer)
Public Declare Sub LoginSendCharacters Lib "BabelUI.dll" ()
Public Declare Sub RequestDeleteCode Lib "BabelUI.dll" ()
Public Declare Sub RemoveCharacterFromList Lib "BabelUI.dll" (ByVal Index As Long)
Public Declare Function GetTelemetry Lib "BabelUI.dll" (ByRef code As Byte, ByRef DataBuff As Byte, ByVal BuffSize As Long) As Long

'Gameplay interface
Public Declare Sub SetUserStats Lib "BabelUI.dll" (ByRef code As t_UserStats)
Public Declare Sub SetUserName Lib "BabelUI.dll" (ByVal username As String)
Public Declare Sub SendChatMessage Lib "BabelUI.dll" (ByRef message As t_ChatMessage)
Public Declare Sub UpdateFps Lib "BabelUI.dll" (ByVal fps As Long)
Public Declare Sub SetInventoryLevel Lib "BabelUI.dll" (ByVal fps As Long)
Public Declare Sub RegisterGameplayCallbacks Lib "BabelUI.dll" (ByRef t_GamePlayCallbacks As t_GamePlayCallbacks)
Public Declare Sub SetInvSlot Lib "BabelUI.dll" (ByRef SlotInfo As t_InvItem)
Public Declare Sub SetSpellSlot Lib "BabelUI.dll" (ByRef SlotInfo As t_SpellSlot)
Public Declare Sub UpdateHpValue Lib "BabelUI.dll" (ByVal Hp As Long, ByVal Shield As Long)
Public Declare Sub UpdateManaValue Lib "BabelUI.dll" (ByVal NewValue As Long)
Public Declare Sub UpdateStaminaValue Lib "BabelUI.dll" (ByVal NewValue As Long)
Public Declare Sub UpdateDrinkValue Lib "BabelUI.dll" (ByVal NewValue As Long)
Public Declare Sub UpdateFoodValue Lib "BabelUI.dll" (ByVal NewValue As Long)
Public Declare Sub UpdateGold Lib "BabelUI.dll" (ByVal NewValue As Long, ByVal SafeGoldForLevel As Long)
Public Declare Sub UpdateExp Lib "BabelUI.dll" (ByVal Current As Long, ByVal max As Long)
Public Declare Sub OpenChat Lib "BabelUI.dll" (ByVal mode As Long)
Public Declare Sub UpdateStrAndAgiBuff Lib "BabelUI.dll" (ByVal str As Byte, ByVal Agi As Byte, ByVal StrState As Byte, ByVal StrState As Byte)
Public Declare Sub UpdateMapInfo Lib "BabelUI.dll" (ByVal MapNumber As Long, ByVal MapName As String, ByVal NpcCount As Integer, ByRef NpcList As t_QuestNPCMapData, ByVal IsSafe As Byte)
Public Declare Sub UpdateUserPos Lib "BabelUI.dll" (ByVal TileX As Integer, ByVal TileY As Integer, ByRef MapPos As t_Position)
Public Declare Sub UpdateGroupPos Lib "BabelUI.dll" (ByRef MapPos As t_Position, ByVal GroupIndex As Integer)
Public Declare Sub SetKeySlot Lib "BabelUI.dll" (ByRef SlotInfo As t_InvItem)
Public Declare Sub UpdateIntervals Lib "BabelUI.dll" (ByRef Intervals As t_Intervals)
Public Declare Sub ActivateInterval Lib "BabelUI.dll" (ByVal IntervalType As Long)
Public Declare Sub SetSafeState Lib "BabelUI.dll" (ByVal SafeType As Long, ByVal State As Long)
Public Declare Sub UpdateOnlines Lib "BabelUI.dll" (ByVal NewValue As Long)
Public Declare Sub UpdateGameTime Lib "BabelUI.dll" (ByVal Hour As Long, ByVal Minutes As Long)
Public Declare Sub UpdateIsGameMaster Lib "BabelUI.dll" (ByVal NewState As Long)
Public Declare Sub UpdateMagicResistance Lib "BabelUI.dll" (ByVal NewValue As Long)
Public Declare Sub UpdateMagicAttack Lib "BabelUI.dll" (ByVal NewValue As Long)
Public Declare Sub SetWhisperTarget Lib "BabelUI.dll" (ByVal UserName As String)
Public Declare Sub PasteText Lib "BabelUI.dll" (ByVal Text As String)
Public Declare Sub ReloadSettings Lib "BabelUI.dll" ()
Public Declare Sub SetRemoteTrackingState Lib "BabelUI.dll" (ByVal State As Long)
Public Declare Sub UpdateInvAndSpellTracking Lib "BabelUI.dll" (ByVal SelectedTab As Long, ByVal SelectedSpell As Long, ByVal FirstSpellToDisplay As Long)
Public Declare Sub HandleRemoteUserClick Lib "BabelUI.dll" ()
Public Declare Sub UpdateRemoteMousePos Lib "BabelUI.dll" (ByVal PosX As Long, ByVal PosY As Long)
Public Declare Sub StartSpellCd Lib "BabelUI.dll" (ByVal SpellId As Long, ByVal CdTime As Long)
Public Declare Sub UpdateCombatAndGlobalChatSettings Lib "BabelUI.dll" (ByVal SpellId As Long, ByVal CdTime As Long)
Public Declare Sub ActivateStunTimer Lib "BabelUI.dll" (ByVal Duration As Long)
Public Declare Sub UpdateHoykeySlot Lib "BabelUI.dll" (ByVal SlotIndex As Long, ByRef SlotInfo As t_HotkeyEntry)
Public Declare Sub ActivateFeatureToggle Lib "BabelUI.dll" (ByVal ToggleName As String)
Public Declare Sub ClearToggles Lib "BabelUI.dll" ()
Public Declare Sub SetHotkeyHideState Lib "BabelUI.dll" (ByVal HideHotkeyState As Long)
Public Declare Sub ShowQuestion Lib "BabelUI.dll" (ByVal QuestionText As String)
Public Declare Sub OpenMerchant Lib "BabelUI.dll" ()
Public Declare Sub UpdateMerchantSlot Lib "BabelUI.dll" (ByRef SlotInfo As t_InvItem)
Public Declare Sub OpenAo20Shop Lib "BabelUI.dll" (ByVal AvailableCredits As Long, ByVal ItemCount As Long, ByRef ItemList As t_ShopItem)

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
Public SaveUseBabelUI As Boolean
Public InputFocus As Boolean
Public IsGameDialogOpen As Boolean

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
       SaveUseBabelUI = UseBabelUI
102    LogError ("Initilalize UI: " & UseBabelUI)
104    CheckAndSetBabelUIUsage = UseBabelUI
    Exit Function
CheckAndSetBabelUIUsage_Err:
    CheckAndSetBabelUIUsage = False
    Call RegistrarError(Err.Number, Err.Description, "BabelUI.CheckAndSetBabelUIUsage", Erl)
End Function

Public Function GetMainHwdn() As String
    If UseBabelUI Then
        GetMainHwdn = frmBabelUI.hwnd
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
116     Call InitializeTexture
118     BabelInitialized = True
        Call RegisterCallbacks(AddressOf LoginCB, AddressOf CloseClientCB, AddressOf BabelUI.CreateAccount, AddressOf SetHostCB, AddressOf ValidateCodeCB, AddressOf ResendValidationCodeCB, AddressOf RequestPasswordResetCB, AddressOf RequestNewPasswordCB, AddressOf SelectCharacterPreviewCB, AddressOf LoginCharacterCB, AddressOf ReturnToLoginCB, AddressOf CreateCharacterCB, AddressOf RequestDeleteCharCB, AddressOf ConfirmDeleteCharCB, AddressOf TransferCharacterCB)
        Dim GameplayCallbacks As t_GamePlayCallbacks
        GameplayCallbacks.HandleConsoleMsg = FARPROC(AddressOf HandleConsoleMsgCB)
        GameplayCallbacks.ShowDialog = FARPROC(AddressOf HandleShowDialogCB)
        GameplayCallbacks.SelectInvSlot = FARPROC(AddressOf HandleSelectInvSlotCB)
        GameplayCallbacks.UseInvSlot = FARPROC(AddressOf HandleUseInvSlotCB)
        GameplayCallbacks.SelectSpellSlot = FARPROC(AddressOf HandleSelectSpellSlotCB)
        GameplayCallbacks.UseSpellSlot = FARPROC(AddressOf HandleUseSpellSlotCB)
        GameplayCallbacks.UpdateOpenDialog = FARPROC(AddressOf IsGameDialogOpenCB)
        GameplayCallbacks.UpdateFocus = FARPROC(AddressOf HandleUpdateFocusStateCB)
        GameplayCallbacks.OpenLink = FARPROC(AddressOf OpenLinkCB)
        GameplayCallbacks.ClickGold = FARPROC(AddressOf ClickGoldCB)
        GameplayCallbacks.MoveInvSlot = FARPROC(AddressOf MoveInvSlotDB)
        GameplayCallbacks.RequestAction = FARPROC(AddressOf RequestActionCB)
        GameplayCallbacks.UseKey = FARPROC(AddressOf UseKeyCB)
        GameplayCallbacks.MoveSpellSlot = FARPROC(AddressOf MoveSpellSlotCB)
        GameplayCallbacks.RequestDeleteItem = FARPROC(AddressOf RequestDeleteItemCB)
        GameplayCallbacks.UpdateScrollPos = FARPROC(AddressOf UpdateSpellScrollPosCB)
        GameplayCallbacks.TeleportToMiniMapPos = FARPROC(AddressOf TeleportToMiniMapPos)
        GameplayCallbacks.UpdateCombatAndGlobalChat = FARPROC(AddressOf UpdateCombatAndGlobalChatCB)
        GameplayCallbacks.UpdateHotKeySlot = FARPROC(AddressOf UpdateHotkeySlotCB)
        GameplayCallbacks.UpdateHideHotkeyState = FARPROC(AddressOf UpdateHideHotkeyCB)
        GameplayCallbacks.HandleQuestionResponse = FARPROC(AddressOf HandleQuestionResponseCB)
        GameplayCallbacks.MoveMerchanSlot = FARPROC(AddressOf HandleMoveMerchanSlotCB)
        GameplayCallbacks.CloseMerchant = FARPROC(AddressOf HandleCloseMerchantCB)
        GameplayCallbacks.BuyItem = FARPROC(AddressOf HandleBuyItemCB)
        GameplayCallbacks.SellItem = FARPROC(AddressOf HandleSellItemCB)
        GameplayCallbacks.BuyAOShop = FARPROC(AddressOf HandleBuyAoShopCB)
        GameplayCallbacks.UpdateIntSetting = FARPROC(AddressOf HandleUpdateIntSetting)
        Call RegisterGameplayCallbacks(GameplayCallbacks)
    Else
        Call RegistrarError(0, "", "Failed to initialize babel UI with w:" & Width & " h:" & Height & " pixelSizee: " & pixelSize, 106)
    End If
    Exit Sub
InitializeUI_Err:
    Call RegistrarError(Err.Number, Err.Description, "BabelUI.InitializeUI", Erl)
End Sub

Public Sub InitializeTexture()
    Set UITexture.Texture = SurfaceDB.CreateTexture(UITexture.TextureWidth, UITexture.TextureHeight)
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
        Call LoginAddCharacter(charlist(i).nombre, charlist(i).Head, charlist(i).Body, charlist(i).Casco, charlist(i).Escudo, charlist(i).Arma, charlist(i).nivel, charlist(i).Criminal, i - 1, charlist(i).Clase)
    Next i
    Call LoginSendCharacters
End Sub

Public Sub HandleConsoleMsgCB(ByRef msg As SINGLESTRINGPARAM)
    Dim MsgStr As String
    MsgStr = GetStringFromPtr(msg.Ptr, msg.Len)
    Call HandleChatMsg(MsgStr)
End Sub

Public Sub HandleShowDialogCB(ByRef dialog As SINGLESTRINGPARAM)
    Dim DialogName As String
    DialogName = GetStringFromPtr(dialog.Ptr, dialog.Len)
    Dim Frm As Form
    Set Frm = Forms.Add(DialogName)
    Frm.Show , frmBabelUI
End Sub

Public Sub HandleSelectInvSlotCB(ByVal Slot As Long)
    Call ModGameplayUI.SelectItemSlot(Slot)
End Sub

Public Sub HandleUseInvSlotCB(ByVal Slot As Long)
    Call ModGameplayUI.UserItemClick
End Sub

Public Sub HandleSelectSpellSlotCB(ByVal Slot As Long)
    SelectedSpellSlot = Slot
    If Seguido = 1 Then
        Call WriteNotifyInventarioHechizos(2, SelectedSpellSlot, FirstSpellInListToRender)
    End If
End Sub

Public Sub HandleUseSpellSlotCB(ByVal Slot As Long)
    Call UseSpell(Slot, "")
End Sub

Public Sub HandleUpdateFocusStateCB(ByVal Focus As Boolean)
    InputFocus = Focus > 0
End Sub

Public Sub IsGameDialogOpenCB(ByVal IsOpen As Boolean)
    IsGameDialogOpen = IsOpen > 0
End Sub

Public Sub OpenLinkCB(ByRef link As SINGLESTRINGPARAM)
    Dim Url As String
    Url = GetStringFromPtr(link.Ptr, link.Len)
    ShellExecute ByVal 0&, "open", _
        Url, _
        vbNullString, vbNullString, _
        vbMaximizedFocus
End Sub

Public Sub ClickGoldCB()
    UserInventory.SelectedSlot = FLAGORO
    If UserStats.GLD > 0 Then
        frmCantidad.Show , frmBabelUI
    End If
End Sub

Public Sub MoveInvSlotDB(ByVal SourceSlot As Long, ByVal DestSlot As Long)
    If DestSlot > 0 And SourceSlot <> DestSlot Then
        Call WriteItemMove(SourceSlot, DestSlot)
    End If
End Sub

Public Sub RequestActionCB(ByVal ActionId As Long)
    Select Case ActionId
    Case e_ActionRequest.eMinimize
        frmBabelUI.WindowState = vbMinimized

    Case e_ActionRequest.eClose
        If frmCerrar.visible Then Exit Sub
        Dim mForm As Form
        For Each mForm In Forms
            If mForm.hwnd <> frmBabelUI.hwnd Or mForm.hwnd <> frmDebugUI.hwnd Then Unload mForm
            Set mForm = Nothing
        Next
        frmCerrar.Show , frmBabelUI

    Case e_ActionRequest.eOpenMinimap
        ExpMult = 1
        OroMult = 1
        Call frmMapaGrande.CalcularPosicionMAPA
        frmMapaGrande.Picture = LoadInterface("ventanamapa.bmp")
        frmMapaGrande.Show , frmBabelUI

    Case e_ActionRequest.eOpenClanDialog
        If frmGuildLeader.visible Then Unload frmGuildLeader
        Call WriteRequestGuildLeaderInfo

    Case e_ActionRequest.eOpenChallenge
        Call ParseUserCommand("/RETAR")

    Case e_ActionRequest.eOpenKeys
        ' Code for eOpenKeys case

    Case e_ActionRequest.eOpenActiveQuest
        Call WriteQuestListRequest

    Case e_ActionRequest.eGoHome
        Call ParseUserCommand("/HOGAR")

    Case e_ActionRequest.eShowStats
        LlegaronAtrib = False
        LlegaronStats = False
        Call WriteRequestAtributes
        Call WriteRequestMiniStats

    Case e_ActionRequest.eUpdateGroupLock
        Call WriteParyToggle

    Case e_ActionRequest.eUpdateClanSafeLock
        Call WriteSeguroClan

    Case e_ActionRequest.eUpdateAttackSafeLock
        Call WriteSafeToggle

    Case e_ActionRequest.eUpdateResurrectionLock
        Call WriteSeguroResu

    Case e_ActionRequest.eReportBug
        FrmGmAyuda.Show vbModeless, frmBabelUI

    Case e_ActionRequest.eRequestSkill
        Call ModGameplayUI.RequestSkills
        
    Case e_ActionRequest.eOpenGroupDialog
        If FrmGrupo.visible = False Then
            Call WriteRequestGrupo
        End If
        
    Case e_ActionRequest.eOpenGmPannel
        frmPanelgm.Width = 4860
        Call WriteSOSShowList
        Call WriteGMPanel
    Case e_ActionRequest.eOpenCreateObjMenu
        Call OpenCreateObjectMenu
        
    Case e_ActionRequest.eOpenSpawnMenu
        Call WriteSpawnListRequest
        
    Case e_ActionRequest.eSetGmInvisible
        Call ParseUserCommand("/INVISIBLE")
    Case e_ActionRequest.eDisplayHPInfo
        Call ParseUserCommand("/PROMEDIO")
    Case e_ActionRequest.eOpenSettings
        Call frmOpciones.init
    Case e_ActionRequest.eDisplayInventory
        Call SelectInvenrotyTab
    Case e_ActionRequest.eDisplaySpells
        Call SelectSpellTab
    Case e_ActionRequest.eSetMeditate
        Call RequestMeditate
    Case e_ActionRequest.eOpenKeySettings
        Call frmCustomKeys.Show(vbModeless, GetGameplayForm)
    Case e_ActionRequest.eSaveSettings
        Call GuardarOpciones
End Select

End Sub

Public Sub UseKeyCB(ByVal KeyIndex As Long)
    Call WriteUseKey(KeyIndex + 1)
End Sub

Public Sub MoveSpellSlotCB(ByVal FromSlot As Long, ByVal ToSlot As Long)
    Dim Diff, AbsDiff, i As Long
    Diff = FromSlot - ToSlot
    AbsDiff = Abs(Diff)
    For i = 0 To AbsDiff - 1
        Call WriteMoveSpell(Diff > 0, FromSlot - (i * (Diff / AbsDiff)))
    Next i
End Sub

Public Sub RequestDeleteItemCB(ByVal SelectedSlot As Long)
    Call WriteDeleteItem(SelectedSlot)
End Sub

Public Sub UpdateSpellScrollPosCB(ByVal FirestSpellInList As Long)
    FirstSpellInListToRender = FirestSpellInList
    If Seguido = 1 Then
        Call WriteNotifyInventarioHechizos(2, SelectedSpellSlot, FirstSpellInListToRender)
    End If
End Sub

Public Sub UpdateHotkeySlotCB(ByVal SlotIndex As Long, ByRef SlotInfo As t_HotkeyEntry)
    Call SetHotkey(SlotInfo.Index, SlotInfo.LastKnownSlot, SlotInfo.Type, SlotIndex)
End Sub

Public Sub UpdateHideHotkeyCB(ByVal State As Integer)
    HideHotkeys = State > 0
    Call SaveHideHotkeys
End Sub

Public Sub TeleportToMiniMapPos(ByVal PosX As Long, ByVal PosY As Long)
    Dim x As Single
    Dim y As Single
    x = PosX
    y = PosY
    Call GetMinimapPosition(x, y)
    Call ParseUserCommand("/TELEP YO " & UserMap & " " & CByte(x) & " " & CByte(y))
End Sub

Public Sub UpdateCombatAndGlobalChatCB(ByVal CombatState As Long, ByVal GlobalState As Long)
    ChatCombate = CombatState
    ChatGlobal = GlobalState
    Call WriteMacroPos
End Sub

Public Sub HandleQuestionResponseCB(ByVal Response As Integer)
    Call HandleQuestionResponse(Response > 0)
End Sub

Public Sub HandleMoveMerchanSlotCB(ByVal FromSlot As Integer, ByVal ToSlot As Integer)
End Sub
Public Sub HandleCloseMerchantCB()
    Call WriteCommerceEnd
End Sub
Public Sub HandleBuyItemCB(ByVal Slot As Integer, ByVal Amount As Integer)
    Call WriteCommerceBuy(Slot, Amount)
End Sub

Public Sub HandleSellItemCB(ByVal Slot As Integer, ByVal Amount As Integer)
    Call WriteCommerceSell(Slot, Amount)
End Sub

Public Sub HandleBuyAoShopCB(ByVal ObjIndex As Integer)
    Call writeBuyShopItem(ObjIndex)
End Sub

Public Sub HandleUpdateIntSetting(ByVal SettingType As Long, ByVal Value As Long)
On Error GoTo HandleUpdateIntSetting_Err
    Select Case SettingType
        Case eCopyDialogsEnabled
            CopiarDialogoAConsola = Value
        Case eWriteAndMove
            PermitirMoverse = Value
        Case eBlockSpellListScroll
            ScrollArrastrar = Value
        Case eThrowSpellLockBehavior
            ModoHechizos = Value
        Case eMouseSens
            SensibilidadMouse = Value
            Call General_Set_Mouse_Speed(Value)
        Case eUserGraphicCursor
            CursoresGraficos = Value > 0
            Call SaveSetting("VIDEO", "CursoresGraficos", Value)
        Case eLanguage
            Call SaveSetting("OPCIONES", "Localization", Value)
        Case eRenderNpcText
            npcs_en_render = Value
            Call SaveSetting("OPCIONES", "NpcsEnRender", Value)
        Case eTutorialEnabled
            MostrarTutorial = Value
            If MostrarTutorial Then
                Dim i As Long
                
                For i = 1 To UBound(tutorial)
                    Call SaveSetting("TUTORIAL" & i, "Activo", 1)
                    tutorial(i).Activo = 1
                Next i
            End If
            Call SaveSetting("INITTUTORIAL", "MostrarTutorial", Value)
        Case eShowFps
            FPSFLAG = Value
        Case eMoveGameWindow
            MoverVentana = Value
        Case eCharacterBreathing
            MostrarRespiracion = Value > 0
        Case eFullScreen
            Debug.Print "Update full screen " & Value
            If PantallaCompleta = (Value > 0) Then Exit Sub
            PantallaCompleta = Value > 0
            If PantallaCompleta Then
                Call SetResolution
            Else
                Call ResetResolution
            End If
        Case eDisplayFloorItemInfo
            InfoItemsEnRender = Value > 0
        Case eDisplayFullNumbersInventory
            NumerosCompletosInventario = Value > 0
        Case eEnableBabelUI
            SaveUseBabelUI = Value
            Call SaveSetting("OPCIONES", "UseExperimentalUI", Value)
        Case eEnableMusic
            If Value = 0 Then
                Sound.Music_Stop
                Musica = CONST_DESHABILITADA
            Else
                Musica = CONST_MP3
                Sound.NextMusic = MapDat.music_numberHi
                Sound.Fading = 100
            End If
        Case eEnableFx
            Fx = Value
            If Value = 0 Then
                Call Sound.Sound_Stop_All
            End If
        Case eEnableAmbient
            If Value = 0 Then
                AmbientalActivated = 0
                Sound.LastAmbienteActual = 0
                Sound.AmbienteActual = 0
                Sound.Ambient_Stop
            Else
                AmbientalActivated = 1
                Call AmbientarAudio(UserMap)

            End If
        Case eSailFx
            FxNavega = Value
        Case eInvertChannels
            InvertirSonido = Value
            Sound.InvertirSonido = Value > 0
        Case eMusicVolume
            If Musica <> CONST_DESHABILITADA Then
                Sound.Music_Volume_Set Value
                Sound.VolumenActualMusicMax = Value
                VolMusic = Sound.VolumenActualMusicMax
            End If
        Case eFxVolume
            Sound.VolumenActual = Value
            VolFX = Sound.VolumenActual
        Case eAmbientVolume
            Sound.Ambient_Volume_Set Value
            VolAmbient = Value
        Case eLightSettings
            Call SaveSetting("VIDEO", "LuzGlobal", Value)
            selected_light = Value
    End Select
    Exit Sub
HandleUpdateIntSetting_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmOpciones.Command1_Click", Erl)
End Sub
        
Public Function BabelEditWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
        '<EhHeader>
        On Error GoTo BabelEditWndProc_Err
        '</EhHeader>

100     Select Case uMsg
           Case WM_MOUSEWHEEL
            Dim delta As Long
102         delta = HiWord(wParam)
106         Call SendScrollEvent(delta)
            '[other messages will go here later]
108        Case WM_DESTROY
110          Call UnSubclass(hwnd, PtrEditWndProc)
        End Select
       
112     BabelEditWndProc = DefSubclassProc(hwnd, uMsg, wParam, lParam)
        '<EhFooter>
        Exit Function

BabelEditWndProc_Err:
        Debug.Print Err.Description & vbCrLf & _
               "in Argentum20.BabelUI.BabelEditWndProc " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Function
    
Private Function PtrEditWndProc() As Long
    PtrEditWndProc = FARPROC(AddressOf BabelEditWndProc)
End Function

Public Sub UpdateBuffState()
    Call UpdateStrAndAgiBuff(UserStats.str, UserStats.Agi, UserStats.StrState, UserStats.AgiState)
End Sub

Public Function GetCDMaskForItem(ByRef Item As ObjDatas) As Long
    If Item.ObjType = 2 Then
        If Item.proyectil > 0 Then
            GetCDMaskForItem = e_CDTypeMask.eRangedAttack
        End If
        If Item.Amunition = 0 Then
            GetCDMaskForItem = GetCDMaskForItem Or e_CDTypeMask.eBasicAttack
        End If
    End If
    If IsUsableItem(Item) Then
        GetCDMaskForItem = GetCDMaskForItem Or e_CDTypeMask.eUsable
    End If
    If Item.CDType > 0 Then
        GetCDMaskForItem = GetCDMaskForItem Or e_CDTypeMask.eCustom
    End If
End Function
