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
    Password As Long
    PasswordLen As Long
    storeCredentials As Long
End Type

Private Type NewAccountData
    User As Long
    UserLen As Long
    Password As Long
    PasswordLen As Long
    Name As Long
    NameLen As Long
    Surname As Long
    SurnameLen As Long
End Type

Private Type NEWCHARACTERDATA
    Name As Long
    NameLen As Long
    Gender As Long
    Race As Long
    Class As Long
    Head As Long
    City As Long
End Type

Public Type SINGLESTRINGPARAM
    Ptr As Long
    Len As Long
End Type

Public Type DOUBLESTRINGPARAM
    FirstPtr As Long
    FirstLen As Long
    SecondPtr As Long
    SecondLen As Long
End Type

Public Type TRIPLESTRINGPARAM
    FirstPtr As Long
    FirstLen As Long
    SecondPtr As Long
    SecondLen As Long
    ThirdPtr As Long
    ThirdLen As Long
End Type

Public Type BABELSETTINGS
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
    desc As String
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
    CreateNewScenario As Long
    JoinScenario As Long
    UpdateSkillList As Long
    SendGuildRequest As Long
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
    eTeleportToMap = 28
    eDisplayGuildDetails = 29
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
    eDisableDungeonLighting = 26
End Enum

Public Enum e_SafeType
    eGroup = 1
    eClan = 2
    eAttack = 3
    eResurrecion = 4
End Enum

Public Type t_NewScenearioSettings
    MinLevel As Byte
    MaxLevel As Byte
    MinPlayers As Byte
    MaxPlayers As Byte
    TeamSize As Byte
    TeamType As Byte
    InscriptionFee As Long
    ScenearioType As Byte
    RoundAmount As Byte
End Type

Public Type t_GuildInfo
    Name As String
    Founder As String
    CreationDate As String
    Leader As String
    MemberCount As Integer
    Aligment As String
    Description As String
    level As Byte
End Type




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




Public GetRemoteError As Boolean

Public IsGameDialogOpen As Boolean

Public Function GetStringFromPtr(ByVal Ptr As Long, ByVal size As Long) As String
    On Error Goto GetStringFromPtr_Err
    Dim Buffer() As Byte
    ReDim Buffer(0 To (size - 1)) As Byte
    CopyMemory Buffer(0), ByVal Ptr, size
    GetStringFromPtr = StrConv(Buffer, vbUnicode)
    Exit Function
GetStringFromPtr_Err:
    Call TraceError(Err.Number, Err.Description, "BabelUI.GetStringFromPtr", Erl)
End Function

Public Sub DisplayError(ByVal message As String, ByVal LocalizationStr As String)
    On Error Goto DisplayError_Err
    Call MsgBox(message)
    Exit Sub
DisplayError_Err:
    Call TraceError(Err.Number, Err.Description, "BabelUI.DisplayError", Erl)
End Sub


