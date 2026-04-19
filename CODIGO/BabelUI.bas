Attribute VB_Name = "BabelUI"
Option Explicit
'windows messages


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


Public Type SINGLESTRINGPARAM
    Ptr As Long
    Len As Long
End Type




Public Type t_Color
    R As Byte
    G As Byte
    B As Byte
End Type


Public Enum e_NpcInfoMask
    AlmostDead = 1
    SeriouslyWounded = 2
    Wounded = 4
    LightlyWounded = 8
    Intact = 16
    Poisoned = 32
    Paralized = 64
    Inmovilized = 128
    Fighting = 256
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


Public Enum MouseButton
    kButton_None = 0
    kButton_Left = 1
    kButton_Middle = 2
    kButton_Right = 3
End Enum


Public GetRemoteError   As Boolean
Public IsGameDialogOpen As Boolean

Public Function GetStringFromPtr(ByVal Ptr As Long, ByVal size As Long) As String
    Dim Buffer() As Byte
    ReDim Buffer(0 To (size - 1)) As Byte
    CopyMemory Buffer(0), ByVal Ptr, size
    GetStringFromPtr = StrConv(Buffer, vbUnicode)
End Function

Public Sub DisplayError(ByVal message As String, ByVal LocalizationStr As String)
    Call MsgBox(message)
End Sub
