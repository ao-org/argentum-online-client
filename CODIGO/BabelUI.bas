Attribute VB_Name = "BabelUI"
Option Explicit


Private Type LOGINDATA
    user As Long
    userLen As Long
    password As Long
    passwordLen As Long
End Type

Public Declare Function InitializeBabel Lib "BabelUI.dll" Alias "_InitializeBabel@8" (ByVal Width As Long, ByVal Height As Long) As Boolean
Public Declare Function GetBebelImageBuffer Lib "BabelUI.dll" Alias "_GetImageBuffer@8" (ByRef Buffer As Byte, ByVal size As Long) As Boolean
Public Declare Sub BabelSendMouseEvent Lib "BabelUI.dll" Alias "_SendMouseEvent@16" (ByVal posX As Long, ByVal posY As Long, ByVal EvtType As Long, ByVal Button As Long)
Public Declare Sub BabelSendKeyEvent Lib "BabelUI.dll" Alias "_SendKeyEvent@20" (ByVal KeyCode As Integer, ByVal Shift As Boolean, ByVal EvtType As Long, ByVal CapsState As Boolean, ByVal Inspector As Boolean)
Public Declare Function NextPowerOf2 Lib "BabelUI.dll" Alias "_NextPowerOf2@4" (ByVal original As Long) As Long
Public Declare Sub RegisterCallbacks Lib "BabelUI.dll" Alias "_RegisterCallbacks@8" (ByVal LoginCallback As Long, ByVal Closeclient As Long)
Public Declare Sub SendErrorMessage Lib "BabelUI.dll" Alias "_SendErrorMessage@12" (ByVal Message As String, ByVal MessageType As Long, ByVal Action As Long)
'debug info
Public Declare Function CreateDebugWindow Lib "BabelUI.dll" Alias "_CreateDebugWindow@8" (ByVal Width As Long, ByVal Height As Long) As Boolean
Public Declare Function GetDebugImageBuffer Lib "BabelUI.dll" Alias "_GetDebugImageBuffer@8" (ByRef Buffer As Byte, ByVal size As Long) As Boolean
Public Declare Sub SendDebugMouseEvent Lib "BabelUI.dll" Alias "_SendDebugMouseEvent@16" (ByVal posX As Long, ByVal posY As Long, ByVal EvtType As Long, ByVal Button As Long)


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

Public Sub InitializeUI(ByVal Width As Long, ByVal Height As Long, ByVal pixelSize As Long)
On Error GoTo InitializeUI_Err
100 Dim initSuccess As Boolean
102 UITexture.Height = Height
104 UITexture.Width = Width
106 initSuccess = InitializeBabel(UITexture.Width, UITexture.Height)
108 UITexture.TextureHeight = NextPowerOf2(Height)
110 UITexture.TextureWidth = NextPowerOf2(Width)
112 ReDim UITexture.ImageBuffer(UITexture.Height * UITexture.Width * pixelSize)
114 UITexture.pixelSize = pixelSize
116 Set UITexture.Texture = SurfaceDB.CreateTexture(UITexture.TextureWidth, UITexture.TextureHeight)
118 BabelInitialized = True
    Call RegisterCallbacks(AddressOf LoginCB, AddressOf CloseClientCB)
    Exit Sub
InitializeUI_Err:
    Call RegistrarError(Err.Number, Err.Description, "BabelUI.InitializeUI", Erl)
End Sub

Public Sub InitializeInspectorUI(ByVal Width As Long, ByVal Height As Long)
On Error GoTo InitializeInspectorUI_Err
100 Dim initSuccess As Boolean
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
110         If .Texture Is Nothing Then
112             Call RegistrarError(102, "texture undefined ", "BabelUI.DrawTexture", 202)
                Exit Sub
            End If
116         Call SpriteBatch.SetTexture(.Texture)
118         Call SpriteBatch.SetAlpha(False)
120         Call SpriteBatch.Draw(0, 0, .Width, .Height, COLOR_WHITE, , , .Width / .TextureWidth, .Height / .TextureHeight, 0)
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
    Dim user, password As String
    If LoginValue.userLen > 0 Then
        user = GetStringFromPtr(LoginValue.user, LoginValue.userLen)
    End If
    If LoginValue.passwordLen > 0 Then
        password = GetStringFromPtr(LoginValue.password, LoginValue.passwordLen)
    End If
    'Call DoLogin(user, password, False)
    Call SendErrorMessage("test", 57, 509)
End Sub

Public Sub CloseClientCB()
    Call Closeclient
End Sub


