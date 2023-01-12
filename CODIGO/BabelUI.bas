Attribute VB_Name = "BabelUI"
Option Explicit

Public Declare Function InitializeBabel Lib "BabelUI.dll" Alias "_InitializeBabel@8" (ByVal Width As Long, ByVal Height As Long) As Boolean
Public Declare Function GetBebelImageBuffer Lib "BabelUI.dll" Alias "_GetImageBuffer@8" (ByRef Buffer As Byte, ByVal size As Long) As Boolean
Public Declare Sub BabelSendMouseEvent Lib "BabelUI.dll" Alias "_SendMouseEvent@16" (ByVal posX As Long, ByVal posY As Long, ByVal EvtType As Long, ByVal Button As Long)
Public Declare Sub BabelSendKeyEvent Lib "BabelUI.dll" Alias "_SendKeyEvent@20" (ByVal KeyCode As Integer, ByVal Shift As Boolean, ByVal EvtType As Long, ByVal CapsState As Boolean, ByVal Inspector As Boolean)
Public Declare Function NextPowerOf2 Lib "BabelUI.dll" Alias "_NextPowerOf2@4" (ByVal original As Long) As Long

'debug info
Public Declare Function CreateDebugWindow Lib "BabelUI.dll" Alias "_CreateDebugWindow@8" (ByVal Width As Long, ByVal Height As Long) As Boolean
Public Declare Function GetDebugImageBuffer Lib "BabelUI.dll" Alias "_GetDebugImageBuffer@8" (ByRef Buffer As Byte, ByVal size As Long) As Boolean
Public Declare Sub SendDebugMouseEvent Lib "BabelUI.dll" Alias "_SendDebugMouseEvent@16" (ByVal posX As Long, ByVal posY As Long, ByVal EvtType As Long, ByVal Button As Long)

Public Enum MouseEvent
    kType_MouseMoved
    kType_MouseDown
    kType_MouseUp
End Enum

Public Enum MouseButton
    kButton_None = 0
    kButton_Left
    kButton_Middle
    kButton_Right
End Enum

Public Enum KeyEventType
    kType_KeyDown
    kType_KeyUp
    kType_RawKeyDown
    kType_Char
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

