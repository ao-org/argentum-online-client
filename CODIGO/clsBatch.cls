VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
 
Private m_VBuffer As Direct3DVertexBuffer8
Private m_IBuffer As Direct3DIndexBuffer8
Private m_Texture As Direct3DTexture8
Private m_Alpha   As Boolean

'
' Represent the dynamic data of the batch
'
Private m_Data()   As TYPE_VERTEX
Private m_Position As Long

' Precalculated
Private m_Start As Long
Private m_Pre_Data() As TYPE_VERTEX
Private m_Fullscreen As Boolean

Private m_Pre_VBuffer As Direct3DVertexBuffer8
Private m_Pre_IBuffer As Direct3DIndexBuffer8

Private m_Pre_Data_Full() As TYPE_VERTEX

Private m_Pre_VBuffer_Full As Direct3DVertexBuffer8
Private m_Pre_IBuffer_Full As Direct3DIndexBuffer8

Private Const WATER_WAVE_SIZE_X = 4
Private Const WATER_WAVE_SIZE_Y = 4
Private Const WATER_WAVE_HEIGHT = 3
Private Const WATER_OBJECT_WAVE_HEIGHT = 3
Private Const WATER_WAVE_FACTOR_X = (2 * PI) / (WATER_WAVE_SIZE_X * TilePixelWidth)
Private Const WATER_WAVE_FACTOR_Y = (2 * PI) / (WATER_WAVE_SIZE_Y * TilePixelHeight)
Private Const WATER_TILE_FACTOR_X = TilePixelWidth * WATER_WAVE_FACTOR_X
Private Const WATER_TILE_FACTOR_Y = TilePixelHeight * WATER_WAVE_FACTOR_Y

Private Const LAVA_WAVE_SIZE_X = 4
Private Const LAVA_WAVE_SIZE_Y = 4
Private Const LAVA_WAVE_HEIGHT = 6
Private Const LAVA_WAVE_FACTOR_X = (2 * PI) / (LAVA_WAVE_SIZE_X * TilePixelWidth)
Private Const LAVA_WAVE_FACTOR_Y = (2 * PI) / (LAVA_WAVE_SIZE_Y * TilePixelHeight)
Private Const LAVA_TILE_FACTOR_X = TilePixelWidth * LAVA_WAVE_FACTOR_X
Private Const LAVA_TILE_FACTOR_Y = TilePixelHeight * LAVA_WAVE_FACTOR_Y

Public Sub Begin()
    '
    ' Re-Initialize the states of the batch
    '
    
    On Error GoTo Begin_Err
    
    Call DirectDevice.SetStreamSource(0, m_VBuffer, 24)
    Call DirectDevice.SetIndices(m_IBuffer, 0)

    Call SetAlpha(False)
    Call SetTexture(Nothing)
    
    Exit Sub

Begin_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsBatch.Begin", Erl)
    Resume Next
    
End Sub

Public Sub BeginPrecalculated(ByVal x As Single, ByVal y As Single)
    
    On Error GoTo BeginPrecalculated_Err
    
    Call SetAlpha(False)

    Dim Translation As D3DMATRIX
    Call D3DXMatrixTranslation(Translation, x, y, 0)
    Call DirectDevice.SetTransform(D3DTS_WORLD, Translation)
    
    Exit Sub

BeginPrecalculated_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsBatch.BeginPrecalculated", Erl)
    Resume Next
    
End Sub

Public Sub EndPrecalculated()
    
    On Error GoTo EndPrecalculated_Err
    
    Call Flush
    Call DirectDevice.SetTransform(D3DTS_WORLD, IdentityMatrix)
    
    Exit Sub

EndPrecalculated_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsBatch.EndPrecalculated", Erl)
    Resume Next
    
End Sub
 
Public Sub SetAlpha(ByVal IsAlpha As Boolean)
    
    On Error GoTo SetAlpha_Err
    
 
    If IsAlpha <> m_Alpha Then
        
        '
        ' Flush the batch since we're using a diferent state
        '
        Call Flush
 
        m_Alpha = IsAlpha
       
        '
        ' Update the new state of the batch
        '
        If IsAlpha Then
            Call DirectDevice.SetRenderState(D3DRS_SRCBLEND, D3DBLEND_SRCALPHA)
            Call DirectDevice.SetRenderState(D3DRS_DESTBLEND, D3DBLEND_ONE)
        Else
            Call DirectDevice.SetRenderState(D3DRS_SRCBLEND, D3DBLEND_SRCALPHA)
            Call DirectDevice.SetRenderState(D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA)
        End If
 
    End If

    
    Exit Sub

SetAlpha_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsBatch.SetAlpha", Erl)
    Resume Next
    
End Sub
 
Public Sub SetTexture(ByRef Texture As Direct3DTexture8)
    
    On Error GoTo SetTexture_Err
    
 
    If ObjPtr(m_Texture) <> ObjPtr(Texture) Then 'te cabe vb6
        
        '
        ' Flush the batch since we're using a diferent texture
        '
        Call Flush
       
        Set m_Texture = Texture
       
        '
        ' Update the new texture of the batch
        '
        Call DirectDevice.SetTexture(0, Texture)
 
    End If
 
    
    Exit Sub

SetTexture_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsBatch.SetTexture", Erl)
    Resume Next
    
End Sub
 
Friend Sub Draw(ByVal x As Long, _
                ByVal y As Long, _
                ByVal Width As Long, _
                ByVal Height As Long, _
                ByRef Color() As RGBA, _
                Optional ByVal Tx1 As Single = 0, _
                Optional ByVal Ty1 As Single = 0, _
                Optional ByVal Tx2 As Single = 1, _
                Optional ByVal Ty2 As Single = 1, _
                Optional ByVal Angle As Single = 0)
    
    On Error GoTo Draw_Err
    
               
    If Angle > 0 And Angle <= 360 Then
        Dim NewX     As Single
        Dim NewY     As Single
        Dim SinRad   As Single
        Dim CosRad   As Single
        Dim RadAngle As Single
        Dim CenterX  As Single
        Dim CenterY  As Single
       
        RadAngle = Angle * 0.017453292519
   
        CenterX = x + (Width * 0.5)
        CenterY = y + (Height * 0.5)
       
        SinRad = Sin(RadAngle)
        CosRad = Cos(RadAngle)
       
    End If
   
    m_Data(m_Position).x = x
    m_Data(m_Position).y = y + Height
    m_Data(m_Position).z = 0
    m_Data(m_Position).Color = Color(0)
    m_Data(m_Position).tX = Tx1
    m_Data(m_Position).tY = Ty2
 
    If Angle > 0 And Angle <= 360 Then
   
        NewX = CenterX + (m_Data(m_Position).x - CenterX) * CosRad - (m_Data(m_Position).y - CenterY) * SinRad
        NewY = CenterY + (m_Data(m_Position).y - CenterY) * CosRad + (m_Data(m_Position).x - CenterX) * SinRad
       
        m_Data(m_Position).x = NewX
        m_Data(m_Position).y = NewY
       
    End If
   
    m_Position = m_Position + 1
    
    m_Data(m_Position).x = x
    m_Data(m_Position).y = y
    m_Data(m_Position).z = 0
    m_Data(m_Position).Color = Color(1)
    m_Data(m_Position).tX = Tx1
    m_Data(m_Position).tY = Ty1
   
    If Angle > 0 And Angle <= 360 Then
   
        NewX = CenterX + (m_Data(m_Position).x - CenterX) * CosRad - (m_Data(m_Position).y - CenterY) * SinRad
        NewY = CenterY + (m_Data(m_Position).y - CenterY) * CosRad + (m_Data(m_Position).x - CenterX) * SinRad
       
        m_Data(m_Position).x = NewX
        m_Data(m_Position).y = NewY
       
    End If
   
    m_Position = m_Position + 1
   
    m_Data(m_Position).x = x + Width
    m_Data(m_Position).y = y + Height
    m_Data(m_Position).z = 0
    m_Data(m_Position).Color = Color(2)
    m_Data(m_Position).tX = Tx2
    m_Data(m_Position).tY = Ty2
 
    If Angle > 0 And Angle <= 360 Then
   
        NewX = CenterX + (m_Data(m_Position).x - CenterX) * CosRad - (m_Data(m_Position).y - CenterY) * SinRad
        NewY = CenterY + (m_Data(m_Position).y - CenterY) * CosRad + (m_Data(m_Position).x - CenterX) * SinRad
       
        m_Data(m_Position).x = NewX
        m_Data(m_Position).y = NewY
       
    End If
   
    m_Position = m_Position + 1
   
    m_Data(m_Position).x = x + Width
    m_Data(m_Position).y = y
    m_Data(m_Position).z = 0
    m_Data(m_Position).Color = Color(3)
    m_Data(m_Position).tX = Tx2
    m_Data(m_Position).tY = Ty1
 
    If Angle > 0 And Angle <= 360 Then
   
        NewX = CenterX + (m_Data(m_Position).x - CenterX) * CosRad - (m_Data(m_Position).y - CenterY) * SinRad
        NewY = CenterY + (m_Data(m_Position).y - CenterY) * CosRad + (m_Data(m_Position).x - CenterX) * SinRad
       
        m_Data(m_Position).x = NewX
        m_Data(m_Position).y = NewY
       
    End If
   
    m_Position = m_Position + 1
   
    If m_Position >= UBound(m_Data) Then
        Call Flush
    End If
   
    
    Exit Sub

Draw_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsBatch.Draw", Erl)
    Resume Next
    
End Sub

Friend Sub DrawTriangle(ByVal x1 As Long, _
                        ByVal y1 As Long, _
                        ByVal x2 As Long, _
                        ByVal y2 As Long, _
                        ByVal x3 As Long, _
                        ByVal y3 As Long, _
                        ByRef Color() As RGBA)

    m_Data(m_Position).x = x1
    m_Data(m_Position).y = y1
    m_Data(m_Position).z = 0
    m_Data(m_Position).Color = Color(0)
    m_Data(m_Position).tX = 0
    m_Data(m_Position).tY = 0

    m_Position = m_Position + 1
   
    m_Data(m_Position).x = x2
    m_Data(m_Position).y = y2
    m_Data(m_Position).z = 0
    m_Data(m_Position).Color = Color(1)
    m_Data(m_Position).tX = 0
    m_Data(m_Position).tY = 0
  
    m_Position = m_Position + 1
  
    m_Data(m_Position).x = x3
    m_Data(m_Position).y = y3
    m_Data(m_Position).z = 0
    m_Data(m_Position).Color = Color(2)
    m_Data(m_Position).tX = 0
    m_Data(m_Position).tY = 0

    m_Position = m_Position + 1
  
    m_Data(m_Position).x = x3
    m_Data(m_Position).y = y3
    m_Data(m_Position).z = 0
    m_Data(m_Position).Color = Color(3)
    m_Data(m_Position).tX = 0
    m_Data(m_Position).tY = 0

    m_Position = m_Position + 1
End Sub

Friend Sub DrawLoad(ByVal x As Long, _
                ByVal y As Long, _
                ByVal Width As Long, _
                ByVal Height As Long, _
                ByRef Color() As RGBA, _
                ByVal Angle As Single, _
                Optional ByVal Tx1 As Single = 0, _
                Optional ByVal Ty1 As Single = 0, _
                Optional ByVal Tx2 As Single = 1, _
                Optional ByVal Ty2 As Single = 1)
    
    On Error GoTo Draw_Err
    
    Dim AngleR As Double
    
    
    AngleR = (Angle - Int(Angle / 45) * 45) '* PI / 180
    
    Dim W As Long
    Dim H As Long
    
    W = Width / 2
    H = Height / 2
    
    
    If Angle < 45 Then Call DrawTriangle(x, y, x + AngleR * W / 45, y - H, x + W, y - H, Color)
    If Angle < 90 Then Call DrawTriangle(x, y, x + W, y - H + IIf(Angle >= 45, AngleR * H / 45, 0), x + W, y, Color)
    If Angle < 135 Then Call DrawTriangle(x, y, x + W, y + IIf(Angle >= 90, AngleR * H / 45, 0), x + W, y + H, Color)
    If Angle < 180 Then Call DrawTriangle(x, y, x + W - IIf(Angle >= 135, AngleR * W / 45, 0), y + H, x, y + H, Color)
    If Angle < 225 Then Call DrawTriangle(x, y, x - IIf(Angle >= 180, AngleR * W / 45, 0), y + H, x - W, y + H, Color)
    If Angle < 270 Then Call DrawTriangle(x, y, x - W, y + H - IIf(Angle >= 225, AngleR * H / 45, 0), x - W, y, Color)
    If Angle < 315 Then Call DrawTriangle(x, y, x - W, y - IIf(Angle >= 270, AngleR * H / 45, 0), x - W, y - H, Color)
    If Angle < 360 Then Call DrawTriangle(x, y, x - W + IIf(Angle >= 315, AngleR * H / 45, 0), y - H, x, y - H, Color)
   
    If m_Position >= UBound(m_Data) Then
        Call Flush
    End If
   
    
    Exit Sub

Draw_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsBatch.Draw", Erl)
    Resume Next
    
End Sub


Friend Sub DrawShadow(ByVal x As Long, _
                    ByVal y As Long, _
                    ByVal Width As Long, _
                    ByVal Height As Long, _
                    ByRef ColorLight() As RGBA, _
                    Optional ByVal Tx1 As Single = 0, _
                    Optional ByVal Ty1 As Single = 0, _
                    Optional ByVal Tx2 As Single = 1, _
                    Optional ByVal Ty2 As Single = 1)
    
    On Error GoTo DrawShadow_Err
    
                
    Dim Color(3) As RGBA

    Dim i As Integer
    For i = 0 To 3
        Color(i).A = (0.2126 * ColorLight(i).r + 0.7152 * ColorLight(i).G + 0.0722 * ColorLight(i).B) ^ 2 * 0.000625
    Next

    Dim dX As Single, dy As Single

    For i = 0 To 3
        Select Case i
            Case 0: dX = x: dy = y - 1
            Case 1: dX = x + 1: dy = y
            Case 2: dX = x: dy = y + 1
            Case 3: dX = x - 1: dy = y
        End Select
    
        m_Data(m_Position).x = dX
        m_Data(m_Position).y = dy + Height - 2
        m_Data(m_Position).z = 0
        m_Data(m_Position).Color = Color(0)
        m_Data(m_Position).tX = Tx1
        m_Data(m_Position).tY = Ty2
     
        m_Position = m_Position + 1
        
        m_Data(m_Position).x = dX + Height * 0.25
        m_Data(m_Position).y = dy + Height * 0.25 - 2
        m_Data(m_Position).z = 0
        m_Data(m_Position).Color = Color(1)
        m_Data(m_Position).tX = Tx1
        m_Data(m_Position).tY = Ty1
       
        m_Position = m_Position + 1
       
        m_Data(m_Position).x = dX + Width
        m_Data(m_Position).y = dy + Height - 2
        m_Data(m_Position).z = 0
        m_Data(m_Position).Color = Color(2)
        m_Data(m_Position).tX = Tx2
        m_Data(m_Position).tY = Ty2
    
        m_Position = m_Position + 1
       
        m_Data(m_Position).x = dX + Width + Height * 0.25
        m_Data(m_Position).y = dy + Height * 0.25 - 2
        m_Data(m_Position).z = 0
        m_Data(m_Position).Color = Color(3)
        m_Data(m_Position).tX = Tx2
        m_Data(m_Position).tY = Ty1
    
        m_Position = m_Position + 1
       
        If m_Position >= UBound(m_Data) Then
            Call Flush
        End If
    Next
   
    
    Exit Sub

DrawShadow_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsBatch.DrawShadow", Erl)
    Resume Next
    
End Sub

Friend Sub DrawReflection(ByVal x As Long, _
                        ByVal y As Long, _
                        ByVal Width As Long, _
                        ByVal Height As Long, _
                        ByRef ColorLight() As RGBA, _
                        Optional ByVal Tx1 As Single = 0, _
                        Optional ByVal Ty1 As Single = 0, _
                        Optional ByVal Tx2 As Single = 1, _
                        Optional ByVal Ty2 As Single = 1)
    
    On Error GoTo DrawReflection_Err
    

    Dim Color(3) As RGBA

    Dim i As Integer
    For i = 0 To 3
        Color(i).A = ColorLight(i).A * 0.5
        Color(i).r = ColorLight(i).r * 0.2
        Color(i).G = ColorLight(i).G * 0.5
        Color(i).B = ColorLight(i).B
    Next
    
    Dim dt As Single, dX As Long, dy As Long
    dt = (FrameTime Mod 10000) * 0.0006283
    dX = Sin(dt) * 8
    dy = Sin(dt * 2 + 1) * 5

    m_Data(m_Position).x = x + dX
    m_Data(m_Position).y = y + Height * 1.75 - dy
    m_Data(m_Position).z = 0.5
    m_Data(m_Position).Color = Color(0)
    m_Data(m_Position).tX = Tx1
    m_Data(m_Position).tY = Ty1
 
    m_Position = m_Position + 1
    
    m_Data(m_Position).x = x
    m_Data(m_Position).y = y + Height - 5
    m_Data(m_Position).z = 0.5
    m_Data(m_Position).Color = Color(1)
    m_Data(m_Position).tX = Tx1
    m_Data(m_Position).tY = Ty2
   
    m_Position = m_Position + 1
   
    m_Data(m_Position).x = x + Width + dX
    m_Data(m_Position).y = y + Height * 1.75 - dy
    m_Data(m_Position).z = 0.5
    m_Data(m_Position).Color = Color(2)
    m_Data(m_Position).tX = Tx2
    m_Data(m_Position).tY = Ty1

    m_Position = m_Position + 1
   
    m_Data(m_Position).x = x + Width
    m_Data(m_Position).y = y + Height - 5
    m_Data(m_Position).z = 0.5
    m_Data(m_Position).Color = Color(3)
    m_Data(m_Position).tX = Tx2
    m_Data(m_Position).tY = Ty2

    m_Position = m_Position + 1
   
    If m_Position >= UBound(m_Data) Then
        Call Flush
    End If
   
    
    Exit Sub

DrawReflection_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsBatch.DrawReflection", Erl)
    Resume Next
    
End Sub

Friend Sub DrawBreathing(ByVal x As Long, _
                        ByVal y As Long, _
                        ByVal Width As Long, _
                        ByVal Height As Long, _
                        ByVal ease As Single, _
                        ByRef Color() As RGBA, _
                        Optional ByVal Tx1 As Single = 0, _
                        Optional ByVal Ty1 As Single = 0, _
                        Optional ByVal Tx2 As Single = 1, _
                        Optional ByVal Ty2 As Single = 1)
    
    On Error GoTo DrawBreathing_Err
    

    Dim dy As Integer
    dy = Int(2 * ease)

    m_Data(m_Position).x = x + 0.25
    m_Data(m_Position).y = y + Height
    m_Data(m_Position).z = 0
    m_Data(m_Position).Color = Color(0)
    m_Data(m_Position).tX = Tx1
    m_Data(m_Position).tY = Ty2
 
    m_Position = m_Position + 1
    
    m_Data(m_Position).x = x + 0.25
    m_Data(m_Position).y = y - dy
    m_Data(m_Position).z = 0
    m_Data(m_Position).Color = Color(1)
    m_Data(m_Position).tX = Tx1
    m_Data(m_Position).tY = Ty1
   
    m_Position = m_Position + 1
   
    m_Data(m_Position).x = x + Width + 0.25
    m_Data(m_Position).y = y + Height
    m_Data(m_Position).z = 0
    m_Data(m_Position).Color = Color(2)
    m_Data(m_Position).tX = Tx2
    m_Data(m_Position).tY = Ty2

    m_Position = m_Position + 1
   
    m_Data(m_Position).x = x + Width + 0.25
    m_Data(m_Position).y = y - dy
    m_Data(m_Position).z = 0
    m_Data(m_Position).Color = Color(3)
    m_Data(m_Position).tX = Tx2
    m_Data(m_Position).tY = Ty1

    m_Position = m_Position + 1
   
    If m_Position >= UBound(m_Data) Then
        Call Flush
    End If
   
    
    Exit Sub

DrawBreathing_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsBatch.DrawBreathing", Erl)
    Resume Next
    
End Sub

Friend Sub DrawWater(ByVal x As Single, _
                    ByVal y As Single, _
                    ByVal Width As Long, _
                    ByVal Height As Long, _
                    ByRef Color() As RGBA, _
                    ByVal Tx1 As Single, _
                    ByVal Ty1 As Single, _
                    ByVal Tx2 As Single, _
                    ByVal Ty2 As Single, _
                    ByVal MapX As Integer, _
                    ByVal MapY As Integer, _
                    ByVal Top As Byte, _
                    ByVal Right As Byte, _
                    ByVal Bottom As Byte, _
                    ByVal Left As Byte)
    
    On Error GoTo DrawWater_Err
    

    Dim t As Single
    t = (FrameTime Mod 62831) * 0.001
    
    Dim tX As Single, tY As Single
    tX = x * WATER_WAVE_FACTOR_X
    tY = y * WATER_WAVE_FACTOR_Y

    m_Data(m_Position).x = x + Cos(t + tX + tY + WATER_TILE_FACTOR_Y) * WATER_WAVE_HEIGHT * Left
    m_Data(m_Position).y = y + Height + Sin(t + tX + tY + WATER_TILE_FACTOR_Y) * WATER_WAVE_HEIGHT * Bottom
    m_Data(m_Position).z = 1
    m_Data(m_Position).Color = Color(0)
    m_Data(m_Position).tX = Tx1
    m_Data(m_Position).tY = Ty2

    m_Position = m_Position + 1
    
    m_Data(m_Position).x = x + Cos(t + tX + tY) * WATER_WAVE_HEIGHT * Left
    m_Data(m_Position).y = y + Sin(t + tX + tY) * WATER_WAVE_HEIGHT * Top
    m_Data(m_Position).z = 1
    m_Data(m_Position).Color = Color(1)
    m_Data(m_Position).tX = Tx1
    m_Data(m_Position).tY = Ty1
   
    m_Position = m_Position + 1
   
    m_Data(m_Position).x = x + Width + Cos(t + tX + tY + WATER_TILE_FACTOR_X + WATER_TILE_FACTOR_Y) * WATER_WAVE_HEIGHT * Right
    m_Data(m_Position).y = y + Height + Sin(t + tX + tY + WATER_TILE_FACTOR_X + WATER_TILE_FACTOR_Y) * WATER_WAVE_HEIGHT * Bottom
    m_Data(m_Position).z = 1
    m_Data(m_Position).Color = Color(2)
    m_Data(m_Position).tX = Tx2
    m_Data(m_Position).tY = Ty2

    m_Position = m_Position + 1
   
    m_Data(m_Position).x = x + Width + Cos(t + tX + tY + WATER_TILE_FACTOR_X) * WATER_WAVE_HEIGHT * Right
    m_Data(m_Position).y = y + Sin(t + tX + tY + WATER_TILE_FACTOR_X) * WATER_WAVE_HEIGHT * Top
    m_Data(m_Position).z = 1
    m_Data(m_Position).Color = Color(3)
    m_Data(m_Position).tX = Tx2
    m_Data(m_Position).tY = Ty1

    m_Position = m_Position + 1
   
    If m_Position >= UBound(m_Data) Then
        Call Flush
    End If
   
    
    Exit Sub

DrawWater_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsBatch.DrawWater", Erl)
    Resume Next
    
End Sub

Friend Sub DrawItemInWater(ByVal x As Single, _
                    ByVal y As Single, _
                    ByVal Width As Long, _
                    ByVal Height As Long, _
                    ByRef Color() As RGBA, _
                    ByVal Tx1 As Single, _
                    ByVal Ty1 As Single, _
                    ByVal Tx2 As Single, _
                    ByVal Ty2 As Single, _
                    ByVal Angle As Single)
    
    On Error GoTo DrawWater_Err
    
    Dim t As Double
    t = (FrameTime Mod 62831) * 0.001
    
    If Angle > 0 And Angle <= 360 Then
        Dim NewX     As Single
        Dim NewY     As Single
        Dim SinRad   As Single
        Dim CosRad   As Single
        Dim RadAngle As Single
        Dim CenterX  As Single
        Dim CenterY  As Single
       
        RadAngle = Angle * 0.017453292519
   
        CenterX = x + (Width * 0.5)
        CenterY = y + (Height * 0.5)
       
        SinRad = Sin(RadAngle)
        CosRad = Cos(RadAngle)
       
    End If
    
    m_Data(m_Position).x = x + Cos(t + WATER_TILE_FACTOR_Y) * WATER_OBJECT_WAVE_HEIGHT
    m_Data(m_Position).y = y + Height + Sin(t + WATER_TILE_FACTOR_Y) * WATER_OBJECT_WAVE_HEIGHT
    m_Data(m_Position).z = 1
    m_Data(m_Position).Color = Color(0)
    m_Data(m_Position).tX = Tx1
    m_Data(m_Position).tY = Ty2

    If Angle > 0 And Angle <= 360 Then
   
        NewX = CenterX + (m_Data(m_Position).x - CenterX) * CosRad - (m_Data(m_Position).y - CenterY) * SinRad
        NewY = CenterY + (m_Data(m_Position).y - CenterY) * CosRad + (m_Data(m_Position).x - CenterX) * SinRad
       
        m_Data(m_Position).x = NewX
        m_Data(m_Position).y = NewY
       
    End If

    m_Position = m_Position + 1
    
    m_Data(m_Position).x = x + Cos(t) * WATER_OBJECT_WAVE_HEIGHT
    m_Data(m_Position).y = y + Sin(t) * WATER_OBJECT_WAVE_HEIGHT
    m_Data(m_Position).z = 1
    m_Data(m_Position).Color = Color(1)
    m_Data(m_Position).tX = Tx1
    m_Data(m_Position).tY = Ty1
   
    If Angle > 0 And Angle <= 360 Then
   
        NewX = CenterX + (m_Data(m_Position).x - CenterX) * CosRad - (m_Data(m_Position).y - CenterY) * SinRad
        NewY = CenterY + (m_Data(m_Position).y - CenterY) * CosRad + (m_Data(m_Position).x - CenterX) * SinRad
       
        m_Data(m_Position).x = NewX
        m_Data(m_Position).y = NewY
       
    End If
   
    m_Position = m_Position + 1
   
    m_Data(m_Position).x = x + Width + Cos(t + WATER_TILE_FACTOR_X + WATER_TILE_FACTOR_Y) * WATER_OBJECT_WAVE_HEIGHT
    m_Data(m_Position).y = y + Height + Sin(t + WATER_TILE_FACTOR_Y) * WATER_OBJECT_WAVE_HEIGHT
    m_Data(m_Position).z = 1
    m_Data(m_Position).Color = Color(2)
    m_Data(m_Position).tX = Tx2
    m_Data(m_Position).tY = Ty2

    If Angle > 0 And Angle <= 360 Then
   
        NewX = CenterX + (m_Data(m_Position).x - CenterX) * CosRad - (m_Data(m_Position).y - CenterY) * SinRad
        NewY = CenterY + (m_Data(m_Position).y - CenterY) * CosRad + (m_Data(m_Position).x - CenterX) * SinRad
       
        m_Data(m_Position).x = NewX
        m_Data(m_Position).y = NewY
       
    End If

    m_Position = m_Position + 1
   
    m_Data(m_Position).x = x + Width + Cos(t + WATER_TILE_FACTOR_X) * WATER_OBJECT_WAVE_HEIGHT
    m_Data(m_Position).y = y + Sin(t) * WATER_OBJECT_WAVE_HEIGHT
    m_Data(m_Position).z = 1
    m_Data(m_Position).Color = Color(3)
    m_Data(m_Position).tX = Tx2
    m_Data(m_Position).tY = Ty1

    If Angle > 0 And Angle <= 360 Then
   
        NewX = CenterX + (m_Data(m_Position).x - CenterX) * CosRad - (m_Data(m_Position).y - CenterY) * SinRad
        NewY = CenterY + (m_Data(m_Position).y - CenterY) * CosRad + (m_Data(m_Position).x - CenterX) * SinRad
       
        m_Data(m_Position).x = NewX
        m_Data(m_Position).y = NewY
       
    End If

    m_Position = m_Position + 1
   
    If m_Position >= UBound(m_Data) Then
        Call Flush
    End If
   
    
    Exit Sub

DrawWater_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsBatch.DrawWater", Erl)
    Resume Next
    
End Sub

Friend Sub DrawLava(ByVal x As Single, _
                    ByVal y As Single, _
                    ByVal Width As Long, _
                    ByVal Height As Long, _
                    ByRef Color() As RGBA, _
                    ByVal Tx1 As Single, _
                    ByVal Ty1 As Single, _
                    ByVal Tx2 As Single, _
                    ByVal Ty2 As Single, _
                    ByVal MapX As Integer, _
                    ByVal MapY As Integer, _
                    ByVal Top As Byte, _
                    ByVal Right As Byte, _
                    ByVal Bottom As Byte, _
                    ByVal Left As Byte)
    
    On Error GoTo DrawLava_Err
    

    Dim t As Single
    t = (FrameTime Mod 62831) * 0.0005
    
    Dim tX As Single, tY As Single
    tX = x * LAVA_WAVE_FACTOR_X
    tY = y * LAVA_WAVE_FACTOR_Y

    m_Data(m_Position).x = x + Cos(t + tX + tY + LAVA_TILE_FACTOR_Y) * LAVA_WAVE_HEIGHT * Left
    m_Data(m_Position).y = y + Height + Sin(t + tX + tY + LAVA_TILE_FACTOR_Y) * LAVA_WAVE_HEIGHT * Bottom
    m_Data(m_Position).z = 1
    m_Data(m_Position).Color = Color(0)
    m_Data(m_Position).tX = Tx1
    m_Data(m_Position).tY = Ty2

    m_Position = m_Position + 1
    
    m_Data(m_Position).x = x + Cos(t + tX + tY) * LAVA_WAVE_HEIGHT * Left
    m_Data(m_Position).y = y + Sin(t + tX + tY) * LAVA_WAVE_HEIGHT * Top
    m_Data(m_Position).z = 1
    m_Data(m_Position).Color = Color(1)
    m_Data(m_Position).tX = Tx1
    m_Data(m_Position).tY = Ty1
   
    m_Position = m_Position + 1
   
    m_Data(m_Position).x = x + Width + Cos(t + tX + tY + LAVA_TILE_FACTOR_X + LAVA_TILE_FACTOR_Y) * LAVA_WAVE_HEIGHT * Right
    m_Data(m_Position).y = y + Height + Sin(t + tX + tY + LAVA_TILE_FACTOR_X + LAVA_TILE_FACTOR_Y) * LAVA_WAVE_HEIGHT * Bottom
    m_Data(m_Position).z = 1
    m_Data(m_Position).Color = Color(2)
    m_Data(m_Position).tX = Tx2
    m_Data(m_Position).tY = Ty2

    m_Position = m_Position + 1
   
    m_Data(m_Position).x = x + Width + Cos(t + tX + tY + LAVA_TILE_FACTOR_X) * LAVA_WAVE_HEIGHT * Right
    m_Data(m_Position).y = y + Sin(t + tX + tY + LAVA_TILE_FACTOR_X) * LAVA_WAVE_HEIGHT * Top
    m_Data(m_Position).z = 1
    m_Data(m_Position).Color = Color(3)
    m_Data(m_Position).tX = Tx2
    m_Data(m_Position).tY = Ty1

    m_Position = m_Position + 1
   
    If m_Position >= UBound(m_Data) Then
        Call Flush
    End If
   
    
    Exit Sub

DrawLava_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsBatch.DrawLava", Erl)
    Resume Next
    
End Sub

Friend Sub DrawVertices(vertices() As TYPE_VERTEX)
    
    On Error GoTo DrawVertices_Err
    

    Dim i As Long

    For i = LBound(vertices) To UBound(vertices)
        m_Data(m_Position) = vertices(i)
        m_Position = m_Position + 1

        If m_Position >= UBound(m_Data) Then
            Call Flush
        End If
    Next

    
    Exit Sub

DrawVertices_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsBatch.DrawVertices", Erl)
    Resume Next
    
End Sub
 
Public Sub Finish()
    
    On Error GoTo Finish_Err
    

    '
    ' Flush the batch since we're finished
    '
    Call Flush
 
    
    Exit Sub

Finish_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsBatch.Finish", Erl)
    Resume Next
    
End Sub
 
Public Sub Flush()
    
    On Error GoTo Flush_Err
    
 
    '
    ' Only draw if there is any data in the buffer
    '
    If m_Position > 0 Then
        
        '
        ' Set the new data of the vertex buffer
        '
        Call D3DVertexBuffer8SetData(m_VBuffer, 0, m_Position * LenB(m_Data(0)), D3DLOCK_DISCARD, m_Data(0))
   
        Call DirectDevice.DrawIndexedPrimitive(D3DPT_TRIANGLELIST, 0, (m_Position / 2) * 3, 0, m_Position / 2)
 
        '
        ' We start all over again
        '
        m_Position = 0
 
    End If
 
    
    Exit Sub

Flush_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsBatch.Flush", Erl)
    Resume Next
    
End Sub
 
Public Sub Initialize(ByVal Capacity As Long)
    
    On Error GoTo Initialize_Err
    

    ReDim m_Data(Capacity - 1) As TYPE_VERTEX
   
    '
    '  Create the vertice buffer
    '
    Set m_VBuffer = DirectDevice.CreateVertexBuffer(24 * Capacity, D3DUSAGE_DYNAMIC, D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_TEX1, D3DPOOL_DEFAULT)
 
    '
    '  Create the indice buffer, fill it with pre-baked indices
    '
    Set m_IBuffer = DirectDevice.CreateIndexBuffer(3 * Capacity, D3DUSAGE_WRITEONLY, D3DFMT_INDEX16, D3DPOOL_DEFAULT)
   
    Dim lpIndices() As Integer
   
    ReDim lpIndices(Capacity * 3 / 2 - 1) As Integer
   
    Dim i As Long, j As Integer
   
    For i = 0 To UBound(lpIndices) Step 6
        lpIndices(i) = j
        lpIndices(i + 1) = j + 1
        lpIndices(i + 2) = j + 2
        lpIndices(i + 3) = j + 2
        lpIndices(i + 4) = j + 3
        lpIndices(i + 5) = j + 1
       
        j = j + 4
    Next
   
    Call D3DIndexBuffer8SetData(m_IBuffer, 0, 3 * Capacity, 0, lpIndices(0))

    
    Exit Sub

Initialize_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsBatch.Initialize", Erl)
    Resume Next
    
End Sub
