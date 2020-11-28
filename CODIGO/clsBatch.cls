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

Public Sub Begin()
    '
    ' Re-Initialize the states of the batch
    '
    Call DirectDevice.SetTexture(0, Nothing)
    Call DirectDevice.SetStreamSource(0, m_VBuffer, 24)
    Call DirectDevice.SetIndices(m_IBuffer, 0)

    Call SetAlpha(False)

End Sub

Public Sub BeginPrecalculated(ByVal x As Single, ByVal y As Single)
    Call SetAlpha(False)

    Dim Translation As D3DMATRIX
    Call D3DXMatrixTranslation(Translation, x, y, 0)
    Call DirectDevice.SetTransform(D3DTS_WORLD, Translation)
End Sub

Public Sub EndPrecalculated()
    Call Flush
    Call DirectDevice.SetTransform(D3DTS_WORLD, IdentityMatrix)
End Sub
 
Public Sub SetAlpha(ByVal IsAlpha As Boolean)
 
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
            Call DirectDevice.SetRenderState(D3DRS_SRCBLEND, D3DBLEND_ONE)
            Call DirectDevice.SetRenderState(D3DRS_DESTBLEND, D3DBLEND_ONE)
        Else
            Call DirectDevice.SetRenderState(D3DRS_SRCBLEND, D3DBLEND_SRCALPHA)
            Call DirectDevice.SetRenderState(D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA)
        End If
 
    End If

End Sub
 
Public Sub SetTexture(ByRef Texture As Direct3DTexture8)
 
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
 
End Sub
 
Public Sub Draw(ByVal x As Long, _
                ByVal y As Long, _
                ByVal Width As Long, _
                ByVal Height As Long, _
                ByRef color_light() As Long, _
                Optional ByVal Tx1 As Single = 0, _
                Optional ByVal Ty1 As Single = 0, _
                Optional ByVal Tx2 As Single = 1, _
                Optional ByVal Ty2 As Single = 1, _
                Optional ByVal angle As Single = 0, _
                Optional ByVal z As Single = 0)
                
    Dim Color(3) As Long
                
    Dim i As Integer
    For i = 0 To 3
        Color(i) = color_light(i)
        If Color(i) = 0 Then Color(i) = map_base_light
    Next
               
    If angle > 0 And angle <= 360 Then
        Dim NewX     As Single
        Dim NewY     As Single
        Dim SinRad   As Single
        Dim CosRad   As Single
        Dim RadAngle As Single
        Dim CenterX  As Single
        Dim CenterY  As Single
       
        RadAngle = angle * 0.017453292519
   
        CenterX = x + (Width * 0.5)
        CenterY = y + (Height * 0.5)
       
        SinRad = Sin(RadAngle)
        CosRad = Cos(RadAngle)
       
    End If
   
    m_Data(m_Position).x = x
    m_Data(m_Position).y = y + Height
    m_Data(m_Position).z = z
    m_Data(m_Position).Color = Color(0)
    m_Data(m_Position).TX = Tx1
    m_Data(m_Position).TY = Ty2
 
    If angle > 0 And angle <= 360 Then
   
        NewX = CenterX + (m_Data(m_Position).x - CenterX) * -CosRad - (m_Data(m_Position).y - CenterY) * -SinRad
        NewY = CenterY + (m_Data(m_Position).y - CenterY) * -CosRad + (m_Data(m_Position).x - CenterX) * -SinRad
       
        m_Data(m_Position).x = NewX
        m_Data(m_Position).y = NewY
       
    End If
   
    m_Position = m_Position + 1
    
    m_Data(m_Position).x = x
    m_Data(m_Position).y = y
    m_Data(m_Position).z = z
    m_Data(m_Position).Color = Color(1)
    m_Data(m_Position).TX = Tx1
    m_Data(m_Position).TY = Ty1
   
    If angle > 0 And angle <= 360 Then
   
        NewX = CenterX + (m_Data(m_Position).x - CenterX) * -CosRad - (m_Data(m_Position).y - CenterY) * -SinRad
        NewY = CenterY + (m_Data(m_Position).y - CenterY) * -CosRad + (m_Data(m_Position).x - CenterX) * -SinRad
       
        m_Data(m_Position).x = NewX
        m_Data(m_Position).y = NewY
       
    End If
   
    m_Position = m_Position + 1
   
    m_Data(m_Position).x = x + Width
    m_Data(m_Position).y = y + Height
    m_Data(m_Position).z = z
    m_Data(m_Position).Color = Color(2)
    m_Data(m_Position).TX = Tx2
    m_Data(m_Position).TY = Ty2
 
    If angle > 0 And angle <= 360 Then
   
        NewX = CenterX + (m_Data(m_Position).x - CenterX) * -CosRad - (m_Data(m_Position).y - CenterY) * -SinRad
        NewY = CenterY + (m_Data(m_Position).y - CenterY) * -CosRad + (m_Data(m_Position).x - CenterX) * -SinRad
       
        m_Data(m_Position).x = NewX
        m_Data(m_Position).y = NewY
       
    End If
   
    m_Position = m_Position + 1
   
    m_Data(m_Position).x = x + Width
    m_Data(m_Position).y = y
    m_Data(m_Position).z = z
    m_Data(m_Position).Color = Color(3)
    m_Data(m_Position).TX = Tx2
    m_Data(m_Position).TY = Ty1
 
    If angle > 0 And angle <= 360 Then
   
        NewX = CenterX + (m_Data(m_Position).x - CenterX) * -CosRad - (m_Data(m_Position).y - CenterY) * -SinRad
        NewY = CenterY + (m_Data(m_Position).y - CenterY) * -CosRad + (m_Data(m_Position).x - CenterX) * -SinRad
       
        m_Data(m_Position).x = NewX
        m_Data(m_Position).y = NewY
       
    End If
   
    m_Position = m_Position + 1
   
    If m_Position >= UBound(m_Data) Then
        Call Flush
    End If
   
End Sub

Public Sub DrawShadow(ByVal x As Long, _
                ByVal y As Long, _
                ByVal Width As Long, _
                ByVal Height As Long, _
                ByRef Color() As Long, _
                Optional ByVal Tx1 As Single = 0, _
                Optional ByVal Ty1 As Single = 0, _
                Optional ByVal Tx2 As Single = 1, _
                Optional ByVal Ty2 As Single = 1)

    Dim ColorShadow(3) As Long
    Dim i As Integer

    For i = 0 To 3
        If Color(i) = 0 Then Color(i) = map_base_light
        
        Dim IntensidadSombra As Single, tmpColor As D3DCOLORVALUE
        
        Call ARGBtoD3DCOLORVALUE(Color(0), tmpColor)
        
        IntensidadSombra = (0.2126 * tmpColor.r + 0.7152 * tmpColor.g + 0.0722 * tmpColor.b) ^ 2 * 0.000625

        ColorShadow(i) = D3DColorARGB(IntensidadSombra, 0, 0, 0)
    Next
    
    Dim dx As Single, dy As Single
   
    For i = 0 To 3
        Select Case i
            Case 0: dx = x: dy = y - 1
            Case 1: dx = x + 1: dy = y
            Case 2: dx = x: dy = y + 1
            Case 3: dx = x - 1: dy = y
        End Select
    
        m_Data(m_Position).x = dx
        m_Data(m_Position).y = dy + Height - 2
        m_Data(m_Position).z = 0
        m_Data(m_Position).Color = ColorShadow(0)
        m_Data(m_Position).TX = Tx1
        m_Data(m_Position).TY = Ty2
     
        m_Position = m_Position + 1
        
        m_Data(m_Position).x = dx + Height * 0.25
        m_Data(m_Position).y = dy + Height * 0.25 - 2
        m_Data(m_Position).z = 0
        m_Data(m_Position).Color = ColorShadow(1)
        m_Data(m_Position).TX = Tx1
        m_Data(m_Position).TY = Ty1
       
        m_Position = m_Position + 1
       
        m_Data(m_Position).x = dx + Width
        m_Data(m_Position).y = dy + Height - 2
        m_Data(m_Position).z = 0
        m_Data(m_Position).Color = ColorShadow(2)
        m_Data(m_Position).TX = Tx2
        m_Data(m_Position).TY = Ty2
    
        m_Position = m_Position + 1
       
        m_Data(m_Position).x = dx + Width + Height * 0.25
        m_Data(m_Position).y = dy + Height * 0.25 - 2
        m_Data(m_Position).z = 0
        m_Data(m_Position).Color = ColorShadow(3)
        m_Data(m_Position).TX = Tx2
        m_Data(m_Position).TY = Ty1
    
        m_Position = m_Position + 1
       
        If m_Position >= UBound(m_Data) Then
            Call Flush
        End If
    Next
   
End Sub

Public Sub DrawReflection(ByVal x As Long, _
                ByVal y As Long, _
                ByVal Width As Long, _
                ByVal Height As Long, _
                ByRef Color() As Long, _
                Optional ByVal Tx1 As Single = 0, _
                Optional ByVal Ty1 As Single = 0, _
                Optional ByVal Tx2 As Single = 1, _
                Optional ByVal Ty2 As Single = 1)

    Dim ColorReflection(3) As Long
    Dim i As Integer

    For i = 0 To 3
        If Color(i) = 0 Then Color(i) = map_base_light
        
        Dim ColorReflejo As Single, tmpColor As D3DCOLORVALUE
        
        Call ARGBtoD3DCOLORVALUE(Color(0), tmpColor)

        ColorReflection(i) = D3DColorARGB(tmpColor.a * 0.5, tmpColor.r * 0.2, tmpColor.g * 0.5, tmpColor.b)
    Next
    
    Dim dt As Single, dx As Single, dy As Single
    dt = ((GetTickCount And &H7FFFFFFF) Mod 10000) * 0.0006283
    dx = Sin(dt) * 8
    dy = Sin(dt * 2 + 1) * 5

    m_Data(m_Position).x = x + dx
    m_Data(m_Position).y = y + Height * 1.75 - dy
    m_Data(m_Position).z = 0.5
    m_Data(m_Position).Color = ColorReflection(0)
    m_Data(m_Position).TX = Tx1
    m_Data(m_Position).TY = Ty1
 
    m_Position = m_Position + 1
    
    m_Data(m_Position).x = x
    m_Data(m_Position).y = y + Height - 5
    m_Data(m_Position).z = 0.5
    m_Data(m_Position).Color = ColorReflection(1)
    m_Data(m_Position).TX = Tx1
    m_Data(m_Position).TY = Ty2
   
    m_Position = m_Position + 1
   
    m_Data(m_Position).x = x + Width + dx
    m_Data(m_Position).y = y + Height * 1.75 - dy
    m_Data(m_Position).z = 0.5
    m_Data(m_Position).Color = ColorReflection(2)
    m_Data(m_Position).TX = Tx2
    m_Data(m_Position).TY = Ty1

    m_Position = m_Position + 1
   
    m_Data(m_Position).x = x + Width
    m_Data(m_Position).y = y + Height - 5
    m_Data(m_Position).z = 0.5
    m_Data(m_Position).Color = ColorReflection(3)
    m_Data(m_Position).TX = Tx2
    m_Data(m_Position).TY = Ty2

    m_Position = m_Position + 1
   
    If m_Position >= UBound(m_Data) Then
        Call Flush
    End If
   
End Sub

Friend Sub DrawVertices(vertices() As TYPE_VERTEX)

    Dim i As Long

    For i = LBound(vertices) To UBound(vertices)
        m_Data(m_Position) = vertices(i)
        m_Position = m_Position + 1

        If m_Position >= UBound(m_Data) Then
            Call Flush
        End If
    Next

End Sub
 
Public Sub Finish()

    '
    ' Flush the batch since we're finished
    '
    Call Flush
 
End Sub
 
Public Sub Flush()
 
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
 
End Sub
 
Public Sub Initialize(ByVal Capacity As Long)

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

End Sub
