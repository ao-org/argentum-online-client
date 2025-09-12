VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUIRenderer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' clsUIRenderer.cls
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Requires UI_VERTEX UDT and CopyMemory declaration in a standard module:
' Public Type UI_VERTEX
'     x As Single: y As Single: z As Single: rhw As Single: color As Long
' End Type
' Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
'     (ByVal Destination As Long, ByRef Source As Any, ByVal Length As Long)

Private m_Device As Direct3DDevice8
Private m_VB As Direct3DVertexBuffer8
Private vertices() As UI_VERTEX
Private vertexCount As Long
Private vertexCapacity As Long

Private oldShader As Long
Private oldAlpha As Long
Private oldSrc As Long
Private oldDest As Long

Private Type UI_VERTEX
    x As Single         'x in screen space
    y As Single         'y in screen space
    z  As Single        'Normalized z
    rhw As Single       'Normalized z rhw
    Color As Long       'Vertex color
    u  As Single
    v  As Single
End Type

Private Const D3DFVF_UI_VERTEX = (D3DFVF_XYZRHW Or D3DFVF_DIFFUSE Or D3DFVF_TEX1)


' Initialize with device and max quads capacity
Public Sub Init(ByVal dev As Direct3DDevice8, ByVal maxQuads As Long)
    Set m_Device = dev
    vertexCapacity = maxQuads * 4
    ReDim vertices(0 To vertexCapacity - 1)
    vertexCount = 0
    ' Create dynamic vertex buffer
    Dim vbSize As Long
    vbSize = vertexCapacity * LenB(vertices(0))
    Set m_VB = dev.CreateVertexBuffer(vbSize, D3DUSAGE_DYNAMIC Or D3DUSAGE_WRITEONLY, D3DFVF_XYZRHW Or D3DFVF_DIFFUSE, D3DPOOL_DEFAULT)
    
End Sub

Public Function ColorToBytes(ByVal Red As Byte, _
                             ByVal Green As Byte, _
                             ByVal Blue As Byte, _
                             Optional ByVal Alpha As Byte = 255) As Byte()
    Dim Bytes(0 To 3) As Byte
    
    ' Store as RGBA
    Bytes(0) = Red
    Bytes(1) = Green
    Bytes(2) = Blue
    Bytes(3) = Alpha
    
    ColorToBytes = Bytes
End Function



Public Sub RenderText(Texto As String, ByVal x As Integer, ByVal y As Integer, ByRef text_color() As Byte, Optional ByVal font_index As Integer = 1, _
                            Optional multi_line As Boolean = False, Optional CharIndex As Integer = 0, Optional ByVal Alpha As Byte = 255)
                            
    
'    On Error GoTo Engine_Text_Render_Err
'
'
'
'
'    Dim a, b, c, d, e, f As Integer
'
'    Dim graf          As Grh
'
'    Dim temp_array(3) As RGBA
'
'    If CharIndex = 0 Then
'        a = 255
'    Else
'        a = Clamp(charlist(CharIndex).AlphaText, 0, 255)
'    End If
'
'    If Alpha <> 255 Then
'        a = Alpha
'    End If
'
'    Call RGBAList(temp_array, text_color(0).r, text_color(0).G, text_color(0).b, a)
'
'    Dim i              As Long
'
'    Dim removedDialogs As Long
'
'    For i = 0 To dialogCount - 1
'
'        'Decrease index to prevent jumping over a dialog
'        'Crappy VB will cache the limit of the For loop, so even if it changed, it won't matter
'        With dialogs(i - removedDialogs)
'
'            If FrameTime - .startTime >= .lifeTime Then
'                Call Char_Dialog_Remove(.CharIndex, CharIndex)
'
'                If a <= 0 Then
'                    removedDialogs = removedDialogs + 1
'
'                End If
'
'            Else
'
'            End If
'
'        End With
'
'    Next i
'
'    Dim Sombra(3) As RGBA 'Sombra
'    Call RGBAList(Sombra, text_color(0).r / 6, text_color(0).G / 6, text_color(0).b / 6, 0.8 * a)
'
'    If (Len(Texto) = 0) Then Exit Sub
'
'    d = 0
'
'    If multi_line = False Then
'        e = 0
'        f = 0
'
'        For a = 1 To Len(Texto)
'            b = Asc(mid(Texto, a, 1))
'            graf.GrhIndex = Fuentes(font_index).Caracteres(b)
'
'            If b = 32 Or b = 13 Then
'                If e >= 32 Then 'reemplazar por lo que os plazca
'                    f = f + 1
'                    e = 0
'                    d = 0
'                Else
'
'                    If b = 32 Then d = d + 4
'
'                End If
'
'            Else
'
'                If graf.GrhIndex > 12 Then
'
'                    'mega sombra O-matica
'                    graf.GrhIndex = Fuentes(font_index).Caracteres(b)
'
'                    If font_index <> 3 Then
'                        Call Draw_GrhFont(graf.GrhIndex, (x + d) + 1, y + 1 + f * 14, Sombra())
'
'                    End If
'
'                    Call Draw_GrhFont(graf.GrhIndex, (x + d), y + f * 14, temp_array())
'
'                    ' graf.grhindex = Fuentes(font_index).Caracteres(b)
'                    ' Grh_Render graf, (X + d), Y + f * 14, temp_array, False, False, False '14 es el height de esta fuente dsp lo hacemos dinamico
'                    d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth
'
'                End If
'
'            End If
'
'            e = e + 1
'        Next a
'
'    Else
'        e = 0
'        f = 0
'
'        For a = 1 To Len(Texto)
'            b = Asc(mid(Texto, a, 1))
'            graf.GrhIndex = Fuentes(font_index).Caracteres(b)
'
'            If b = 32 Or b = 13 Then
'                If e >= 20 Then 'reemplazar por lo que os plazca
'                    f = f + 1
'                    e = 0
'                    d = 0
'                Else
'
'                    If b = 32 Then d = d + 4
'
'                End If
'
'            Else
'
'                If graf.GrhIndex > 12 Then
'
'                    'mega sombra O-matica
'                    graf.GrhIndex = Fuentes(font_index).Caracteres(b)
'                    Call Draw_GrhFont(graf.GrhIndex, (x + d) + 1, y + 1 + f * 14, Sombra())
'                    Call Draw_GrhFont(graf.GrhIndex, (x + d), y + f * 14, temp_array())
'
'                    ' graf.grhindex = Fuentes(font_index).Caracteres(b)
'                    'Grh_Render graf, (x + d), y + f * 14, temp_array, False, False, False '14 es el height de esta fuente dsp lo hacemos dinamico
'                    If font_index = 4 Then
'                        d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth - 1
'                    Else
'                        d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth
'
'                    End If
'
'                End If
'
'            End If
'
'            e = e + 1
'        Next a
'
'    End If

    
    Exit Sub

Engine_Text_Render_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Textos.Engine_Text_Render", Erl)
    Resume Next
    
End Sub

' Add a quad defined by position, size, per-vertex colors array, and optional alpha
Public Sub AddQuad(ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, ByRef colors() As Long, Optional ByVal Alpha As Byte = 255)
    If vertexCount + 4 > vertexCapacity Then Exit Sub
    Dim idx As Long: idx = vertexCount
    Dim clr As Long, clrAlpha As Long
    Dim i As Long
    ' Build per-vertex color with correct signed alpha shift
    For i = 0 To 3
        If Alpha < 128 Then
            clrAlpha = Alpha * &H1000000
        Else
            clrAlpha = (Alpha - 256) * &H1000000
        End If
        ' Combine ARGB with RGB from colors(i)
        clr = clrAlpha Or (colors(i) And &HFFFFFF)
        ' Assign to vertex
        vertices(idx + i).x = IIf(i Mod 2 = 1, x + w, x)
        vertices(idx + i).y = IIf(i > 1, y + h, y)
        vertices(idx + i).z = 0
        vertices(idx + i).rhw = 1
        vertices(idx + i).Color = clr
    Next i
    vertexCount = vertexCount + 4
End Sub

Public Sub AddTexturedQuad(ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, ByRef colors() As Long, Optional ByVal Alpha As Byte = 255)
    'TODO
End Sub


Private Sub SaveRenderState(ByVal dev As Direct3DDevice8)
    oldShader = dev.GetVertexShader
    oldAlpha = dev.GetRenderState(D3DRS_ALPHABLENDENABLE)
    oldSrc = dev.GetRenderState(D3DRS_SRCBLEND)
    oldDest = dev.GetRenderState(D3DRS_DESTBLEND)
End Sub

Private Sub RestoreRenderState(ByVal dev As Direct3DDevice8)
    ' Restore vertex shader (or FVF)
    dev.SetVertexShader oldShader
    ' Restore alpha blend states
    dev.SetRenderState D3DRS_ALPHABLENDENABLE, oldAlpha
    dev.SetRenderState D3DRS_SRCBLEND, oldSrc
    dev.SetRenderState D3DRS_DESTBLEND, oldDest
End Sub

' Render all queued quads and clear buffer
Public Sub Draw(ByVal dev As Direct3DDevice8)
    If vertexCount = 0 Then Exit Sub
    ' Save states
    SaveRenderState dev

    ' Prepare VB for write
    Dim pData As Long
    m_VB.Lock 0, vertexCount * LenB(vertices(0)), pData, D3DLOCK_DISCARD
    CopyMemory ByVal pData, vertices(0), vertexCount * LenB(vertices(0))
    m_VB.Unlock
    

    ' Set FVF and stream source
    dev.SetStreamSource 0, m_VB, LenB(vertices(0))
    dev.SetVertexShader D3DFVF_UI_VERTEX
    dev.SetTexture 0, Nothing
    dev.SetRenderState D3DRS_ALPHABLENDENABLE, True
    dev.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    dev.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA

    ' Draw quads as triangle strips
    Dim i As Long
    For i = 0 To vertexCount - 1 Step 4
        dev.DrawPrimitive D3DPT_TRIANGLESTRIP, i, 2
    Next

    ' Restore states
    RestoreRenderState dev

    ' Clear for next frame
    vertexCount = 0
End Sub

' IUIRenderer interface
Public Sub AddComponentQuad(ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, ByRef colors() As Long)
    Me.AddQuad x, y, w, h, colors
End Sub
Public Sub RenderAll(ByVal dev As Direct3DDevice8)
    Draw dev
End Sub

