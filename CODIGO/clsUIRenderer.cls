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
Option Explicit
' =====================================================================================
' clsUIRenderer
' - Unified vertex (XYZRHW + DIFFUSE + TEX1)
' - Buffers allocated once in Init()
' - Public API preserved: Init, AddQuad, Draw, Save/RestoreRenderState
' - New optional helpers: AddTexturedQuad, AddTexturedQuadPx
' =====================================================================================
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

' ---------------------------------------------------------------------
' Unified vertex type
Private Type UI_VERTEX
    x   As Single
    y   As Single
    z   As Single
    rhw As Single
    color As Long
    tu  As Single
    tv  As Single
End Type

Private Const D3DFVF_UI_VERTEX As Long = (D3DFVF_XYZRHW Or D3DFVF_DIFFUSE Or D3DFVF_TEX1)

Private Type UIDrawCmd
    startV As Long
    tex    As Direct3DTexture8
End Type

' ---------------------------------------------------------------------
' Buffers
Private vertices()  As UI_VERTEX
Private cmds()      As UIDrawCmd
Private vertexCount As Long
Private cmdCount    As Long
Private m_VB        As Direct3DVertexBuffer8
Private maxVerts    As Long
Private maxCmds     As Long
Private mHalfPixel  As Single

' ---------------------------------------------------------------------
' Saved render state
Private Type SavedState
    Shader      As Long
    VB0Stride   As Long
    BaseVBIndex As Long
    VB0         As Direct3DVertexBuffer8
    ib          As Direct3DIndexBuffer8
    Tex0        As Direct3DBaseTexture8
    ZEnable     As Long
    AlphaBlend  As Long
    SrcBlend    As Long
    DstBlend    As Long
    ColorOp     As Long
    ColorArg1   As Long
    ColorArg2   As Long
    AlphaOp     As Long
    AlphaArg1   As Long
    AlphaArg2   As Long
    MinFilter   As Long
    MagFilter   As Long
    MipFilter   As Long
End Type

Private mSaved      As SavedState
Private gFrameCount As Long
Private gLastTick   As Long
Private gFPS        As Long


Private mTexMan As clsTexManager

Private Sub InitFPS()
    gFrameCount = 0
    gLastTick = GetTickCount()
    gFPS = 0
End Sub

Private Sub UpdateAndRenderFPS()
    ' --- Update ---
    gFrameCount = gFrameCount + 1
    Dim nowT As Long: nowT = GetTickCount()
    If nowT - gLastTick >= 1000 Then
        gFPS = gFrameCount
        gFrameCount = 0
        gLastTick = nowT
    End If
    ' --- Draw ---
    Dim col(3) As RGBA
    col(0) = RGBA_From_Comp(255, 255, 255) ' white
    col(1) = col(0)
    col(2) = col(0)
    col(3) = col(0)
    simple_text_render "FPS: " & CStr(gFPS), 10, 10, col, 1, True
End Sub

Public Function ColorToBytes(ByVal r As Byte, ByVal G As Byte, ByVal b As Byte, Optional ByVal a As Byte = 255) As Long
    Dim p(0 To 3) As Byte   ' [0]=B, [1]=G, [2]=R, [3]=A
    p(0) = b
    p(1) = G
    p(2) = r
    p(3) = a
    CopyMemory ColorToBytes, p(0), 4
End Function

' =====================================================================================
' Init (allocate once)
' =====================================================================================
Public Sub Init(ByVal dev As Direct3DDevice8, ByVal maxQuads As Long)
    InitFPS
    If maxQuads < 1 Then maxQuads = 1
    maxVerts = maxQuads * 4
    maxCmds = maxQuads
    ReDim vertices(0 To maxVerts - 1)
    ReDim cmds(0 To maxCmds - 1)
    Set m_VB = Nothing
    Set m_VB = dev.CreateVertexBuffer(maxVerts * LenB(vertices(0)), D3DUSAGE_WRITEONLY Or D3DUSAGE_DYNAMIC, D3DFVF_UI_VERTEX, D3DPOOL_DEFAULT)
    vertexCount = 0
    cmdCount = 0
    mHalfPixel = 0.5!
End Sub

' =====================================================================================
' Public API
' =====================================================================================
' Pack bytes to ARGB (what D3D expects). Memory order = [B,G,R,A].
Private Function ARGB(ByVal a As Byte, ByVal r As Byte, ByVal G As Byte, ByVal b As Byte) As Long
    Dim p(0 To 3) As Byte
    p(0) = b: p(1) = G: p(2) = r: p(3) = a
    CopyMemory ARGB, p(0), 4
End Function

' Convert a VB/OLE color (BGR: 0x00BBGGRR) + alpha -> D3D ARGB.
Private Function FromVBColor(ByVal vbColor As Long, ByVal a As Byte) As Long
    FromVBColor = ARGB(a, vbColor And &HFF&, (vbColor \ &H100&) And &HFF&, (vbColor \ &H10000) And &HFF&)
End Function

Public Sub AddQuad(ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, ByRef colors() As Long, Optional ByVal alpha As Byte = 255)
    If cmdCount >= maxCmds Then Exit Sub
    Dim cTL  As Long, cTR As Long, cBL As Long, cBR As Long
    Dim base As Long, has4 As Boolean
    has4 = (UBound(colors) - LBound(colors) + 1) >= 4
    If has4 Then
        cTL = FromVBColor(colors(LBound(colors) + 0), alpha)
        cTR = FromVBColor(colors(LBound(colors) + 1), alpha)
        cBL = FromVBColor(colors(LBound(colors) + 2), alpha)
        cBR = FromVBColor(colors(LBound(colors) + 3), alpha)
    Else
        base = FromVBColor(colors(LBound(colors)), alpha)
        cTL = base: cTR = base: cBL = base: cBR = base
    End If
    Dim x0!, y0!, x1!, y1!
    x0 = CSng(x) - mHalfPixel
    y0 = CSng(y) - mHalfPixel
    x1 = CSng(x + w) - mHalfPixel
    y1 = CSng(y + h) - mHalfPixel
    Const Z_UI! = 0!, RHW_UI! = 1!
    With vertices(vertexCount + 0): .x = x0: .y = y0: .z = Z_UI: .rhw = RHW_UI: .color = cTL: .tu = 0!: .tv = 0!: End With
    With vertices(vertexCount + 1): .x = x1: .y = y0: .z = Z_UI: .rhw = RHW_UI: .color = cTR: .tu = 1!: .tv = 0!: End With
    With vertices(vertexCount + 2): .x = x0: .y = y1: .z = Z_UI: .rhw = RHW_UI: .color = cBL: .tu = 0!: .tv = 1!: End With
    With vertices(vertexCount + 3): .x = x1: .y = y1: .z = Z_UI: .rhw = RHW_UI: .color = cBR: .tu = 1!: .tv = 1!: End With
    cmds(cmdCount).startV = vertexCount
    Set cmds(cmdCount).tex = Nothing
    cmdCount = cmdCount + 1
    vertexCount = vertexCount + 4
End Sub

' --- Textured quad (UVs normalized) ---
Public Sub AddTexturedQuad(ByVal x As Long, _
                           ByVal y As Long, _
                           ByVal w As Long, _
                           ByVal h As Long, _
                           ByVal tex As Direct3DTexture8, _
                           ByVal u0 As Single, _
                           ByVal v0 As Single, _
                           ByVal u1 As Single, _
                           ByVal v1 As Single, _
                           Optional ByVal colorARGB As Long = &HFFFFFFFF, _
                           Optional ByVal alpha As Byte = 255)
    If tex Is Nothing Then Exit Sub
    If cmdCount >= maxCmds Then Exit Sub
    Dim c As Long: c = ColorWithAlpha(colorARGB, alpha)
    Dim x0!, y0!, x1!, y1!
    x0 = CSng(x) - mHalfPixel
    y0 = CSng(y) - mHalfPixel
    x1 = CSng(x + w) - mHalfPixel
    y1 = CSng(y + h) - mHalfPixel
    Const Z_UI! = 0!, RHW_UI! = 1!
    With vertices(vertexCount + 0)
        .x = x0: .y = y0: .z = Z_UI: .rhw = RHW_UI: .color = c: .tu = u0: .tv = v0
    End With
    With vertices(vertexCount + 1)
        .x = x1: .y = y0: .z = Z_UI: .rhw = RHW_UI: .color = c: .tu = u1: .tv = v0
    End With
    With vertices(vertexCount + 2)
        .x = x0: .y = y1: .z = Z_UI: .rhw = RHW_UI: .color = c: .tu = u0: .tv = v1
    End With
    With vertices(vertexCount + 3)
        .x = x1: .y = y1: .z = Z_UI: .rhw = RHW_UI: .color = c: .tu = u1: .tv = v1
    End With
    cmds(cmdCount).startV = vertexCount
    Set cmds(cmdCount).tex = tex
    cmdCount = cmdCount + 1
    vertexCount = vertexCount + 4
End Sub

' --- Textured quad with atlas pixel coords ---
Public Sub AddTexturedQuadPx(ByVal x As Long, _
                             ByVal y As Long, _
                             ByVal w As Long, _
                             ByVal h As Long, _
                             ByVal tex As Direct3DTexture8, _
                             ByVal texW As Long, _
                             ByVal texH As Long, _
                             ByVal srcX As Long, _
                             ByVal srcY As Long, _
                             ByVal srcW As Long, _
                             ByVal srcH As Long, _
                             Optional ByVal colorARGB As Long = &HFFFFFFFF, _
                             Optional ByVal alpha As Byte = 255)
    If tex Is Nothing Or texW <= 0 Or texH <= 0 Then Exit Sub
    Dim u0!, v0!, u1!, v1!
    u0 = srcX / texW: v0 = srcY / texH
    u1 = (srcX + srcW) / texW
    v1 = (srcY + srcH) / texH
    AddTexturedQuad x, y, w, h, tex, u0, v0, u1, v1, colorARGB, alpha
End Sub

Public Sub Draw(ByVal dev As Direct3DDevice8)
    If vertexCount = 0 Then Exit Sub
    ' Upload
    Dim pData As Long
    m_VB.Lock 0, vertexCount * LenB(vertices(0)), pData, D3DLOCK_DISCARD
    CopyMemory ByVal pData, vertices(0), vertexCount * LenB(vertices(0))
    m_VB.Unlock
    ' Save states (your existing helpers)
    SaveRenderState dev
    ' --- Force a clean 2D overlay state ---
    dev.SetStreamSource 0, m_VB, LenB(vertices(0))
    ' Force a clean overlay state (add this block)
    dev.SetVertexShader D3DFVF_UI_VERTEX
    dev.SetRenderState D3DRS_ZENABLE, D3DZB_FALSE
    dev.SetRenderState D3DRS_LIGHTING, 0
    dev.SetRenderState D3DRS_FOGENABLE, 0
    dev.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
    dev.SetRenderState D3DRS_ALPHABLENDENABLE, 1
    dev.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    dev.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    dev.SetRenderState D3DRS_COLORWRITEENABLE, &HF           ' RGBA
    dev.SetRenderState D3DRS_ALPHATESTENABLE, 0
    dev.SetRenderState D3DRS_STENCILENABLE, 0
    ' Neutralize any leftover tint factor
    dev.SetRenderState D3DRS_TEXTUREFACTOR, &HFFFFFFFF       ' white
    ' Hard-disable every stage =1 so they cannot tint UI
    Dim s As Long
    For s = 1 To 7
        dev.SetTexture s, Nothing
        dev.SetTextureStageState s, D3DTSS_COLOROP, D3DTOP_DISABLE
        dev.SetTextureStageState s, D3DTSS_ALPHAOP, D3DTOP_DISABLE
    Next
    ' Stage 0 set per-quad later, but you can also set defaults here:
    dev.SetTexture 0, Nothing
    dev.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_SELECTARG2
    dev.SetTextureStageState 0, D3DTSS_COLORARG2, D3DTA_DIFFUSE
    dev.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG2
    dev.SetTextureStageState 0, D3DTSS_ALPHAARG2, D3DTA_DIFFUSE
    ' --- Draw queued quads ---
    Dim i As Long
    For i = 0 To cmdCount - 1
        If cmds(i).tex Is Nothing Then
            ' Solid = DIFFUSE only
            dev.SetTexture 0, Nothing
            dev.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_SELECTARG2
            dev.SetTextureStageState 0, D3DTSS_COLORARG2, D3DTA_DIFFUSE
            dev.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG2
            dev.SetTextureStageState 0, D3DTSS_ALPHAARG2, D3DTA_DIFFUSE
        Else
            ' Textured = Texture * Diffuse
            dev.SetTexture 0, cmds(i).tex
            dev.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
            dev.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
            dev.SetTextureStageState 0, D3DTSS_COLORARG2, D3DTA_DIFFUSE
            dev.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
            dev.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
            dev.SetTextureStageState 0, D3DTSS_ALPHAARG2, D3DTA_DIFFUSE
        End If
        dev.DrawPrimitive D3DPT_TRIANGLESTRIP, cmds(i).startV, 2
    Next
    ' Restore game states
    RestoreRenderState dev
    ' Reset batch
    vertexCount = 0
    cmdCount = 0
    UpdateAndRenderFPS
End Sub

' =====================================================================================
' State helpers
' =====================================================================================
Public Sub SaveRenderState(ByVal dev As Direct3DDevice8)
    Dim vb     As Direct3DVertexBuffer8, ib As Direct3DIndexBuffer8
    Dim stride As Long, baseIdx As Long
    Set mSaved.Tex0 = dev.GetTexture(0)
    mSaved.Shader = dev.GetVertexShader
    dev.GetStreamSource 0, vb, stride: Set mSaved.VB0 = vb: mSaved.VB0Stride = stride
    dev.GetIndices ib, baseIdx: Set mSaved.ib = ib: mSaved.BaseVBIndex = baseIdx
    mSaved.ZEnable = dev.GetRenderState(D3DRS_ZENABLE)
    mSaved.AlphaBlend = dev.GetRenderState(D3DRS_ALPHABLENDENABLE)
    mSaved.SrcBlend = dev.GetRenderState(D3DRS_SRCBLEND)
    mSaved.DstBlend = dev.GetRenderState(D3DRS_DESTBLEND)
    mSaved.ColorOp = dev.GetTextureStageState(0, D3DTSS_COLOROP)
    mSaved.ColorArg1 = dev.GetTextureStageState(0, D3DTSS_COLORARG1)
    mSaved.ColorArg2 = dev.GetTextureStageState(0, D3DTSS_COLORARG2)
    mSaved.AlphaOp = dev.GetTextureStageState(0, D3DTSS_ALPHAOP)
    mSaved.AlphaArg1 = dev.GetTextureStageState(0, D3DTSS_ALPHAARG1)
    mSaved.AlphaArg2 = dev.GetTextureStageState(0, D3DTSS_ALPHAARG2)
    mSaved.MinFilter = dev.GetTextureStageState(0, D3DTSS_MINFILTER)
    mSaved.MagFilter = dev.GetTextureStageState(0, D3DTSS_MAGFILTER)
    mSaved.MipFilter = dev.GetTextureStageState(0, D3DTSS_MIPFILTER)
End Sub

Public Sub RestoreRenderState(ByVal dev As Direct3DDevice8)
    dev.SetTexture 0, mSaved.Tex0
    Set mSaved.Tex0 = Nothing
    dev.SetVertexShader mSaved.Shader
    dev.SetStreamSource 0, mSaved.VB0, mSaved.VB0Stride
    dev.SetIndices mSaved.ib, mSaved.BaseVBIndex
    dev.SetRenderState D3DRS_ZENABLE, mSaved.ZEnable
    dev.SetRenderState D3DRS_ALPHABLENDENABLE, mSaved.AlphaBlend
    dev.SetRenderState D3DRS_SRCBLEND, mSaved.SrcBlend
    dev.SetRenderState D3DRS_DESTBLEND, mSaved.DstBlend
    dev.SetTextureStageState 0, D3DTSS_COLOROP, mSaved.ColorOp
    dev.SetTextureStageState 0, D3DTSS_COLORARG1, mSaved.ColorArg1
    dev.SetTextureStageState 0, D3DTSS_COLORARG2, mSaved.ColorArg2
    dev.SetTextureStageState 0, D3DTSS_ALPHAOP, mSaved.AlphaOp
    dev.SetTextureStageState 0, D3DTSS_ALPHAARG1, mSaved.AlphaArg1
    dev.SetTextureStageState 0, D3DTSS_ALPHAARG2, mSaved.AlphaArg2
    dev.SetTextureStageState 0, D3DTSS_MINFILTER, mSaved.MinFilter
    dev.SetTextureStageState 0, D3DTSS_MAGFILTER, mSaved.MagFilter
    dev.SetTextureStageState 0, D3DTSS_MIPFILTER, mSaved.MipFilter
End Sub

' =====================================================================================
' Helpers
' =====================================================================================
Private Function ColorWithAlpha(ByVal baseColor As Long, ByVal a As Byte) As Long
    ' baseColor is a VB/OLE color (0x00BBGGRR). Pack to D3D ARGB (0xAARRGGBB)
    Dim out(0 To 3) As Byte
    out(0) = (baseColor \ &H10000) And &HFF&  ' B
    out(1) = (baseColor \ &H100&) And &HFF&   ' G
    out(2) = baseColor And &HFF&              ' R
    out(3) = a                                ' A
    CopyMemory ColorWithAlpha, out(0), 4
End Function

Private Function SafeColor(ByRef arr() As Long, ByVal idx As Long, ByVal Fallback As Long) As Long
    On Error GoTo Fallback
    SafeColor = arr(LBound(arr) + idx)
    Exit Function
Fallback:
    SafeColor = Fallback
End Function




Private Function ARGBToVBColor(ByVal argbColor As Long) As Long
    Dim r As Long, G As Long, b As Long
    r = (argbColor \ &H10000) And &HFF&
    G = (argbColor \ &H100&) And &HFF&
    b = argbColor And &HFF&
    ARGBToVBColor = (b * &H10000) Or (G * &H100) Or r
End Function

Private Function NormalizeTextColor(ByVal inputColor As Long) As Long
    If (inputColor And &HFF000000) = 0 Then
        NormalizeTextColor = inputColor
    Else
        NormalizeTextColor = ARGBToVBColor(inputColor)
    End If
End Function

Private Sub AddGlyphQuad(ByVal glyphIndex As Long, ByVal drawX As Long, ByVal drawY As Long, ByVal vbColor As Long, ByVal alphaVal As Byte)
    If glyphIndex <= 0 Then Exit Sub

    Dim finalIndex As Long
    finalIndex = glyphIndex

    If GrhData(finalIndex).NumFrames > 0 Then
        If GrhData(finalIndex).NumFrames >= 1 Then
            finalIndex = GrhData(finalIndex).Frames(1)
        End If
    End If

    If finalIndex <= 0 Then Exit Sub

    Dim glyphWidth As Long
    Dim glyphHeight As Long
    glyphWidth = GrhData(finalIndex).pixelWidth + 1
    glyphHeight = GrhData(finalIndex).pixelHeight + 1

    Dim cols(3) As Long
    cols(0) = vbColor
    cols(1) = vbColor
    cols(2) = vbColor
    cols(3) = vbColor

    Dim startCmd As Long
    Dim startVert As Long
    startCmd = cmdCount
    startVert = vertexCount

    AddQuad drawX, drawY, glyphWidth, glyphHeight, cols, alphaVal

    If cmdCount <= startCmd Then Exit Sub

    If SurfaceDB Is Nothing Then Exit Sub

    Dim texwidth As Long, texheight As Long
    Dim glyphTex As Direct3DTexture8
    Set glyphTex = SurfaceDB.GetTexture(GrhData(finalIndex).FileNum, texwidth, texheight)
    Set cmds(startCmd).tex = glyphTex

    Dim u0 As Single, v0 As Single, u1 As Single, v1 As Single
    If texwidth <> 0 And texheight <> 0 Then
        u0 = (GrhData(finalIndex).sX + 0.25) / texwidth
        v0 = (GrhData(finalIndex).sY + 0.25) / texheight
        u1 = (GrhData(finalIndex).sX + glyphWidth) / texwidth
        v1 = (GrhData(finalIndex).sY + glyphHeight) / texheight
    Else
        u0 = 0!
        v0 = 0!
        u1 = 1!
        v1 = 1!
    End If

    With vertices(startVert + 0)
        .tu = u0
        .tv = v0
    End With
    With vertices(startVert + 1)
        .tu = u1
        .tv = v0
    End With
    With vertices(startVert + 2)
        .tu = u0
        .tv = v1
    End With
    With vertices(startVert + 3)
        .tu = u1
        .tv = v1
    End With
End Sub

Public Sub RenderText(Texto As String, _
                      ByVal x As Integer, _
                      ByVal y As Integer, _
                      ByVal color As Long, _
                      Optional ByVal font_index As Integer = 1, _
                      Optional ByVal alpha As Byte = 255)
   
 
 On Error GoTo RenderText_Err

    If Len(Texto) = 0 Then Exit Sub

    Dim graf As Grh
    Dim AlphaValue As Byte
    AlphaValue = alpha

    Dim vbColor As Long
    vbColor = NormalizeTextColor(color)

    Dim advanceX As Long
    Dim charsOnLine As Long
    Dim lineIndex As Long
    Dim idx As Long
    Const WRAP_THRESHOLD As Long = 33

    advanceX = 0
    For idx = 1 To Len(Texto)
        Dim charCode As Long
        charCode = Asc(mid$(Texto, idx, 1))

        If charCode = 13 Or charCode = 10 Then
            lineIndex = lineIndex + 1
            advanceX = 0
            charsOnLine = 0
        Else
            graf.GrhIndex = Fuentes(font_index).Caracteres(charCode)

            If charCode = 32 Then
                If charsOnLine >= WRAP_THRESHOLD Then
                    lineIndex = lineIndex + 1
                    advanceX = 0
                    charsOnLine = 0
                Else
                    advanceX = advanceX + 2
                    charsOnLine = charsOnLine + 1
                End If
            ElseIf graf.GrhIndex > 12 Then
                Dim glyphIdx As Long
                glyphIdx = graf.GrhIndex

                AddGlyphQuad glyphIdx, x + advanceX, y + lineIndex * 14, vbColor, AlphaValue

                If GrhData(glyphIdx).NumFrames > 0 Then
                    glyphIdx = GrhData(glyphIdx).Frames(1)
                End If

                If glyphIdx > 0 Then
                    If font_index = 4 Then
                        advanceX = advanceX + GrhData(glyphIdx).pixelWidth - 1
                    Else
                        advanceX = advanceX + GrhData(glyphIdx).pixelWidth
                    End If
                End If

                charsOnLine = charsOnLine + 1
            End If
        End If
    Next idx

    Exit Sub
RenderText_Err:
    
    Call RegistrarError(Err.Number, Err.Description, "clsUIRenderer.RenderText", Erl)
    Resume Next
End Sub
