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
    VB0Stride As Long
    BaseVBIndex As Long
    VB0 As Direct3DVertexBuffer8
    ib  As Direct3DIndexBuffer8
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

Public Function ColorToBytes(ByVal R As Byte, ByVal G As Byte, ByVal B As Byte, Optional ByVal A As Byte = 255) As Long
    Dim p(0 To 3) As Byte   ' [0]=B, [1]=G, [2]=R, [3]=A
    p(0) = B
    p(1) = G
    p(2) = R
    p(3) = A
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
Private Function ARGB(ByVal A As Byte, ByVal R As Byte, ByVal G As Byte, ByVal B As Byte) As Long
    Dim p(0 To 3) As Byte
    p(0) = B: p(1) = G: p(2) = R: p(3) = A
    CopyMemory ARGB, p(0), 4
End Function

' Convert a VB/OLE color (BGR: 0x00BBGGRR) + alpha -> D3D ARGB.
Private Function FromVBColor(ByVal vbColor As Long, ByVal A As Byte) As Long
    FromVBColor = ARGB(A, vbColor And &HFF&, (vbColor \ &H100&) And &HFF&, (vbColor \ &H10000) And &HFF&)
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
Private Function ColorWithAlpha(ByVal baseColor As Long, ByVal A As Byte) As Long
    ' baseColor is a VB/OLE color (0x00BBGGRR). Pack to D3D ARGB (0xAARRGGBB)
    Dim out(0 To 3) As Byte
    out(0) = (baseColor \ &H10000) And &HFF&  ' B
    out(1) = (baseColor \ &H100&) And &HFF&   ' G
    out(2) = baseColor And &HFF&              ' R
    out(3) = A                                ' A
    CopyMemory ColorWithAlpha, out(0), 4
End Function

Private Function SafeColor(ByRef arr() As Long, ByVal idx As Long, ByVal Fallback As Long) As Long
    On Error GoTo Fallback
    SafeColor = arr(LBound(arr) + idx)
    Exit Function
Fallback:
    SafeColor = Fallback
End Function

Public Sub RenderText(Texto As String, _
                      ByVal x As Integer, _
                      ByVal y As Integer, _
                      ByVal color As Long, _
                      Optional ByVal font_index As Integer = 1, _
                      Optional multi_line As Boolean = False, _
                      Optional charindex As Integer = 0, _
                      Optional ByVal alpha As Byte = 255)
    '    On Error GoTo Engine_Text_Render_Err
End Sub
