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
End Type

Private Const D3DFVF_UI_VERTEX = (D3DFVF_XYZRHW Or D3DFVF_DIFFUSE)


' Initialize with device and max quads capacity
Public Sub Init(ByVal dev As Direct3DDevice8, ByVal maxQuads As Long)
    On Error Goto Init_Err
    Set m_Device = dev
    vertexCapacity = maxQuads * 4
    ReDim vertices(0 To vertexCapacity - 1)
    vertexCount = 0
    ' Create dynamic vertex buffer
    Dim vbSize As Long
    vbSize = vertexCapacity * LenB(vertices(0))
    Set m_VB = dev.CreateVertexBuffer(vbSize, D3DUSAGE_DYNAMIC Or D3DUSAGE_WRITEONLY, D3DFVF_XYZRHW Or D3DFVF_DIFFUSE, D3DPOOL_DEFAULT)
    
    Exit Sub
Init_Err:
    Call TraceError(Err.Number, Err.Description, "clsUIRenderer.Init", Erl)
End Sub

' Add a quad defined by position, size, per-vertex colors array, and optional alpha
Public Sub AddQuad(ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, ByRef colors() As Long, Optional ByVal Alpha As Byte = 255)
    On Error Goto AddQuad_Err
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
    Exit Sub
AddQuad_Err:
    Call TraceError(Err.Number, Err.Description, "clsUIRenderer.AddQuad", Erl)
End Sub


Private Sub SaveRenderState(ByVal dev As Direct3DDevice8)
    On Error Goto SaveRenderState_Err
    oldShader = dev.GetVertexShader
    oldAlpha = dev.GetRenderState(D3DRS_ALPHABLENDENABLE)
    oldSrc = dev.GetRenderState(D3DRS_SRCBLEND)
    oldDest = dev.GetRenderState(D3DRS_DESTBLEND)
    Exit Sub
SaveRenderState_Err:
    Call TraceError(Err.Number, Err.Description, "clsUIRenderer.SaveRenderState", Erl)
End Sub

Private Sub RestoreRenderState(ByVal dev As Direct3DDevice8)
    On Error Goto RestoreRenderState_Err
    ' Restore vertex shader (or FVF)
    dev.SetVertexShader oldShader
    ' Restore alpha blend states
    dev.SetRenderState D3DRS_ALPHABLENDENABLE, oldAlpha
    dev.SetRenderState D3DRS_SRCBLEND, oldSrc
    dev.SetRenderState D3DRS_DESTBLEND, oldDest
    Exit Sub
RestoreRenderState_Err:
    Call TraceError(Err.Number, Err.Description, "clsUIRenderer.RestoreRenderState", Erl)
End Sub

' Render all queued quads and clear buffer
Public Sub Draw(ByVal dev As Direct3DDevice8)
    On Error Goto Draw_Err
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
    Exit Sub
Draw_Err:
    Call TraceError(Err.Number, Err.Description, "clsUIRenderer.Draw", Erl)
End Sub

' IUIRenderer interface
Public Sub AddComponentQuad(ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, ByRef colors() As Long)
    On Error Goto AddComponentQuad_Err
    Me.AddQuad x, y, w, h, colors
    Exit Sub
AddComponentQuad_Err:
    Call TraceError(Err.Number, Err.Description, "clsUIRenderer.AddComponentQuad", Erl)
End Sub
Public Sub RenderAll(ByVal dev As Direct3DDevice8)
    On Error Goto RenderAll_Err
    Draw dev
    Exit Sub
RenderAll_Err:
    Call TraceError(Err.Number, Err.Description, "clsUIRenderer.RenderAll", Erl)
End Sub

