VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTexManager"
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
'    Argentum 20 - Game Client Program
'    Copyright (C) ...
'
'    clsTexManager — ad-hoc hash table version (no Scripting.Dictionary)

Option Explicit

'======================
' Hash table parameters
'======================
Private Const HASH_TABLE_SIZE As Long = 1024   ' Power of two preferred (fast AND masking)

'======================
' Entry & bucket layout
'======================
Private Type SURFACE_ENTRY
    filename        As Long
    UltimoAcceso    As Long
    texture_width   As Long
    texture_height  As Long
    size            As Long                  ' bytes
    Texture         As Direct3DTexture8
End Type

Private Type bucket
    surfaceCount    As Long
    SurfaceEntry()  As SURFACE_ENTRY         ' 1..surfaceCount
End Type

'======================
' State
'======================
Private TexList()               As bucket            ' 0..HASH_TABLE_SIZE-1
Private mD3D                    As D3DX8
Private device                  As Direct3DDevice8
Private mTotalTextureCount      As Long              ' number of cached textures
Private mTotalAllocatedBytes    As Long              ' sum of .size
Private mMaxEntries             As Long

Private Const E_OUTOFMEMORY = 7

'======================
' Public stats
'======================
Public Function GetAllocatedBytes() As Long
    GetAllocatedBytes = mTotalAllocatedBytes
End Function

Public Function GetAllocatedTextureCount() As Long
    GetAllocatedTextureCount = mTotalTextureCount
End Function

'======================
' Lifecycle
'======================
Public Function Init(ByRef D3D8 As D3DX8, ByRef d3d_device As Direct3DDevice8) As Boolean
On Error GoTo Init_Err
    Set mD3D = D3D8
    Set device = d3d_device
    Init = True
    Exit Function
Init_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsTexManager.Init", Erl)
    Resume Next
End Function

Private Sub Class_Initialize()
    Dim i As Long
    ReDim TexList(0 To HASH_TABLE_SIZE - 1) As bucket
    For i = 0 To HASH_TABLE_SIZE - 1
        TexList(i).surfaceCount = 0
        ' SurfaceEntry() left unallocated until first insert
    Next i
    mTotalTextureCount = 0
    mTotalAllocatedBytes = 0
End Sub

Private Sub Class_Terminate()
    Dim i As Long, j As Long
    On Error Resume Next
    For i = 0 To HASH_TABLE_SIZE - 1
        With TexList(i)
            For j = 1 To .surfaceCount
                Set .SurfaceEntry(j).Texture = Nothing
            Next j
            Erase .SurfaceEntry
            .surfaceCount = 0
        End With
    Next i
    On Error GoTo 0
End Sub

'======================
' Helpers
'======================
Private Function BucketIndex(ByVal key As Long) As Long
    ' Use positive mask then modulo (robust for negatives)
    BucketIndex = (key And &H7FFFFFFF) And (HASH_TABLE_SIZE - 1)
End Function

Private Sub EvictEntry(ByVal bucket As Long, ByVal idx As Long)
    Dim sz As Long, id As Long
    Dim k As Long
    With TexList(bucket).SurfaceEntry(idx)
        id = .filename
        sz = .size
        On Error Resume Next
        Set .Texture = Nothing
        On Error GoTo 0
        .filename = 0
        .texture_width = 0
        .texture_height = 0
        .size = 0
        .UltimoAcceso = 0
    End With

    ' Shift down within the bucket
    With TexList(bucket)
        For k = idx To .surfaceCount - 1
            .SurfaceEntry(k) = .SurfaceEntry(k + 1)
        Next k
        .surfaceCount = .surfaceCount - 1
        If .surfaceCount > 0 Then
            ReDim Preserve .SurfaceEntry(1 To .surfaceCount) As SURFACE_ENTRY
        Else
            Erase .SurfaceEntry
        End If
    End With

    ' Update totals & log
    mTotalTextureCount = mTotalTextureCount - 1
    mTotalAllocatedBytes = mTotalAllocatedBytes - sz
    If mTotalTextureCount < 0 Then mTotalTextureCount = 0
    If mTotalAllocatedBytes < 0 Then mTotalAllocatedBytes = 0

    Debug.Print "--- Evicted tex=" & id & _
                " (" & format$(sz / 1048576#, "0.00") & "MB);" & _
                " Count=" & mTotalTextureCount & _
                " Mem=" & format$(mTotalAllocatedBytes / 1048576#, "0.00") & "MB"
End Sub

'======================
' Public: get-or-load
'======================
Public Function GetTexture( _
    ByVal filename As Integer, _
    ByRef textwidth As Long, _
    ByRef textheight As Long _
) As Direct3DTexture8
On Error GoTo errhandler

    Dim key As Long: key = filename
    Dim b As Long: b = BucketIndex(key)
    Dim j As Long

    ' 1) Cache hit in bucket
    With TexList(b)
        For j = 1 To .surfaceCount
            If .SurfaceEntry(j).filename = key Then
                .SurfaceEntry(j).UltimoAcceso = GetTickCount()
                textwidth = .SurfaceEntry(j).texture_width
                textheight = .SurfaceEntry(j).texture_height
                Set GetTexture = .SurfaceEntry(j).Texture
                Exit Function
            End If
        Next j
    End With

    ' 2) Cache miss ? load texture
    Dim tex As Direct3DTexture8
    Set tex = CreateDirect3DTexture(key)
    If tex Is Nothing Then
        frmDebug.add_text_tracebox "Failed to load texture " & filename
        Exit Function
    End If

    ' Query size
    Dim desc As D3DSURFACE_DESC
    Call tex.GetLevelDesc(0, desc)
    textwidth = desc.Width
    textheight = desc.Height

    ' 3) Insert into bucket (append)
    With TexList(b)
        .surfaceCount = .surfaceCount + 1
        If .surfaceCount = 1 Then
            ReDim .SurfaceEntry(1 To 1) As SURFACE_ENTRY
        Else
            ReDim Preserve .SurfaceEntry(1 To .surfaceCount) As SURFACE_ENTRY
        End If

        With .SurfaceEntry(.surfaceCount)
            .filename = key
            .UltimoAcceso = GetTickCount()
            Set .Texture = tex
            .texture_width = desc.Width
            .texture_height = desc.Height
            .size = desc.size
        End With
    End With

    ' 4) Update counters & log
    mTotalTextureCount = mTotalTextureCount + 1
    mTotalAllocatedBytes = mTotalAllocatedBytes + desc.size
    Debug.Print "+++ Loaded tex=" & key & _
                " (" & format$(desc.size / 1048576#, "0.00") & "MB);" & _
                " Count=" & mTotalTextureCount & _
                " Mem=" & format$(mTotalAllocatedBytes / 1048576#, "0.00") & "MB"

    Set GetTexture = tex
    Exit Function

errhandler:
    Call RegistrarError(Err.Number, Err.Description, "clsTexManager.GetTexture", Erl)
    Resume Next
End Function

'======================
' Eviction: first >60s
'======================
Private Function RemoveLRU() As Boolean
On Error GoTo errhandler

    Const AGE_THRESHOLD_MS As Long = 60000
    Dim thresholdTick As Long: thresholdTick = GetTickCount() - AGE_THRESHOLD_MS

    Dim i As Long, j As Long
    For i = 0 To HASH_TABLE_SIZE - 1
        With TexList(i)
            For j = 1 To .surfaceCount
                If .SurfaceEntry(j).size > 0 And .SurfaceEntry(j).UltimoAcceso <= thresholdTick Then
                    Call EvictEntry(i, j)
                    RemoveLRU = True
                    Exit Function
                End If
            Next j
        End With
    Next i

    RemoveLRU = False
    Exit Function

errhandler:
    Call RegistrarError(Err.Number, Err.Description, "clsTexManager.RemoveLRU", Erl)
    Resume Next
End Function

Public Sub ReleaseUnusedTextures(ByVal num_textures As Integer)
    Dim i As Long
    For i = 1 To num_textures
        If Not RemoveLRU() Then Exit For
    Next
End Sub

'======================
' Texture creation
'======================
Private Function CreateDirect3DTexture(ByVal FileNum As Integer) As Direct3DTexture8
On Error GoTo ErrHandler

    Dim Texture As Direct3DTexture8
    Dim loadResult As Long
    
    Set CreateDirect3dTexture = Nothing
    loadResult = LoadTexture(str$(FileNum), Texture)
    
    Select Case loadResult
        Case D3DERR_INVALIDCALL
            frmDebug.add_text_tracebox "LoadTexture failed with D3DERR_INVALIDCALL"
        Case D3DERR_NOTAVAILABLE
            frmDebug.add_text_tracebox "LoadTexture failed with D3DERR_NOTAVAILABLE"
        Case D3DXERR_INVALIDDATA
            frmDebug.add_text_tracebox "LoadTexture failed with D3DXERR_INVALIDDATA"
        Case D3DERR_OUTOFVIDEOMEMORY, E_OUTOFMEMORY
            frmDebug.add_text_tracebox "LoadTexture failed with D3DERR_OUTOFVIDEOMEMORY"
            Call ReleaseUnusedTextures(250)
            loadResult = LoadTexture(str$(FileNum), Texture)
    End Select
    
    If loadResult <> D3D_OK Then
       frmDebug.add_text_tracebox "Failed to load texture file " & FileNum
       Exit Function
    End If
        
    Set CreateDirect3DTexture = Texture
    Exit Function

ErrHandler:
    frmDebug.add_text_tracebox "ERROR EN GRHLOAD>" & FileNum & ".png"
End Function

'======================
' Load from disk / resource
'======================
Private Function LoadTexture(ByVal filename As String, ByRef Dest As Direct3DTexture8) As Long
On Error GoTo LoadTexture_ErrHandler
    Set Dest = Nothing
    Dim bytArr() As Byte

#If Compresion = 1 Then
    If Not Extract_File_To_Memory(Graphics, App.path & "\..\Recursos\OUTPUT\", LTrim$(filename) & ".png", bytArr, ResourcesPassword) Then
        frmDebug.add_text_tracebox "¡No se puede cargar el grafico numero " & filename & "!"
        Exit Function
    End If

    Set Dest = mD3D.CreateTextureFromFileInMemoryEx( _
        device, bytArr(0), UBound(bytArr) + 1, _
        D3DX_DEFAULT, D3DX_DEFAULT, 1, 0, _
        D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT, _
        D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, &HFF000000, _
        ByVal 0, ByVal 0 _
    )
#Else
    Dim PathToFile As String
    PathToFile = App.path & "\..\Recursos\Graficos\" & LTrim$(filename) & ".png"

    If Not FileExist(PathToFile, vbArchive) Then
        frmDebug.add_text_tracebox "¡No se puede cargar el grafico numero " & PathToFile & "!"
        Exit Function
    End If

    Set Dest = mD3D.CreateTextureFromFileEx( _
        device, PathToFile, _
        0, 0, 1, 0, _
        D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT, _
        D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, &HFF000000, _
        ByVal 0, ByVal 0 _
    )
#End If

    Debug.Assert (Not (Dest Is Nothing))
    Exit Function

LoadTexture_ErrHandler:
    LoadTexture = Err.Number
End Function

'======================
' Procedural texture
'======================
Public Function CreateTexture(ByVal Width As Long, ByVal Height As Long) As Direct3DTexture8
On Error GoTo ErrHandler
    Dim Texture As Direct3DTexture8
    Set Texture = mD3D.CreateTexture(device, Width, Height, 1, 0, D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT)
    Set CreateTexture = Texture
    Exit Function
ErrHandler:
    frmDebug.add_text_tracebox "Failed to generate texture, " & Err.Description
End Function


