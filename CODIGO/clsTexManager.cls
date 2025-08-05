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

Option Explicit

'— Maps filename?Direct3DTexture8
Private mDictTextures        As Scripting.Dictionary
'— Maps filename?Meta array: (0)=UltimoAcceso, (1)=width, (2)=height, (3)=size
Private mDictMeta            As Scripting.Dictionary

Const E_OUTOFMEMORY = 7

Private mD3D                         As D3DX8
Private device                       As Direct3DDevice8
Private mMaxEntries                  As Integer

Private mTotalTextureCount   As Long    ' # of textures currently cached
Private mTotalAllocatedBytes As Long    ' sum of .size for all cached textures

Public Function GetAllocatedBytes() As Long
    GetAllocatedBytes = mTotalAllocatedBytes
End Function

Public Function GetAllocatedTextureCount() As Long
    GetAllocatedTextureCount = mTotalTextureCount
End Function

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
    Set mDictTextures = New Scripting.Dictionary
    Set mDictMeta = New Scripting.Dictionary
    mTotalTextureCount = 0
    mTotalAllocatedBytes = 0
End Sub

Private Sub Class_Terminate()
    Dim key As Variant
    ' Release D3D textures
    For Each key In mDictTextures.Keys
        On Error Resume Next
        Set mDictTextures(key) = Nothing
    Next key
    mDictTextures.RemoveAll
    mDictMeta.RemoveAll
    Set mDictTextures = Nothing
    Set mDictMeta = Nothing
End Sub


'———————————————————————————————————————————————————————————
' Public: get-or-load a texture (LRU eviction by age)
'———————————————————————————————————————————————————————————
Public Function GetTexture( _
    ByVal filename As Integer, _
    ByRef textwidth As Long, _
    ByRef textheight As Long _
) As Direct3DTexture8
    On Error GoTo errhandler

    Dim key      As String
    key = CStr(filename)

    ' — 1) Cache hit? update access time and return —
    If mDictTextures.Exists(key) Then
        Dim meta As Variant
        meta = mDictMeta.Item(key)
        meta(0) = GetTickCount()                ' refresh access time
        mDictMeta.Item(key) = meta              ' write back

        textwidth = meta(1)
        textheight = meta(2)
        Set GetTexture = mDictTextures.Item(key)
        Exit Function
    End If

    ' — 2) Cache miss ? create (may trigger eviction) —
    Dim tex As Direct3DTexture8
    Set tex = CreateDirect3DTexture(filename)
    If tex Is Nothing Then Exit Function

    ' Query its size
    Dim desc As D3DSURFACE_DESC
    Call tex.GetLevelDesc(0, desc)
    textwidth = desc.Width
    textheight = desc.Height

    ' — 3) Insert into dictionaries —
    Dim newMeta(0 To 3) As Variant
    newMeta(0) = GetTickCount()  ' UltimoAcceso
    newMeta(1) = desc.Width
    newMeta(2) = desc.Height
    newMeta(3) = desc.size

    mDictTextures.Add key, tex
    mDictMeta.Add key, newMeta

    ' — 4) Update counters & debug log —
    mTotalTextureCount = mDictTextures.count
    mTotalAllocatedBytes = mTotalAllocatedBytes + desc.size
    Debug.Print "+++ Loaded tex=" & filename & _
                " (" & format$(desc.size / 1048576#, "0.00") & "MB);" & _
                " Count=" & mTotalTextureCount & _
                " Mem=" & format$(mTotalAllocatedBytes / 1048576#, "0.00") & "MB"

    Set GetTexture = tex
    Exit Function

errhandler:
    Call RegistrarError(Err.Number, Err.Description, "clsTexManager.GetTexture", Erl)
    Resume Next
End Function

'———————————————————————————————————————————————————————————
' Private: evict the first texture not accessed in the last 60s
'———————————————————————————————————————————————————————————
Private Function RemoveLRU() As Boolean
    On Error GoTo errhandler

    Const AGE_THRESHOLD_MS As Long = 60000
    Dim thresholdTick As Long
    thresholdTick = GetTickCount() - AGE_THRESHOLD_MS

    Dim key As Variant
    For Each key In mDictMeta.Keys
        Dim meta As Variant
        meta = mDictMeta.Item(key)
        If meta(0) <= thresholdTick Then
            ' — capture for logging —
            Dim sz As Long
            sz = meta(3)
            ' — release the texture object —
            On Error Resume Next
            Set mDictTextures.Item(key) = Nothing
            On Error GoTo errhandler
            ' — remove from both dicts —
            mDictTextures.Remove key
            mDictMeta.Remove key
            ' — update counters —
            mTotalTextureCount = mDictTextures.count
            mTotalAllocatedBytes = mTotalAllocatedBytes - sz
            Debug.Print "--- Evicted tex=" & key & _
                        " (" & format$(sz / 1048576#, "0.00") & "MB);" & _
                        " Count=" & mTotalTextureCount & _
                        " Mem=" & format$(mTotalAllocatedBytes / 1048576#, "0.00") & "MB"
            RemoveLRU = True
            Exit Function
        End If
    Next

    RemoveLRU = False
    Exit Function

errhandler:
    Call RegistrarError(Err.Number, Err.Description, "clsTexManager.RemoveLRU", Erl)
    Resume Next
End Function



Private Function LoadTexture(ByVal FileName As String, ByRef Dest As Direct3DTexture8) As Long

On Error GoTo LoadTexture_ErrHandler
    Set Dest = Nothing
    Dim bytArr() As Byte

    #If Compresion = 1 Then
        If Not Extract_File_To_Memory(Graphics, App.path & "\..\Recursos\OUTPUT\", LTrim(FileName) & ".png", bytArr, ResourcesPassword) Then
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
        PathToFile = App.path & "\..\Recursos\Graficos\" & LTrim(FileName) & ".png"

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

Public Sub ReleaseUnusedTextures(ByVal num_textures As Integer)
    Dim i As Long
    For i = 0 To num_textures
       Call RemoveLRU
    Next
End Sub

Private Function CreateDirect3DTexture(ByVal FileNum As Integer) As Direct3DTexture8

On Error GoTo ErrHandler

    Dim Texture         As Direct3DTexture8
    Dim loadResult      As Long
    
    Set CreateDirect3dTexture = Nothing
    loadResult = LoadTexture(str(FileNum), Texture)
    
    Select Case loadResult
        Case D3DERR_INVALIDCALL
            Debug.Assert False
            frmDebug.add_text_tracebox "LoadTexture failed with D3DERR_INVALIDCALL"
        Case D3DERR_NOTAVAILABLE
            Debug.Assert False
            frmDebug.add_text_tracebox "LoadTexture failed with D3DERR_NOTAVAILABLE"
        Case D3DXERR_INVALIDDATA
            Debug.Assert False
            frmDebug.add_text_tracebox "LoadTexture failed with D3DXERR_INVALIDDATA"
        Case D3DERR_OUTOFVIDEOMEMORY
        Case E_OUTOFMEMORY
            frmDebug.add_text_tracebox "LoadTexture failed with D3DERR_OUTOFVIDEOMEMORY"
            Call ReleaseUnusedTextures(250)
            'Try to load the texture again, if it fails we've run out of options.
            loadResult = LoadTexture(str(FileNum), Texture)
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


Public Function CreateTexture(ByVal Width As Long, ByVal Height As Long) As Direct3DTexture8
On Error GoTo ErrHandler
    Dim Texture As Direct3DTexture8
    Set Texture = mD3D.CreateTexture(device, Width, Height, 1, 0, D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT)

    Set CreateTexture = Texture
    Exit Function
ErrHandler:
    frmDebug.add_text_tracebox "Failed to generate texture, " & Err.Description
End Function




