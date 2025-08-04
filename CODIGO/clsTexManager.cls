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

Const HASH_TABLE_SIZE As Long = 500
Const E_OUTOFMEMORY = 7

Private Type SURFACE_ENTRY_DYN
    FileName As Integer
    UltimoAcceso As Long
    Texture As Direct3DTexture8
    size As Long
    texture_width As Integer
    texture_height As Integer
End Type

Private Type HashNode
    surfaceCount As Integer
    SurfaceEntry() As SURFACE_ENTRY_DYN
End Type

Private TexList(HASH_TABLE_SIZE - 1) As HashNode

Private mD3D                         As D3DX8
Private device                       As Direct3DDevice8
Private mMaxEntries                  As Integer

Private Sub Class_Terminate()
    On Error GoTo Class_Terminate_Err
    Dim i As Long
    Dim j As Long
    'Destroy every surface in memory
    For i = 0 To HASH_TABLE_SIZE - 1
        With TexList(i)
            For j = 1 To .surfaceCount
                Set .SurfaceEntry(j).Texture = Nothing
            Next j
            'Destroy the arrays
            Erase .SurfaceEntry
        End With

    Next i
    Exit Sub

Class_Terminate_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsTexManager.Class_Terminate", Erl)
    Resume Next

End Sub

Public Function GetTexture( _
    ByVal filename As Integer, _
    ByRef textwidth As Long, _
    ByRef textheight As Long _
) As Direct3DTexture8

    On Error GoTo GetTexture_Err
    Debug.Assert FileName > 0

    ' default return
    Set GetTexture = Nothing
    textwidth = 0
    textheight = 0

    '——— 1) Try cache hit ———
    Dim i As Long
    With TexList(FileName Mod HASH_TABLE_SIZE)
        For i = 1 To .surfaceCount
            If .SurfaceEntry(i).FileName = FileName Then
                .SurfaceEntry(i).UltimoAcceso = GetTickCount()
                textwidth = .SurfaceEntry(i).texture_width
                textheight = .SurfaceEntry(i).texture_height
                Set GetTexture = .SurfaceEntry(i).Texture
                Exit Function
            End If
        Next i
    End With

    '——— 2) Cache miss ? actually load & cache now ———
    Dim Texture     As Direct3DTexture8
    Dim surfaceDesc As D3DSURFACE_DESC

    ' call into the bare-bones creator
    Set Texture = CreateDirect3DTexture(filename)

    ' moved in from CreateDirect3DTexture:
    Debug.Assert Not Texture Is Nothing
    If Texture Is Nothing Then
        frmDebug.add_text_tracebox "Failed to load texture file " & filename
        Exit Function
    End If

    ' pull out width/height/size
    Call Texture.GetLevelDesc(0, surfaceDesc)

    ' insert into LRU cache
    With TexList(filename Mod HASH_TABLE_SIZE)
        .surfaceCount = .surfaceCount + 1
        ReDim Preserve .SurfaceEntry(1 To .surfaceCount) As SURFACE_ENTRY_DYN

        With .SurfaceEntry(.surfaceCount)
            .filename = filename
            .UltimoAcceso = GetTickCount()
            Set .Texture = Texture
            .texture_width = surfaceDesc.Width
            .texture_height = surfaceDesc.Height
            .size = surfaceDesc.size
        End With
    End With

    ' now fill the out parameters
    textwidth = surfaceDesc.Width
    textheight = surfaceDesc.Height

    ' finally return the new texture
    Set GetTexture = Texture
    Exit Function

GetTexture_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsTexManager.GetTexture", Erl)
    Resume Next
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
            Call ReleaseUnusedTextures(10)
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

Private Function RemoveLRU() As Boolean
    On Error GoTo RemoveLRU_Err

    Const AGE_THRESHOLD_MS As Long = 60000   ' 60 seconds
    Dim nowTick      As Long
    Dim oldestTick   As Long
    Dim bucketIdx    As Long
    Dim entryIdx     As Long
    Dim i As Long, J As Long

    nowTick = GetTickCount()
    oldestTick = nowTick - AGE_THRESHOLD_MS

    ' Find the absolute oldest entry that is at least 60s stale
    bucketIdx = -1
    entryIdx = 0

    For i = 0 To HASH_TABLE_SIZE - 1
        With TexList(i)
            For J = 1 To .surfaceCount
                With .SurfaceEntry(J)
                    ' candidate if more than threshold old and valid size
                    If .size > 0 And .UltimoAcceso <= oldestTick Then
                        ' choose the very oldest among those candidates
                        If bucketIdx = -1 Or .UltimoAcceso < oldestTick Then
                            bucketIdx = i
                            entryIdx = J
                            oldestTick = .UltimoAcceso
                        End If
                    End If
                End With
            Next J
        End With
    Next i

    ' If we found none that are >= 60s old, do not evict
    If bucketIdx = -1 Or entryIdx = 0 Then
        RemoveLRU = False
        Exit Function
    End If

    ' Evict that entry
    With TexList(bucketIdx).SurfaceEntry(entryIdx)
        ' Debug log if you like:
        frmDebug.add_text_tracebox "Evicting texture FileID=" & .filename & " last used " & format(((nowTick - .UltimoAcceso) / 1000), "0.0") & "s ago"
        Set .Texture = Nothing
        .filename = 0
    End With

    ' Shift entries down and shrink array
    With TexList(bucketIdx)
        Dim k As Long
        For k = entryIdx To .surfaceCount - 1
            .SurfaceEntry(k) = .SurfaceEntry(k + 1)
        Next k
        .surfaceCount = .surfaceCount - 1
        If .surfaceCount > 0 Then
            ReDim Preserve .SurfaceEntry(1 To .surfaceCount) As SURFACE_ENTRY_DYN
        Else
            Erase .SurfaceEntry
        End If
    End With

    RemoveLRU = True
    Exit Function

RemoveLRU_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsTexManager.RemoveLRU", Erl)
    Resume Next
End Function

