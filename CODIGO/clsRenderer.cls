VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRenderer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'    Argentum 20 - Game Client Program
'    Copyright (C) 2023 - Noland Studios
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

Option Base 0

Option Explicit


Private mDirectX As New DirectX8
Private mDirectD3D8 As D3DX8
Private mDirectD3D As Direct3D8

Private mDirectDevice As Direct3DDevice8

Private mD3DPP As D3DPRESENT_PARAMETERS

Private mProjection As D3DMATRIX
Private mIdentityMatrix As D3DMATRIX

Const ScreenWidth  As Integer = 1024
Const ScreenHeight As Integer = 768


Private Type NEW_TYPE_VERTEX
    x       As Single
    y       As Single
    z       As Single
    color   As RGBA
    u      As Single
    v      As Single
End Type

Private mDisplayMode As D3DDISPLAYMODE

Private Sub get_game_resolution(ByRef mode As D3DDISPLAYMODE)
    On Error Goto get_game_resolution_Err
    'For the time being we hard code it to 1024x768x32
    mode.Width = 1024
    mode.Height = 768
    mode.RefreshRate = 60
    mode.format = D3DFMT_A8R8G8B8
    Exit Sub
get_game_resolution_Err:
    Call TraceError(Err.Number, Err.Description, "clsRenderer.get_game_resolution", Erl)
End Sub

Public Sub list_modes(ByRef d3d As Direct3D8)
    On Error Goto list_modes_Err
    Dim tmpDispMode As D3DDISPLAYMODE
    Dim i As Long
    For i = 0 To d3d.GetAdapterModeCount(0) - 1 'primary adapter
        Call d3d.EnumAdapterModes(0, i, tmpDispMode)
        frmDebug.add_text_tracebox tmpDispMode.Width & "x" & tmpDispMode.Height & " fmt:" & tmpDispMode.format
        
    Next i
    Exit Sub
list_modes_Err:
    Call TraceError(Err.Number, Err.Description, "clsRenderer.list_modes", Erl)
End Sub
Private Function init_dx8_and_d3d() As Long
    On Error Goto init_dx8_and_d3d_Err
On Error Resume Next
    Err.Clear
    Set mDirectX = New DirectX8
    If Err.Number <> 0 Then
                Call MsgBox(JsonLanguage.Item("MENSAJEBOX_FATAL_ERROR_DIRECTX8"), vbCritical, App.title)
                frmDebug.add_text_tracebox "Error Number Returned: " & Err.Number
                Exit Function
    End If
    
    Set mDirectD3D = DirectX.Direct3DCreate()
    If Err.Number <> 0 Then
                Call MsgBox(JsonLanguage.Item("MENSAJEBOX_FATAL_ERROR_DIRECTD3D"), vbCritical, App.title)
                frmDebug.add_text_tracebox "Error Number Returned: " & Err.Number
                Exit Function
    End If
    
    Set mDirectD3D8 = New D3DX8
    If Err.Number <> 0 Then
                Call MsgBox(JsonLanguage.Item("MENSAJEBOX_FATAL_ERROR_DIRECTD3D8"), vbCritical, App.title)
                frmDebug.add_text_tracebox "Error Number Returned: " & Err.Number
                Exit Function
    End If
    init_dx8_and_d3d = Err.Number
    Exit Function
init_dx8_and_d3d_Err:
    Call TraceError(Err.Number, Err.Description, "clsRenderer.init_dx8_and_d3d", Erl)
End Function

Public Function init(ByVal window As Long) As Long
    On Error Goto init_Err
    init = init_dx8_and_d3d()
    If init <> 0 Then
        frmDebug.add_text_tracebox "init_dx8_and_d3d() failed with code " & Init
        Exit Function
    End If
    init = init_dx_device(window, mDirectD3D, mDisplayMode)
    If init <> 0 Then
        frmDebug.add_text_tracebox "init_dx_device() failed with code " & Init
        Exit Function
    End If
    
    Exit Function
init_Err:
    Call TraceError(Err.Number, Err.Description, "clsRenderer.init", Erl)
End Function

Private Function init_dx_device(ByVal window_handle As Long, _
    On Error Goto init_dx_device_Err
                                ByRef d3d_obj As Direct3D8, _
                                ByRef disp_mode As D3DDISPLAYMODE) As Long
On Error Resume Next
    Debug.Assert Not d3d_obj Is Nothing
    Dim Caps As D3DCAPS8
    Dim DevType As CONST_D3DDEVTYPE
    Dim DevBehaviorFlags As Long
    Dim d3dDispMode  As D3DDISPLAYMODE
    Err.Clear
    d3d_obj.GetDeviceCaps D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Caps
    If Err.Number = D3DERR_NOTAVAILABLE Then
        frmDebug.add_text_tracebox "HAL Is Not available, using; software; vertex; processing"
        DevType = D3DDEVTYPE_REF
        DevBehaviorFlags = D3DCREATE_SOFTWARE_VERTEXPROCESSING
    Else
        DevType = D3DDEVTYPE_HAL
        frmDebug.add_text_tracebox "VertexProcessingCaps = " & Caps.VertexProcessingCaps
        If Caps.VertexProcessingCaps = 0 Then
            frmDebug.add_text_tracebox "HAL Is available, " & "Using; software; vertex; processing"
            DevBehaviorFlags = D3DCREATE_SOFTWARE_VERTEXPROCESSING
        ElseIf Caps.VertexProcessingCaps = &H4B Then
            frmDebug.add_text_tracebox "HAL Is available, " & "Using; hardware; vertex; processing; "
            DevBehaviorFlags = D3DCREATE_HARDWARE_VERTEXPROCESSING
        Else
            frmDebug.add_text_tracebox "HAL Is available, " & "Using; mixed; vertex; processing; "
            DevBehaviorFlags = D3DCREATE_MIXED_VERTEXPROCESSING
        End If
    End If
    d3d_obj.GetAdapterDisplayMode D3DADAPTER_DEFAULT, d3dDispMode
    
    frmDebug.add_text_tracebox "Using; Windowed; mode"
    Debug.Assert disp_mode.Width > 0
    Debug.Assert disp_mode.Height > 0
    mD3DPP.Windowed = 1
    mD3DPP.BackBufferWidth = disp_mode.Width
    mD3DPP.BackBufferHeight = disp_mode.Height
    mD3DPP.BackBufferFormat = disp_mode.format
    mD3DPP.SwapEffect = D3DSWAPEFFECT_DISCARD
    mD3DPP.BackBufferCount = 1
    mD3DPP.hDeviceWindow = window_handle
    mD3DPP.EnableAutoDepthStencil = 1
    mD3DPP.AutoDepthStencilFormat = D3DFMT_D16
    
    Err.Clear
 
    Set mDirectDevice = d3d_obj.CreateDevice(D3DADAPTER_DEFAULT, DevType, window_handle, DevBehaviorFlags, mD3DPP)
    frmDebug.add_text_tracebox "Create Direct3D device: " & Err
    If (Err.Number <> 0) Then
        'if we failed to create the device with D3DFMT_A8R8G8B8 we try to do it with current display fmt
        D3DWindow.BackBufferFormat = d3dDispMode.format
        Err.Clear
        Set mDirectDevice = d3d_obj.CreateDevice(D3DADAPTER_DEFAULT, DevType, window_handle, DevBehaviorFlags, mD3DPP)
        frmDebug.add_text_tracebox "Create Direct3D device: " & Err
    End If
    init_dx_device = Err.Number
    Exit Function
init_dx_device_Err:
    Call TraceError(Err.Number, Err.Description, "clsRenderer.init_dx_device", Erl)
End Function




Private Sub Class_Initialize()
    On Error Goto Class_Initialize_Err
    Set mDirectX = Nothing
    Set mDirectD3D8 = Nothing
    Set mDirectD3D = Nothing
    Set mDirectDevice = Nothing
    mDisplayMode.Width = ScreenWidth
    mDisplayMode.Height = ScreenHeight
    mDisplayMode.RefreshRate = 60
    mDisplayMode.format = D3DFMT_A8R8G8B8
    Exit Sub
Class_Initialize_Err:
    Call TraceError(Err.Number, Err.Description, "clsRenderer.Class_Initialize", Erl)
End Sub
