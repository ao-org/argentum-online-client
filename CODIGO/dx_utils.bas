Attribute VB_Name = "dx_utils"
Option Explicit

Public Sub list_modes(ByRef d3d As Direct3D8)
    Dim tmpDispMode As D3DDISPLAYMODE
    Dim i As Long
    For i = 0 To d3d.GetAdapterModeCount(0) - 1 'primary adapter
        Call d3d.EnumAdapterModes(0, i, tmpDispMode)
        Debug.Print tmpDispMode.Width & "x" & tmpDispMode.Height
    Next i
End Sub
