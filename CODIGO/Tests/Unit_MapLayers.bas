Attribute VB_Name = "Unit_MapLayers"
' Argentum 20 - Game Client Program
' Copyright (C) 2026 Noland Studios
' Licensed under the GNU Affero General Public License, version 3 or later,
' as described in the repository licence.txt.
Option Explicit

#If UNIT_TEST = 1 Then
Public Function test_suite_map_layers() As Boolean
    Call UnitTesting.RunTest("maplayers_signed_five_layers", Recursos.TestCsmLayers(False, 0))
    Call UnitTesting.RunTest("maplayers_legacy_four_layers", Recursos.TestCsmLayers(True, 0))
    Call UnitTesting.RunTest("maplayers_layer2_walkable_water_explicit_block", Recursos.TestCsmLayers(False, 2))
    Call UnitTesting.RunTest("maplayers_layer3_walkable_water_explicit_block", Recursos.TestCsmLayers(False, 3))
    Call UnitTesting.RunTest("maplayers_overlay_steps_height_priority", TestOverlayBehavior())
    Call UnitTesting.RunTest("maplayers_contiguous_tile_graphics", TestContiguousTiles())
    Call UnitTesting.RunTest("maplayers_installed_resources", Recursos.TestInstalledCsmMaps(App.Path & "\..\Recursos\Mapas"))
    test_suite_map_layers = True
End Function

Private Function TestOverlayBehavior() As Boolean
    On Error GoTo Fail
    ReDim MapData(1 To 100, 1 To 100)
    TestOverlayBehavior = GetTerrenoDePaso(20, GetWalkableOverlayGraphic(50, 50)) = CONST_AGUA
    MapData(50, 50).Graphic(2).GrhIndex = 12682
    TestOverlayBehavior = TestOverlayBehavior And GetTerrainHeight(50, 50) = 5
    TestOverlayBehavior = TestOverlayBehavior And GetTerrenoDePaso(20, GetWalkableOverlayGraphic(50, 50)) = CONST_PISO
    MapData(50, 50).Graphic(3).GrhIndex = 12683
    TestOverlayBehavior = TestOverlayBehavior And GetWalkableOverlayGraphic(50, 50) = 12683
    TestOverlayBehavior = TestOverlayBehavior And GetTerrainHeight(50, 50) = 10
    TestOverlayBehavior = TestOverlayBehavior And GetTerrenoDePaso(20, GetWalkableOverlayGraphic(50, 50)) = CONST_PISO
    MapData(50, 50).Graphic(2).GrhIndex = 0
    TestOverlayBehavior = TestOverlayBehavior And GetTerrainHeight(50, 50) = 10
    Exit Function
Fail:
    TestOverlayBehavior = False
End Function

Private Function TestContiguousTiles() As Boolean
    On Error GoTo Fail
    ReDim MapData(1 To 100, 1 To 100)
    Dim stride As Long, layer As Long
    stride = VarPtr(MapData(1, 1).Graphic(2)) - VarPtr(MapData(1, 1).Graphic(1))
    TestContiguousTiles = stride > 0
    For layer = 2 To MAP_LAYER_COUNT
        TestContiguousTiles = TestContiguousTiles And _
            VarPtr(MapData(1, 1).Graphic(layer)) - VarPtr(MapData(1, 1).Graphic(layer - 1)) = stride
    Next layer
    stride = VarPtr(MapData(2, 1)) - VarPtr(MapData(1, 1))
    TestContiguousTiles = TestContiguousTiles And stride > 0 And _
        VarPtr(MapData(1, 2)) - VarPtr(MapData(1, 1)) = stride * 100
    Exit Function
Fail:
    TestContiguousTiles = False
End Function
#End If
