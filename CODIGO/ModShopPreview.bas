Attribute VB_Name = "ModShopPreview"
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
Option Explicit

Private Type tShopPreviewInfo
    BodyOverride(1 To eRaza.Orco, 1 To eGenero.Mujer) As Integer
    DefaultBody As Integer
End Type

Private gShopPreview() As tShopPreviewInfo
Private gPreviewInitialized As Boolean

Public Sub ShopPreview_Reset(ByVal maxObj As Long)
    On Error GoTo ShopPreview_Reset_Err

    If maxObj < 0 Then maxObj = 0
    ReDim gShopPreview(0 To maxObj) As tShopPreviewInfo
    gPreviewInitialized = True
    Exit Sub

ShopPreview_Reset_Err:
    gPreviewInitialized = False
    Call RegistrarError(Err.Number, Err.Description, "ModShopPreview.ShopPreview_Reset", Erl)
    Resume Next
End Sub

Public Sub ShopPreview_SetDefaultBody(ByVal objNum As Long, ByVal bodyIndex As Integer)
    On Error GoTo ShopPreview_SetDefaultBody_Err

    If Not gPreviewInitialized Then Exit Sub
    If objNum < LBound(gShopPreview) Or objNum > UBound(gShopPreview) Then Exit Sub

    gShopPreview(objNum).DefaultBody = bodyIndex
    Exit Sub

ShopPreview_SetDefaultBody_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModShopPreview.ShopPreview_SetDefaultBody", Erl)
    Resume Next
End Sub

Public Sub ShopPreview_SetBodyOverride(ByVal objNum As Long, ByVal race As eRaza, ByVal gender As eGenero, ByVal bodyIndex As Integer)
    On Error GoTo ShopPreview_SetBodyOverride_Err

    If Not gPreviewInitialized Then Exit Sub
    If objNum < LBound(gShopPreview) Or objNum > UBound(gShopPreview) Then Exit Sub
    If race < LBound(gShopPreview(objNum).BodyOverride, 1) Or race > UBound(gShopPreview(objNum).BodyOverride, 1) Then Exit Sub
    If gender < LBound(gShopPreview(objNum).BodyOverride, 2) Or gender > UBound(gShopPreview(objNum).BodyOverride, 2) Then Exit Sub

    gShopPreview(objNum).BodyOverride(race, gender) = bodyIndex
    Exit Sub

ShopPreview_SetBodyOverride_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModShopPreview.ShopPreview_SetBodyOverride", Erl)
    Resume Next
End Sub

Public Function ShopPreview_GetBodyOverride(ByVal objNum As Long, ByVal race As eRaza, ByVal gender As eGenero) As Integer
    On Error GoTo ShopPreview_GetBodyOverride_Err

    If Not gPreviewInitialized Then Exit Function
    If objNum < LBound(gShopPreview) Or objNum > UBound(gShopPreview) Then Exit Function
    If race < LBound(gShopPreview(objNum).BodyOverride, 1) Or race > UBound(gShopPreview(objNum).BodyOverride, 1) Then Exit Function
    If gender < LBound(gShopPreview(objNum).BodyOverride, 2) Or gender > UBound(gShopPreview(objNum).BodyOverride, 2) Then Exit Function

    ShopPreview_GetBodyOverride = gShopPreview(objNum).BodyOverride(race, gender)
    Exit Function

ShopPreview_GetBodyOverride_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModShopPreview.ShopPreview_GetBodyOverride", Erl)
    ShopPreview_GetBodyOverride = 0
End Function

Public Function ShopPreview_GetDefaultBody(ByVal objNum As Long) As Integer
    On Error GoTo ShopPreview_GetDefaultBody_Err

    If Not gPreviewInitialized Then Exit Function
    If objNum < LBound(gShopPreview) Or objNum > UBound(gShopPreview) Then Exit Function

    ShopPreview_GetDefaultBody = gShopPreview(objNum).DefaultBody
    Exit Function

ShopPreview_GetDefaultBody_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModShopPreview.ShopPreview_GetDefaultBody", Erl)
    ShopPreview_GetDefaultBody = 0
End Function

Public Sub ShopPreview_RegisterArmorBodies(ByVal objNum As Long, ByRef reader As clsIniManager)
    On Error GoTo ShopPreview_RegisterArmorBodies_Err

    If reader Is Nothing Then Exit Sub

    Call ShopPreview_SetBodyOverride(objNum, eRaza.Humano, eGenero.Hombre, val(reader.GetValue("OBJ" & objNum, "RopajeHumano")))
    Call ShopPreview_SetBodyOverride(objNum, eRaza.Humano, eGenero.Mujer, val(reader.GetValue("OBJ" & objNum, "RopajeHumana")))
    Call ShopPreview_SetBodyOverride(objNum, eRaza.Elfo, eGenero.Hombre, val(reader.GetValue("OBJ" & objNum, "RopajeElfo")))
    Call ShopPreview_SetBodyOverride(objNum, eRaza.Elfo, eGenero.Mujer, val(reader.GetValue("OBJ" & objNum, "RopajeElfa")))
    Call ShopPreview_SetBodyOverride(objNum, eRaza.ElfoOscuro, eGenero.Hombre, val(reader.GetValue("OBJ" & objNum, "RopajeElfoOscuro")))
    Call ShopPreview_SetBodyOverride(objNum, eRaza.ElfoOscuro, eGenero.Mujer, val(reader.GetValue("OBJ" & objNum, "RopajeElfaOscura")))
    Call ShopPreview_SetBodyOverride(objNum, eRaza.Gnomo, eGenero.Hombre, val(reader.GetValue("OBJ" & objNum, "RopajeGnomo")))
    Call ShopPreview_SetBodyOverride(objNum, eRaza.Gnomo, eGenero.Mujer, val(reader.GetValue("OBJ" & objNum, "RopajeGnoma")))
    Call ShopPreview_SetBodyOverride(objNum, eRaza.Enano, eGenero.Hombre, val(reader.GetValue("OBJ" & objNum, "RopajeEnano")))
    Call ShopPreview_SetBodyOverride(objNum, eRaza.Enano, eGenero.Mujer, val(reader.GetValue("OBJ" & objNum, "RopajeEnana")))
    Call ShopPreview_SetBodyOverride(objNum, eRaza.Orco, eGenero.Hombre, val(reader.GetValue("OBJ" & objNum, "RopajeOrco")))
    Call ShopPreview_SetBodyOverride(objNum, eRaza.Orco, eGenero.Mujer, val(reader.GetValue("OBJ" & objNum, "RopajeOrca")))
    Exit Sub

ShopPreview_RegisterArmorBodies_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModShopPreview.ShopPreview_RegisterArmorBodies", Erl)
    Resume Next
End Sub
