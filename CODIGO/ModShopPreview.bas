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

Private Type tShopBodyLookup
    Race As eRaza
    Gender As eGenero
End Type

Private gShopPreview() As tShopPreviewInfo
Private gPreviewInitialized As Boolean
Private gBodyLookup() As tShopBodyLookup

Public Sub ShopPreview_Reset(ByVal maxObj As Long)
    On Error GoTo ShopPreview_Reset_Err

    If maxObj < 0 Then maxObj = 0
    ReDim gShopPreview(0 To maxObj) As tShopPreviewInfo
    gPreviewInitialized = True
    Erase gBodyLookup
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

Private Sub EnsureBodyLookupSize(ByVal bodyIndex As Long)
    On Error GoTo EnsureBodyLookupSize_Err

    If bodyIndex <= 0 Then Exit Sub

    If IsBodyLookupInitialized() Then
        If bodyIndex <= UBound(gBodyLookup) Then Exit Sub
    End If

    Dim newSize As Long
    If IsBodyLookupInitialized() Then
        newSize = UBound(gBodyLookup)
    Else
        newSize = 0
    End If

    If bodyIndex > newSize Then newSize = bodyIndex
    If IsBodyLookupInitialized() Then
        ReDim Preserve gBodyLookup(0 To newSize) As tShopBodyLookup
    Else
        ReDim gBodyLookup(0 To newSize) As tShopBodyLookup
    End If
    Exit Sub

EnsureBodyLookupSize_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModShopPreview.EnsureBodyLookupSize", Erl)
    Resume Next
End Sub

Private Sub ShopPreview_RegisterBodyMapping(ByVal bodyIndex As Integer, ByVal race As eRaza, ByVal gender As eGenero)
    On Error GoTo ShopPreview_RegisterBodyMapping_Err

    If bodyIndex <= 0 Then Exit Sub
    If race < eRaza.Humano Or race > eRaza.Orco Then Exit Sub
    If gender < eGenero.Hombre Or gender > eGenero.Mujer Then Exit Sub

    Call EnsureBodyLookupSize(bodyIndex)
    If IsBodyLookupInitialized() Then
        gBodyLookup(bodyIndex).Race = race
        gBodyLookup(bodyIndex).Gender = gender
    End If
    Exit Sub

ShopPreview_RegisterBodyMapping_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModShopPreview.ShopPreview_RegisterBodyMapping", Erl)
    Resume Next
End Sub

Public Function ShopPreview_GuessRaceGender(ByVal bodyIndex As Integer, ByRef race As eRaza, ByRef gender As eGenero) As Boolean
    On Error GoTo ShopPreview_GuessRaceGender_Err

    If bodyIndex < 0 Then Exit Function

    If IsBodyLookupInitialized() Then
        If bodyIndex >= LBound(gBodyLookup) And bodyIndex <= UBound(gBodyLookup) Then
            If gBodyLookup(bodyIndex).Race >= eRaza.Humano And gBodyLookup(bodyIndex).Race <= eRaza.Orco Then
                If gBodyLookup(bodyIndex).Gender >= eGenero.Hombre And gBodyLookup(bodyIndex).Gender <= eGenero.Mujer Then
                    race = gBodyLookup(bodyIndex).Race
                    gender = gBodyLookup(bodyIndex).Gender
                    ShopPreview_GuessRaceGender = True
                End If
            End If
        End If
    End If
    Exit Function

ShopPreview_GuessRaceGender_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModShopPreview.ShopPreview_GuessRaceGender", Erl)
    ShopPreview_GuessRaceGender = False
End Function

Private Function IsBodyLookupInitialized() As Boolean
    On Error GoTo IsBodyLookupInitialized_Err

    Dim lb As Long
    lb = LBound(gBodyLookup)
    IsBodyLookupInitialized = True
    Exit Function

IsBodyLookupInitialized_Err:
    IsBodyLookupInitialized = False
End Function

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
    If objNum < LBound(ObjData) Or objNum > UBound(ObjData) Then Exit Sub

    Dim overrides(1 To eRaza.Orco, 1 To eGenero.Mujer) As Integer
    overrides(eRaza.Humano, eGenero.Hombre) = Val(reader.GetValue("OBJ" & objNum, "RopajeHumano"))
    overrides(eRaza.Humano, eGenero.Mujer) = Val(reader.GetValue("OBJ" & objNum, "RopajeHumana"))
    overrides(eRaza.Elfo, eGenero.Hombre) = Val(reader.GetValue("OBJ" & objNum, "RopajeElfo"))
    overrides(eRaza.Elfo, eGenero.Mujer) = Val(reader.GetValue("OBJ" & objNum, "RopajeElfa"))
    overrides(eRaza.ElfoOscuro, eGenero.Hombre) = Val(reader.GetValue("OBJ" & objNum, "RopajeElfoOscuro"))
    overrides(eRaza.ElfoOscuro, eGenero.Mujer) = Val(reader.GetValue("OBJ" & objNum, "RopajeElfaOscura"))
    overrides(eRaza.Gnomo, eGenero.Hombre) = Val(reader.GetValue("OBJ" & objNum, "RopajeGnomo"))
    overrides(eRaza.Gnomo, eGenero.Mujer) = Val(reader.GetValue("OBJ" & objNum, "RopajeGnoma"))
    overrides(eRaza.Enano, eGenero.Hombre) = Val(reader.GetValue("OBJ" & objNum, "RopajeEnano"))
    overrides(eRaza.Enano, eGenero.Mujer) = Val(reader.GetValue("OBJ" & objNum, "RopajeEnana"))
    overrides(eRaza.Orco, eGenero.Hombre) = Val(reader.GetValue("OBJ" & objNum, "RopajeOrco"))
    overrides(eRaza.Orco, eGenero.Mujer) = Val(reader.GetValue("OBJ" & objNum, "RopajeOrca"))

    Dim race As eRaza
    Dim gender As eGenero
    Dim defaultBody As Integer
    Dim bodyIndex As Integer

    For race = eRaza.Humano To eRaza.Orco
        For gender = eGenero.Hombre To eGenero.Mujer
            bodyIndex = overrides(race, gender)
            ObjData(objNum).PreviewBody(race, gender) = bodyIndex
            Call ShopPreview_SetBodyOverride(objNum, race, gender, bodyIndex)
            Call ShopPreview_RegisterBodyMapping(bodyIndex, race, gender)
            If defaultBody = 0 And bodyIndex > 0 Then
                defaultBody = bodyIndex
            End If
        Next gender
    Next race

    ObjData(objNum).PreviewDefaultBody = defaultBody
    Call ShopPreview_SetDefaultBody(objNum, defaultBody)
    Exit Sub

ShopPreview_RegisterArmorBodies_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModShopPreview.ShopPreview_RegisterArmorBodies", Erl)
    Resume Next
End Sub
