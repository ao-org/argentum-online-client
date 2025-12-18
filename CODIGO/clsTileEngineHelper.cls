VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTileEngineHelper"
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
Private ValidTreeIndices() As Variant

Function IsIndexTree(ByVal GrhIndex As Long) As Boolean
    On Error Resume Next
    IsIndexTree = False

    ' Si no se cargaron indices, no hay arboles que verificar
    If (Not Not ValidTreeIndices) = 0 Then Exit Function

    Dim i As Long
    For i = LBound(ValidTreeIndices) To UBound(ValidTreeIndices)
        If GrhIndex = ValidTreeIndices(i) Then
            IsIndexTree = True
            Exit Function
        End If
    Next i
End Function

Private Sub Class_Initialize()
    Call LoadTreeIndices
End Sub

Private Sub LoadTreeIndices()
    Dim IniPath As String
    IniPath = App.path & "\..\Recursos\init\EsArbol.ini"

    On Error GoTo Cleanup

    If FileExist(IniPath, vbNormal) Then
        Dim Reader As New clsIniManager
        Call Reader.Initialize(IniPath)

        Dim RawIndices As String
        RawIndices = Reader.GetValue("ARBOL", "Indices")

        If LenB(RawIndices) <> 0 Then
            Dim Parts() As String
            Parts = Split(RawIndices, ",")

            Dim i As Long
            ReDim ValidTreeIndices(LBound(Parts) To UBound(Parts))

            For i = LBound(Parts) To UBound(Parts)
                ValidTreeIndices(i) = CLng(Trim$(Parts(i)))
            Next i

            Exit Sub
        End If
    End If

Cleanup:
    ' Si la lectura falla, dejamos el array vacio para evitar listas desactualizadas
    ValidTreeIndices = Array()
End Sub
