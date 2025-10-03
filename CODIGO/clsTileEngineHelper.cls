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
    Dim i As Long
    For i = LBound(ValidTreeIndices) To UBound(ValidTreeIndices)
        If GrhIndex = ValidTreeIndices(i) Then
            IsIndexTree = True
            Exit Function
        End If
    Next i
End Function

Private Sub Class_Initialize()
    ValidTreeIndices = Array(643, 644, 647, 735, 1121, 2931, 11903, 11904, 11905, 14775, 11906, 70885, 70884, 71042, 71041, 15698, 14504, 14505, 15697, 15510, 12581, 12582, _
            12583, 12584, 12585, 12586, 12164, 12165, 12166, 12167, 12168, 12169, 12170, 12171, 12172, 12173, 12174, 12175, 12176, 12177, 12178, 12179, 32142, 32143, 32144, _
            32145, 32146, 32147, 32148, 32149, 32150, 32151, 32152, 32154, 55626, 55627, 55628, 55629, 55630, 55631, 55632, 55633, 55634, 55635, 55636, 55637, 55638, 55639, _
            55640, 55642, 50985, 50986, 50987, 50988, 50989, 50990, 50991, 2547, 2548, 2549, 6597, 6598, 15108, 15109, 15110, 12160, 7220, 462, 463, 1877, 1878, 1879, 1880, _
            1881, 1890, 1892, 433, 460, 461, 9513, 9514, 9515, 9518, 9519, 9520, 9529, 14687, 47726, 12333, 12330, 20369, 21120, 21227, 21352, 12332, 21226, 8258, 32118, 32119, _
            32129, 32132, 32133, 32135)
End Sub
