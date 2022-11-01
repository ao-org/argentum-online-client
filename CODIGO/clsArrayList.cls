VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsArrayList"
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
'***************************************************
'Author: FrankoH298
'Count ahora se llama GetLastItem (Recox)
' Funciones:
'           - Clear() - Limpia la lista
'           - itemExist(item) - Te indica si el item existe
'           - itemPos(item) - Te devuelve la posicion del item
'           - Item(index) - Te devuelve el item en la posicion
'           - Add() - Agrega un item a la lista
'           - Remove() - Remueve el item de la lista
'           - GetLastItem() - Te muestra el ultimo item de la lista
'***************************************************

Option Explicit

Private ArrayList() As Byte

Public Sub Clear()
    
    On Error GoTo Clear_Err
    

    '***************************************************
    'Author: FrankoH298
    'Last Modify Date: 07/09/2019
    ' Limpia la lista
    '***************************************************

    Dim i As Long

    For i = LBound(ArrayList) To UBound(ArrayList)
        ArrayList(i) = 0
    Next i

    
    Exit Sub

Clear_Err:
    Call RegistrarError(Err.number, Err.Description, "clsArrayList.Clear", Erl)
    Resume Next
    
End Sub

Public Sub Initialize(ByVal min As Integer, ByVal max As Integer)
    
    On Error GoTo Initialize_Err
    

    '***************************************************
    'Author: FrankoH298
    'Last Modify Date: 07/09/2019
    ' Inicializa la clase
    '***************************************************

    ReDim ArrayList(min To max) As Byte
    Me.Clear

    
    Exit Sub

Initialize_Err:
    Call RegistrarError(Err.number, Err.Description, "clsArrayList.Initialize", Erl)
    Resume Next
    
End Sub

Public Function itemExist(ByVal Item As Integer) As Boolean
    
    On Error GoTo itemExist_Err
    

    '***************************************************
    'Author: FrankoH298
    'Last Modify Date: 08/09/2019
    ' Devuelve si existe el item
    '***************************************************

    Dim i As Long

    For i = LBound(ArrayList) To UBound(ArrayList)

        If ArrayList(i) = Item Then
            itemExist = True
            Exit Function

        End If

    Next i

    itemExist = False

    
    Exit Function

itemExist_Err:
    Call RegistrarError(Err.number, Err.Description, "clsArrayList.itemExist", Erl)
    Resume Next
    
End Function

Public Function itemPos(ByVal Item As Integer) As Byte
    
    On Error GoTo itemPos_Err
    

    '***************************************************
    'Author: FrankoH298
    'Last Modify Date: 08/09/2019
    ' Devuelve la posicion del item
    '***************************************************

    Dim i As Long

    For i = LBound(ArrayList) To UBound(ArrayList)

        If ArrayList(i) = Item Then
            itemPos = i
            Exit Function

        End If

    Next i

    itemPos = 0

    
    Exit Function

itemPos_Err:
    Call RegistrarError(Err.number, Err.Description, "clsArrayList.itemPos", Erl)
    Resume Next
    
End Function

Public Function Item(ByVal Index As Integer) As Byte
    
    On Error GoTo Item_Err
    

    '***************************************************
    'Author: FrankoH298
    'Last Modify Date: 08/09/2019
    ' Devuelve que item esta en esa posicion
    '***************************************************

    Item = ArrayList(Index)

    
    Exit Function

Item_Err:
    Call RegistrarError(Err.number, Err.Description, "clsArrayList.Item", Erl)
    Resume Next
    
End Function

Public Sub Add(ByVal Item As Integer)
    
    On Error GoTo Add_Err
    

    '***************************************************
    'Author: FrankoH298
    'Last Modify Date: 07/09/2019
    ' Agrega el item a la lista
    '***************************************************

    Dim i As Long

    For i = LBound(ArrayList) To UBound(ArrayList)

        If ArrayList(i) = 0 Then
            ArrayList(i) = Item
            Exit Sub

        End If

    Next i

    
    Exit Sub

Add_Err:
    Call RegistrarError(Err.number, Err.Description, "clsArrayList.Add", Erl)
    Resume Next
    
End Sub

Public Sub Remove(ByVal Item As Integer)
    
    On Error GoTo Remove_Err
    

    '***************************************************
    'Author: FrankoH298
    'Last Modify Date: 07/09/2019
    ' Remueve el item de la lista
    '***************************************************

    Dim i As Long

    Dim A As Long

    For i = LBound(ArrayList) To (UBound(ArrayList) - 1)

        If ArrayList(i) = Item Then

            For A = i To (UBound(ArrayList) - 1)
                ArrayList(A) = ArrayList(A + 1)
            Next A

            ArrayList(UBound(ArrayList)) = 0
            Exit Sub

        End If

    Next i

    If ArrayList(UBound(ArrayList)) = Item Then ArrayList(UBound(ArrayList)) = 0

    
    Exit Sub

Remove_Err:
    Call RegistrarError(Err.number, Err.Description, "clsArrayList.Remove", Erl)
    Resume Next
    
End Sub

Public Function GetLastItem() As Byte
    
    On Error GoTo GetLastItem_Err
    

    '***************************************************
    'Author: FrankoH298
    'Last Modify Date: 08/09/2019
    ' Devuelve el ultimo objeto de la lista que sea diferente a 0
    '***************************************************

    Dim i As Long

    For i = UBound(ArrayList) To LBound(ArrayList) Step -1

        If ArrayList(i) <> 0 Then
            GetLastItem = ArrayList(i)
            Exit Function

        End If

    Next i

    GetLastItem = 0

    
    Exit Function

GetLastItem_Err:
    Call RegistrarError(Err.number, Err.Description, "clsArrayList.GetLastItem", Erl)
    Resume Next
    
End Function

