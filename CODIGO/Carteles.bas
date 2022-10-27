Attribute VB_Name = "Carteles"
Option Explicit

'Carteles
Public cartel    As Boolean

Public Leyenda   As String

Public GrhCartel As Integer

Sub InitCartel(Ley As String, grh As Integer)
    
    On Error GoTo InitCartel_Err
    

    If Not cartel Then
        Leyenda = Ley
        GrhCartel = grh
        cartel = True
    Else
        Exit Sub

    End If

    
    Exit Sub

InitCartel_Err:
    Call RegistrarError(Err.number, Err.Description, "Carteles.InitCartel", Erl)
    Resume Next
    
End Sub

