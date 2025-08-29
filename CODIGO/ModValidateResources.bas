Attribute VB_Name = "ModValidateResources"

Public Function ValidateResources() As Boolean
    On Error Goto ValidateResources_Err
    ValidateResources = True
    Exit Function
ValidateResources_Err:
    Call TraceError(Err.Number, Err.Description, "ModValidateResources.ValidateResources", Erl)
End Function
