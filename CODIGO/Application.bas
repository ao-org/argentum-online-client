Attribute VB_Name = "Application"
'RevolucionAo 1.0
'Pablo Mercavides
'**************************************************************************

Option Explicit

Private Declare Function GetActiveWindow Lib "user32" () As Long

''
' Checks if this is the active (foreground) application or not.
'
' @return   True if any of the app's windows are the foreground window, false otherwise.

Public Function IsAppActive() As Boolean
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (maraxus)
    'Last Modify Date: 03/03/2007
    'Checks if this is the active application or not
    '***************************************************
    
    On Error GoTo IsAppActive_Err
    
    IsAppActive = (GetActiveWindow <> 0)

    
    Exit Function

IsAppActive_Err:
    Call RegistrarError(Err.number, Err.Description, "Application.IsAppActive", Erl)
    Resume Next
    
End Function
