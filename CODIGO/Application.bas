Attribute VB_Name = "Application"
'RevolucionAo 1.0
'Pablo Mercavides
'**************************************************************************

Option Explicit

Private Declare Function GetActiveWindow Lib "user32" () As Long

Private Type UltimoError
    Componente As String
    Contador As Byte
    ErrorCode As Long
End Type: Private HistorialError As UltimoError

' Obtener carpetas especiales de Windows
Private Const CSIDL_DESKTOP = &H0 '// The Desktop - virtual folder
Private Const CSIDL_PROGRAMS = 2 '// Program Files
Private Const CSIDL_CONTROLS = 3 '// Control Panel - virtual folder
Private Const CSIDL_PRINTERS = 4 '// Printers - virtual folder
Private Const CSIDL_DOCUMENTS = 5 '// My Documents
Private Const CSIDL_FAVORITES = 6 '// Favourites
Private Const CSIDL_STARTUP = 7 '// Startup Folder
Private Const CSIDL_RECENT = 8 '// Recent Documents
Private Const CSIDL_SENDTO = 9 '// Send To Folder
Private Const CSIDL_BITBUCKET = 10 '// Recycle Bin - virtual folder
Private Const CSIDL_STARTMENU = 11 '// Start Menu
Private Const CSIDL_DESKTOPFOLDER = 16 '// Desktop folder
Private Const CSIDL_DRIVES = 17 '// My Computer - virtual folder
Private Const CSIDL_NETWORK = 18 '// Network Neighbourhood - virtual folder
Private Const CSIDL_NETHOOD = 19 '// NetHood Folder
Private Const CSIDL_FONTS = 20 '// Fonts folder
Private Const CSIDL_SHELLNEW = 21 '// ShellNew folder

Private Const MAX_PATH = 260
Private Const NOERROR = 0

Public CARPETA_LOGS As String

Private Type shiEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As shiEMID
End Type

Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long


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

Code:

Public Function GetSpecialfolder(CSIDL As Long) As String
    Dim IDL     As ITEMIDLIST
    Dim sPath   As String
    Dim iReturn As Long
    
    iReturn = SHGetSpecialFolderLocation(100, CSIDL, IDL)
    
    If iReturn = NOERROR Then
        sPath = Space(512)
        iReturn = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath)
        sPath = RTrim$(sPath)

        If Asc(Right(sPath, 1)) = 0 Then sPath = Left$(sPath, Len(sPath) - 1)
        
        GetSpecialfolder = sPath
        
        Exit Function

    End If

    GetSpecialfolder = vbNullString

End Function

Public Sub RegistrarError(ByVal Numero As Long, ByVal Descripcion As String, ByVal Componente As String, Optional ByVal Linea As Integer)
'**********************************************************
'Author: Jopi
'Guarda una descripcion detallada del error en Errores.log
'**********************************************************
    
    On Error GoTo EH:
    
    ' Si no existe la carpeta, la creamos.
    If Not FileExist(CARPETA_LOGS, vbDirectory) Then Call MkDir(CARPETA_LOGS)
    
    'Si lo del parametro Componente es ES IGUAL, al Componente del anterior error...
    If Componente = HistorialError.Componente And _
       Numero = HistorialError.ErrorCode Then
        
        'Agregamos el error al historial.
        HistorialError.Contador = HistorialError.Contador + 1
        
        'Si ya recibimos error en el mismo componente 10 veces, es bastante probable que estemos en un bucle
        'x lo que no hace falta registrar el error.
        If HistorialError.Contador = 10 Then Exit Sub
        
    Else 'Si NO es igual, reestablecemos el contador.

        HistorialError.Contador = 0
        HistorialError.ErrorCode = Numero
        HistorialError.Componente = Componente
            
    End If
    
    'Registramos el error en Errores.log
    Dim File As Integer: File = FreeFile
        
    Open CARPETA_LOGS & "Errores.log" For Append As #File
    
        Print #File, "Error: " & Numero
        Print #File, "Descripcion: " & Descripcion
        
        If LenB(Linea) <> 0 Then
            Print #File, "Linea: " & Linea
        End If
        
        Print #File, "Componente: " & Componente
        Print #File, "Fecha y Hora: " & Date$ & "-" & Time$
        
        Print #File, vbNullString
        
    Close #File
    
    Debug.Print "Error: " & Numero & vbNewLine & _
                "Descripcion: " & Descripcion & vbNewLine & _
                "Componente: " & Componente & vbNewLine & _
                "Fecha y Hora: " & Date$ & "-" & Time$ & vbNewLine
                
   Exit Sub
                
EH:

    Close #File
    
    Debug.Print "Error: " & Numero & vbNewLine & _
                "Descripcion: " & Descripcion & vbNewLine & _
                "Componente: " & Componente & vbNewLine & _
                "Fecha y Hora: " & Date$ & "-" & Time$ & vbNewLine

End Sub

