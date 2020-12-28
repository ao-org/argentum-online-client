Attribute VB_Name = "ModProcesos"
Option Explicit

Private Const TH32CS_SNAPPROCESS As Long = &H2
Private Const TH32CS_SNAPMODULE As Long = &H8
Private Const MAX_PATH As Integer = 260
Private Const MAX_MODULE_NAME32 As Integer = 256
Private Const GW_HWNDFIRST = 0&
Private Const GW_HWNDNEXT = 2&
Private Const GW_CHILD = 5&
 
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type

Private Type MODULEENTRY32
    dwSize As Long
    th32ModuleID As Long
    th32ProcessID As Long
    GlblcntUsage As Long
    ProccntUsage As Long
    modBaseAddr As Long
    modBaseSize As Long
    hModule As Long
    szModuleName As String * MAX_MODULE_NAME32
    szExeFile As String * MAX_PATH
End Type

Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ModuleFirst Lib "kernel32" Alias "Module32First" (ByVal hSnapShot As Long, uProcess As MODULEENTRY32) As Long
Private Declare Function ModuleNext Lib "kernel32" Alias "Module32Next" (ByVal hSnapShot As Long, uProcess As MODULEENTRY32) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wFlag As Long) As Long
Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)

Public Function ListarVentanasVisibles() As String
    On Error GoTo Handler

    Dim Buf As Long, Handle As Long, titulo As String, LenT As Long, Ret As Long

    Handle = GetWindow(frmMain.hWnd, GW_HWNDFIRST)

    Do While Handle <> 0
        If IsWindowVisible(Handle) Then
            LenT = GetWindowTextLength(Handle)

            If LenT > 0 Then
                titulo = String$(LenT, 0)
                Ret = GetWindowText(Handle, titulo, LenT + 1)
                titulo = Left$(titulo, Ret)

                ListarVentanasVisibles = ListarVentanasVisibles & titulo & vbNewLine
            End If
        End If

        Handle = GetWindow(Handle, GW_HWNDNEXT)
    Loop
    
    Exit Function
    
Handler:
    ListarVentanasVisibles = "** Error al listar ventanas **"
End Function

Public Function ListarProcesos() As String
    On Error GoTo Handler
    
    Dim hSnapShot As Long
    Dim uProcess As PROCESSENTRY32
    Dim r As Long

    hSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    
    If hSnapShot = 0 Then
        ListarProcesos = "** Error al listar procesos **"
        Exit Function
    End If

    uProcess.dwSize = Len(uProcess)
    r = ProcessFirst(hSnapShot, uProcess)

    While r <> 0

        ListarProcesos = ListarProcesos & Left$(uProcess.szExeFile, InStr(1, uProcess.szExeFile, Chr(0)) - 1) & vbNewLine

        r = ProcessNext(hSnapShot, uProcess)

    Wend

    Call CloseHandle(hSnapShot)
    
    Exit Function
    
Handler:
    ListarProcesos = "** Error al listar procesos **"
End Function

Public Function ListarModulos() As String
    On Error GoTo Handler
    
    Dim hSnapShot As Long
    Dim uModule As MODULEENTRY32
    Dim r As Long

    hSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPMODULE, 0&)
    
    If hSnapShot = 0 Then
        ListarModulos = "** Error al listar módulos **"
        Exit Function
    End If

    uModule.dwSize = Len(uModule)
    r = ModuleFirst(hSnapShot, uModule)

    While r <> 0

        ListarModulos = ListarModulos & Left$(uModule.szModuleName, InStr(1, uModule.szModuleName, Chr(0)) - 1) & vbNewLine

        r = ModuleNext(hSnapShot, uModule)

    Wend

    Call CloseHandle(hSnapShot)
    
    Exit Function
    
Handler:
    ListarModulos = "** Error al listar módulos **"
End Function
 
Public Function GetProcessesList() As String
    On Error GoTo Handler
    
    GetProcessesList = "## Ventanas visibles: ##" & vbNewLine & ListarVentanasVisibles & vbNewLine
    
    GetProcessesList = GetProcessesList & "####" & vbNewLine & vbNewLine
    
    GetProcessesList = GetProcessesList & "## Lista de procesos: ##" & vbNewLine & ListarProcesos & vbNewLine
    
    GetProcessesList = GetProcessesList & "####" & vbNewLine & vbNewLine
    
    GetProcessesList = GetProcessesList & "## Módulos/DLLs del juego: ##" & vbNewLine & ListarModulos

    Exit Function
    
Handler:
    GetProcessesList = "ERROR"

End Function
 

