Attribute VB_Name = "ModUrl"
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    'Geodar
Const EM_SETEVENTMASK = &H445
Const EN_LINK = &H70B
Const ENM_LINK = &H4000000
Const EM_AUTOURLDETECT = &H45B
Const EM_GETEVENTMASK = &H43B
Const GWL_WNDPROC = (-4) 'Geodar
Const WM_NOTIFY = &H4E 'Geodar
Const WM_LBUTTONDOWN = &H201
Const EM_GETTEXTRANGE = &H44B
 
Dim lOldProc As Long
Dim hWndRTB As Long
Dim hWndParent As Long
'Geodar
Private Type NMHDR 'Geodar
    hWndFrom As Long
    idFrom As Long
    code As Long
End Type 'Geodar
 
Private Type CHARRANGE
    cpMin As Long
    cpMax As Long
End Type
'Geodar
Private Type ENLINK
    hdr As NMHDR
    msg As Long
    wParam As Long 'Geodar
    lParam As Long
    chrg As CHARRANGE 'Geodar
End Type
'Geodar
Private Type TEXTRANGE
    chrg As CHARRANGE
    lpstrText As String
End Type
 
 
Public Sub Detectar(ByVal hWndTextbox As Long, ByVal hWndOwner As Long)
  'Don't want to subclass twice!
  If lOldProc = 0 Then
    'Subclass!
    lOldProc = SetWindowLong(hWndOwner, GWL_WNDPROC, AddressOf WndProc)
    SendMessage hWndTextbox, EM_SETEVENTMASK, 0, ByVal ENM_LINK Or SendMessage(hWndTextbox, EM_GETEVENTMASK, 0, 0)
    SendMessage hWndTextbox, EM_AUTOURLDETECT, 1, ByVal 0
    hWndParent = hWndOwner 'Geodar
    hWndRTB = hWndTextbox
  End If
End Sub
 
Public Sub NoDetectar()
  If lOldProc Then
    SendMessage hWndRTB, EM_AUTOURLDETECT, 0, ByVal 0
    'Reset the window procedure (stop the subclassing)
    SetWindowLong hWndParent, GWL_WNDPROC, lOldProc
    'Set this to 0 so we can subclass again in future
    lOldProc = 0
  End If
End Sub
 
Public Function WndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim uHead As NMHDR 'Geodar
    Dim eLink As ENLINK
    Dim eText As TEXTRANGE
    Dim sText As String
    Dim lLen As Long 'Geodar
   
    'Which message?
    Select Case uMsg 'Geodar
    Case WM_NOTIFY
        'Copy the notification header into our structure from the pointer
        CopyMemory uHead, ByVal lParam, Len(uHead)
       
        If (uHead.hWndFrom = hWndRTB) And (uHead.code = EN_LINK) Then
            CopyMemory eLink, ByVal lParam, Len(eLink)
           
            'What kind of message?
            Select Case eLink.msg 'Geodar
           
            Case WM_LBUTTONDOWN
              eText.chrg.cpMin = eLink.chrg.cpMin
              eText.chrg.cpMax = eLink.chrg.cpMax 'Geodar
              eText.lpstrText = Space$(1024)
              lLen = SendMessage(hWndRTB, EM_GETTEXTRANGE, 0, eText)
              sText = left$(eText.lpstrText, lLen)
              ShellExecute hWndParent, vbNullString, sText, vbNullString, vbNullString, SW_SHOW
               
            End Select
           
        End If 'Geodar
       
    End Select
    WndProc = CallWindowProc(lOldProc, hwnd, uMsg, wParam, lParam) 'Geodar
End Function
 
