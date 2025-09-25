Attribute VB_Name = "modUrlDetection"
' Argentum 20 Game Client
'
'    Copyright (C) 2023 Noland Studios LTD
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
'
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
'    This program was based on Argentum Online 0.11.6
'    Copyright (C) 2002 Márquez Pablo Ignacio
'
'    Argentum Online is based on Baronsoft's VB6 Online RPG
'    You can contact the original creator of ORE at aaron@baronsoft.com
'    for more information about ORE please visit http://www.baronsoft.com/
'
'
'
Option Explicit

Private Type NMHDR
    hWndFrom As Long
    idFrom As Long
    code As Long
End Type

Private Type CHARRANGE
    cpMin As Long
    cpMax As Long
End Type

Private Type ENLINK
    hdr As NMHDR
    msg As Long
    wParam As Long
    lParam As Long
    chrg As CHARRANGE
End Type

Private Type TEXTRANGE
    chrg As CHARRANGE
    lpstrText As String
End Type

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc _
                Lib "user32" _
                Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                                         ByVal hWnd As Long, _
                                         ByVal msg As Long, _
                                         ByVal wParam As Long, _
                                         ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function ShellExecute _
                Lib "shell32" _
                Alias "ShellExecuteA" (ByVal hWnd As Long, _
                                       ByVal lpOperation As String, _
                                       ByVal lpFile As String, _
                                       ByVal lpParameters As String, _
                                       ByVal lpDirectory As String, _
                                       ByVal nShowCmd As Long) As Long
Private Const WM_NOTIFY = &H4E
Private Const EM_SETEVENTMASK = &H445
Private Const EM_GETEVENTMASK = &H43B
Private Const EM_GETTEXTRANGE = &H44B
Private Const EM_AUTOURLDETECT = &H45B
Private Const EN_LINK = &H70B
Private Const WM_LBUTTONDOWN = &H201
Private Const ENM_LINK = &H4000000
Private Const GWL_WNDPROC = (-4)
Private Const SW_SHOW = 5
Private lOldProc   As Long
Private hWndRTB    As Long
Private hWndParent As Long

Public Function WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error GoTo WndProc_Err
    'Get "Click" event on link and open browser.
    Dim uHead As NMHDR
    Dim eLink As ENLINK
    Dim eText As TEXTRANGE
    Dim sText As String
    Dim lLen  As Long
    If uMsg = WM_NOTIFY Then
        Call CopyMemory(uHead, ByVal lParam, Len(uHead))
        If (uHead.hWndFrom = hWndRTB) And (uHead.code = EN_LINK) Then
            Call CopyMemory(eLink, ByVal lParam, Len(eLink))
            Select Case eLink.msg
                Case WM_LBUTTONDOWN
                    eText.chrg.cpMin = eLink.chrg.cpMin
                    eText.chrg.cpMax = eLink.chrg.cpMax
                    eText.lpstrText = Space$(1024)
                    lLen = SendMessage(hWndRTB, EM_GETTEXTRANGE, 0, eText)
                    sText = Left$(eText.lpstrText, lLen)
                    Call ShellExecute(hWndParent, vbNullString, sText, vbNullString, vbNullString, SW_SHOW)
            End Select
        End If
    End If
    WndProc = CallWindowProc(lOldProc, hWnd, uMsg, wParam, lParam)
    Exit Function
WndProc_Err:
    Call RegistrarError(Err.Number, Err.Description, "modUrlDetection.WndProc", Erl)
    Resume Next
End Function

Public Sub EnableURLDetect(ByVal hWndRichTextbox As Long, ByVal hWndOwner As Long)
    'Enables url detection in richtexbox.
    On Error GoTo EnableURLDetect_Err
    SendMessage hWndRichTextbox, EM_SETEVENTMASK, 0, ByVal ENM_LINK Or SendMessage(hWndRichTextbox, EM_GETEVENTMASK, 0, 0)
    SendMessage hWndRichTextbox, EM_AUTOURLDETECT, 1, ByVal 0
    hWndParent = hWndOwner
    hWndRTB = hWndRichTextbox
    Exit Sub
EnableURLDetect_Err:
    Call RegistrarError(Err.Number, Err.Description, "modUrlDetection.EnableURLDetect", Erl)
    Resume Next
End Sub

Public Sub DisableURLDetect()
    'Disables url detection in richtexbox.
    On Error GoTo DisableURLDetect_Err
    SendMessage hWndRTB, EM_AUTOURLDETECT, 0, ByVal 0
    StopCheckingLinks
    Exit Sub
DisableURLDetect_Err:
    Call RegistrarError(Err.Number, Err.Description, "modUrlDetection.DisableURLDetect", Erl)
    Resume Next
End Sub

Public Sub StartCheckingLinks()
    On Error GoTo StartCheckingLinks_Err
    'Starts checking links (in console range)
    If lOldProc = 0 Then lOldProc = SetWindowLong(hWndParent, GWL_WNDPROC, AddressOf WndProc)
    Exit Sub
StartCheckingLinks_Err:
    Call RegistrarError(Err.Number, Err.Description, "modUrlDetection.StartCheckingLinks", Erl)
    Resume Next
End Sub

Public Sub StopCheckingLinks()
    On Error GoTo StopCheckingLinks_Err
    'Stops checking links (out of console range)
    If lOldProc Then
        SetWindowLong hWndParent, GWL_WNDPROC, lOldProc
        lOldProc = 0
    End If
    Exit Sub
StopCheckingLinks_Err:
    Call RegistrarError(Err.Number, Err.Description, "modUrlDetection.StopCheckingLinks", Erl)
    Resume Next
End Sub
