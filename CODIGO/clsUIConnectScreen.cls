VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUIConnectScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Argentum 20 Game Client
'
'    Copyright (C) 2025 Noland Studios LTD
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
Private Const WINDOW_WIDTH    As Long = 300
Private Const WINDOW_HEIGHT   As Long = 360
Private Const WINDOW_OFFSET_Y As Long = 100  ' adjust this value to move window lower
Private m_UIMgr               As clsUIManager

' Ventna Connect
Private m_lblTitle              As clsUILabel
Private m_imgSeparatorH1        As clsUIImage
Private m_imgSeparatorH2        As clsUIImage
Private m_lblPass               As clsUILabel
Private m_cbOk                  As clsUICheckbox
Private m_btnConnect            As clsUIButton
Private m_inputboxUser          As clsUIInputbox
Private m_inputboxPass          As clsUIInputbox
Private m_boxInfo               As clsUIMsgBox
Private m_Action                As Boolean
Private m_lblNewAccount         As clsUILabel
Private m_lblForgotPass         As clsUILabel

Private m_frameConnect          As clsUIFrame

' Ventana Crear Cuenta
Private m_lblTitle2             As clsUILabel
Private m_btnCreateAccount      As clsUIButton
Private m_btnCloseCA            As clsUIButton
Private m_imgSeparatorH3        As clsUIImage
Private m_imgSeparatorH4        As clsUIImage
Private m_inputboxName          As clsUIInputbox
Private m_inputboxLastName      As clsUIInputbox
Private m_inputboxEmail          As clsUIInputbox
Private m_inputboxCCPass      As clsUIInputbox
Private m_lblCreateAcctNext     As clsUILabel

Private m_frameCreateUser       As clsUIFrame


Private Const EVENT_ID_NOP                                      As Integer = 0
Private Const EVENT_ID_CREATE_ACCOUNT_BUTTON_CLOSE_CLICKED      As Integer = 1
Private Const EVENT_ID_BUTTON_CREATE_ACCOUNT_CLICKED            As Integer = 2
Private Const EVENT_ID_CONNECT_BUTTON_CLICKED                   As Integer = 3
Private Const EVENT_ID_BUTTON_MSG_BOX_OK_CLICKED                As Integer = 4
Private Const EVENT_ID_CREATE_ACCOUNT_NEXT_CLICKED              As Integer = 5
Private Const EVENT_ID_CREATE_ACCOUNT_BACK_CLICKED              As Integer = 6



Private Sub ProcessAction(ByVal actionID As Integer)
    Select Case actionID
        Case EVENT_ID_CREATE_ACCOUNT_BUTTON_CLOSE_CLICKED
            m_frameCreateUser.IUIComponent_SetVisible False
        Case EVENT_ID_BUTTON_CREATE_ACCOUNT_CLICKED
             m_frameCreateUser.IUIComponent_SetVisible True
             m_frameCreateUser.ChangeState 1
        Case EVENT_ID_CONNECT_BUTTON_CLICKED
            'Call SetActiveServer("127.0.0.1", "6501")
            Call DoLogin(m_inputboxUser.text, m_inputboxPass.text, True)
        Case EVENT_ID_BUTTON_MSG_BOX_OK_CLICKED
            m_frameCreateUser.ChangeState 1
        Case EVENT_ID_CREATE_ACCOUNT_NEXT_CLICKED
            m_frameCreateUser.ChangeState 2
        Case EVENT_ID_CREATE_ACCOUNT_BACK_CLICKED
            m_frameCreateUser.ChangeState 1
        Case Else
            Debug.Assert False ' must not happen!
    End Select
End Sub

 
Public Sub Init(ByVal dev As Direct3DDevice8)
    Set m_UIMgr = New clsUIManager: m_UIMgr.Init
    
    Dim sW As Long, sH As Long, wx As Long, wy As Long
    sW = Render_Connect_Rect.Right - Render_Connect_Rect.Left
    sH = Render_Connect_Rect.Bottom - Render_Connect_Rect.Top
    wx = (sW - WINDOW_WIDTH) \ 2 + Render_Connect_Rect.Left
    wy = (sH - WINDOW_HEIGHT) \ 2 + Render_Connect_Rect.Top + WINDOW_OFFSET_Y
    
    ' """"""""""""   CONNECT FRAME   """""""""""""
    Set m_frameConnect = New clsUIFrame
    m_frameConnect.Init wx, wy, 300, 360
    Set m_lblTitle = New clsUILabel
    m_lblTitle.Init "INGRESAR", 0.3, 0.05, UI_COLOR_WHITE, 255, 0.5, 8, 0
    m_frameConnect.AddUIElement m_lblTitle
    Set m_imgSeparatorH1 = New clsUIImage
    m_imgSeparatorH1.Init "", 0.1, 0.15, 0.8, 0.01, UI_COLOR_GOLD, &HFF606060, UI_COLOR_WHITE, 0
    m_frameConnect.AddUIElement m_imgSeparatorH1
    Set m_inputboxUser = New clsUIInputbox
    m_inputboxUser.Init "Email", 0.1, 0.22, 0.8, 0.1, &HFFFFFF, &HFFFFFF, UI_COLOR_GOLD, False, 0
    m_frameConnect.AddUIElement m_inputboxUser
    Set m_inputboxPass = New clsUIInputbox
    m_inputboxPass.Init "Clave", 0.1, 0.38, 0.8, 0.1, &HFFFFFF, &HFFFFFF, UI_COLOR_GOLD, True, 0
    m_frameConnect.AddUIElement m_inputboxPass
    Set m_cbOk = New clsUICheckbox
    m_cbOk.Init "Recordarme", 0.1, 0.55, 0.1, 0.07, &HFFFFFF, &HFFFF00, UI_COLOR_WHITE, 0
    m_frameConnect.AddUIElement m_cbOk
    Set m_boxInfo = New clsUIMsgBox
    m_boxInfo.Init "Default text...", 0, 0, 400, 150, &HFFFF00, &HFFFFFF, UI_COLOR_BLACK, EVENT_ID_BUTTON_MSG_BOX_OK_CLICKED
    m_frameConnect.AddUIElement m_boxInfo
    Set m_imgSeparatorH2 = New clsUIImage
    m_imgSeparatorH2.Init "", 0.1, 0.65, 0.85, 0.01, UI_COLOR_GOLD, &HFF606060, UI_COLOR_WHITE, 0
    m_frameConnect.AddUIElement m_imgSeparatorH2
    Set m_btnConnect = New clsUIButton
    m_btnConnect.Init "Conectar", 0.1, 0.73, 0.8, 0.13, UI_COLOR_GOLD, &HFF606060, UI_COLOR_WHITE, EVENT_ID_CONNECT_BUTTON_CLICKED
    m_frameConnect.AddUIElement m_btnConnect
    Set m_lblNewAccount = New clsUILabel
    m_lblNewAccount.Init "Crear cuenta", 0.1, 0.9, UI_COLOR_WHITE, 100, 1, 1, EVENT_ID_BUTTON_CREATE_ACCOUNT_CLICKED
    m_frameConnect.AddUIElement m_lblNewAccount
    Set m_lblForgotPass = New clsUILabel
    m_lblForgotPass.Init "Olvide mi contraseña", 0.5, 0.9, UI_COLOR_WHITE, 100, 1, 1, EVENT_ID_BUTTON_CREATE_ACCOUNT_CLICKED
    m_frameConnect.AddUIElement m_lblForgotPass
    m_UIMgr.AddComponent m_frameConnect
    
    ' """"""""""""   CREATE ACCOUNT PANEL   """""""""""""
    Set m_frameCreateUser = New clsUIFrame
    m_frameCreateUser.Init 50, 50, 300, 300
    m_frameCreateUser.IUIComponent_SetVisible False
    Set m_lblTitle2 = New clsUILabel
    m_lblTitle2.Init "Crear Cuenta", 0.3, 0.08, UI_COLOR_WHITE, 255, 0.5, 8, 0
    m_lblTitle2.IUIComponent_SetVisible False
    m_frameCreateUser.AddUIElement m_lblTitle2
    ' H LINE
    Set m_imgSeparatorH3 = New clsUIImage
    m_imgSeparatorH3.IUIComponent_SetVisible False
    m_imgSeparatorH3.Init "", 0.1, 0.23, 0.8, 0.01, UI_COLOR_GOLD, &HFF606060, UI_COLOR_WHITE, 0
    m_frameCreateUser.AddUIElement m_imgSeparatorH3
    ' INPUTBOX
    Set m_inputboxName = New clsUIInputbox
    m_inputboxName.Init "Nombre", 0.1, 0.28, 0.8, 0.1, &HFFFFFF, &HFFFFFF, UI_COLOR_GOLD, False, 0
    m_frameCreateUser.AddUIElement m_inputboxName
    Set m_inputboxLastName = New clsUIInputbox
    m_inputboxLastName.Init "Apellido", 0.1, 0.46, 0.8, 0.1, &HFFFFFF, &HFFFFFF, UI_COLOR_GOLD, True, 0
    m_frameCreateUser.AddUIElement m_inputboxLastName
    ' INPUTBOX 2
    Set m_inputboxName = New clsUIInputbox
    m_inputboxName.Init "Email", 0.1, 0.28, 0.8, 0.1, &HFFFFFF, &HFFFFFF, UI_COLOR_GOLD, False, 0
    m_inputboxName.IUIComponent_SetVisible False
    m_frameCreateUser.AddUIElement m_inputboxName
    Set m_inputboxLastName = New clsUIInputbox
    m_inputboxLastName.Init "Clave", 0.1, 0.46, 0.8, 0.1, &HFFFFFF, &HFFFFFF, UI_COLOR_GOLD, True, 0
    m_inputboxLastName.IUIComponent_SetVisible False
    m_frameCreateUser.AddUIElement m_inputboxLastName
    ' H LINE
    Set m_imgSeparatorH4 = New clsUIImage
    m_imgSeparatorH4.IUIComponent_SetVisible False
    m_imgSeparatorH4.Init "", 0.1, 0.63, 0.8, 0.01, UI_COLOR_GOLD, &HFF606060, UI_COLOR_WHITE, 0
    m_frameCreateUser.AddUIElement m_imgSeparatorH4
    ' Next Button
    Set m_lblCreateAcctNext = New clsUILabel
    m_lblCreateAcctNext.Init ">>", 0.8, 0.65, UI_COLOR_WHITE, 100, 2, 1, EVENT_ID_CREATE_ACCOUNT_NEXT_CLICKED
    m_frameCreateUser.AddUIElement m_lblCreateAcctNext
    ' Back Button
    Set m_lblCreateAcctNext = New clsUILabel
    m_lblCreateAcctNext.Init "<<", 0.1, 0.65, UI_COLOR_WHITE, 100, 2, 1, EVENT_ID_CREATE_ACCOUNT_BACK_CLICKED
    m_frameCreateUser.AddUIElement m_lblCreateAcctNext
    Set m_btnCloseCA = New clsUIButton
    m_btnCloseCA.Init "X", 0.9, 0.02, 0.08, 0.08, UI_COLOR_GOLD, &HFF606060, UI_COLOR_WHITE, EVENT_ID_CREATE_ACCOUNT_BUTTON_CLOSE_CLICKED
    m_btnCloseCA.IUIComponent_SetVisible False
    m_frameCreateUser.AddUIElement m_btnCloseCA
    Set m_btnCreateAccount = New clsUIButton
    m_btnCreateAccount.Init "Crear", 0.1, 0.8, 0.8, 0.1, UI_COLOR_GOLD, &HFF606060, UI_COLOR_WHITE, 0
    m_btnCreateAccount.IUIComponent_SetVisible False
    m_frameCreateUser.AddUIElement m_btnCreateAccount
    m_UIMgr.AddComponent m_frameCreateUser
    
End Sub




Public Sub render(ByVal dev As Direct3DDevice8)

    Call Engine_BeginScene

        ' World/background
        RenderConnect 57, 45, 0, 0


        m_UIMgr.RenderAll dev
        UIRenderer.Draw DirectDevice

    Call Engine_EndScene(Render_Connect_Rect, frmConnect.render.hWnd)
End Sub


Public Sub HandleMouse(ByVal mouseX As Long, ByVal mouseY As Long, ByVal mouseButtons As Long)
    Dim actionID As Integer
    m_UIMgr.HandleMouse mouseX, mouseY, mouseButtons, actionID
    If actionID <> EVENT_ID_NOP Then
        ProcessAction actionID
    End If
End Sub

Public Sub HandleKeyboard(ByVal keyCode As Long, ByVal shiftState As Long)
    m_UIMgr.HandleKeyboard keyCode, shiftState
End Sub


Public Sub ShowMessageBox(ByVal title, ByVal message As String)
    m_boxInfo.IUIComponent_SetVisible True
    m_boxInfo.message = message
    m_boxInfo.title = title
End Sub
