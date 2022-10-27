VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGameState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Public Enum e_game_state
    e_state_gameplay_screen = 0
    e_state_connect_screen = 1
    e_state_account_screen = 2
    e_state_createchar_screen = 3
End Enum

Private cur_state As e_game_state

Private Sub Class_Initialize()
cur_state = e_state_connect_screen
End Sub

Public Property Get state() As e_game_state
    state = cur_state
End Property

Public Property Let state(ByVal new_state As e_game_state)
    cur_state = new_state
End Property


