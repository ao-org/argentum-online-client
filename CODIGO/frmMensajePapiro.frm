VERSION 5.00
Begin VB.Form frmMensajePapiro 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9270
   ScaleWidth      =   11580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Stream2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Plus"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   6720
      TabIndex        =   9
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Stream1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Argentum Online Oficial"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   6360
      TabIndex        =   8
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label CafecitoLink 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ayudar con Cafecito"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Index           =   2
      Left            =   4560
      TabIndex        =   7
      Top             =   6225
      Width           =   1095
   End
   Begin VB.Label PatreonLink 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ayudar con Patreon"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Index           =   1
      Left            =   3285
      TabIndex        =   6
      Top             =   6225
      Width           =   1020
   End
   Begin VB.Label GameRulesLink 
      BackStyle       =   0  'Transparent
      Caption         =   "Ir al reglamento"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   3285
      TabIndex        =   5
      Top             =   5625
      Width           =   1455
   End
   Begin VB.Label HelpGuideLink 
      BackStyle       =   0  'Transparent
      Caption         =   "Ir a Guía"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3285
      TabIndex        =   4
      Top             =   5025
      Width           =   1335
   End
   Begin VB.Label BasicInfoLink 
      BackStyle       =   0  'Transparent
      Caption         =   "Ir al manual"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3285
      TabIndex        =   3
      Top             =   4420
      Width           =   5535
   End
   Begin VB.Label eventNews 
      BackStyle       =   0  'Transparent
      Caption         =   "Ver próximos eventos"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3285
      TabIndex        =   2
      Top             =   3825
      Width           =   1815
   End
   Begin VB.Label newsLink 
      BackStyle       =   0  'Transparent
      Caption         =   "Ir a novedades"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3285
      TabIndex        =   1
      Top             =   3225
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7335
      Left            =   1560
      TabIndex        =   0
      Top             =   1320
      Width           =   8655
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   9840
      Top             =   720
      Width           =   495
   End
End
Attribute VB_Name = "frmMensajePapiro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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
Option Explicit

Private Declare Function GetWindowLong Lib "user32" _
    Alias "GetWindowLongW" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long) As Long
    
Private Declare Function SetWindowLong Lib "user32" _
    Alias "SetWindowLongW" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long
    
Private Declare Function SetLayeredWindowAttributes Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal crKey As Long, _
    ByVal bAlpha As Byte, _
    ByVal dwFlags As Long) As Long
    
Private Const GWL_EXSTYLE = -20
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_COLORKEY = &H1&

Private Sub BasicInfoLink_Click()
    On Error Goto BasicInfoLink_Click_Err
    Call OpenLink("https://www.argentumonline.com.ar/wiki")
    Exit Sub
BasicInfoLink_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmMensajePapiro.BasicInfoLink_Click", Erl)
End Sub

Private Sub CafecitoLink_Click(Index As Integer)
    On Error Goto CafecitoLink_Click_Err
    Call OpenLink("https://cafecito.app/nolandstudios")
    Exit Sub
CafecitoLink_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmMensajePapiro.CafecitoLink_Click", Erl)
End Sub

Private Sub eventNews_Click()
    On Error Goto eventNews_Click_Err
    Call OpenLink("https://discord.com/channels/761213868352471040/1225542315507056762")
    Exit Sub
eventNews_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmMensajePapiro.eventNews_Click", Erl)
End Sub

Private Sub Form_Load()
    On Error Goto Form_Load_Err

    Me.Picture = LoadInterface("board.bmp")
    MakeFormTransparent Me, vbBlack    'Set the Form "transparent by color."

    Exit Sub
Form_Load_Err:
    Call TraceError(Err.Number, Err.Description, "frmMensajePapiro.Form_Load", Erl)
End Sub

Private Sub Form_LostFocus()
    On Error Goto Form_LostFocus_Err
    Unload Me
    Exit Sub
Form_LostFocus_Err:
    Call TraceError(Err.Number, Err.Description, "frmMensajePapiro.Form_LostFocus", Erl)
End Sub

Private Sub GameRulesLink_Click(Index As Integer)
    On Error Goto GameRulesLink_Click_Err
    Call OpenLink("https://www.argentumonline.com.ar/reglamento")
    Exit Sub
GameRulesLink_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmMensajePapiro.GameRulesLink_Click", Erl)
End Sub

Private Sub HelpGuideLink_Click()
    On Error Goto HelpGuideLink_Click_Err
    Call OpenLink("https://www.argentumonline.com.ar/wiki/guia-general/skills-clases")
    Exit Sub
HelpGuideLink_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmMensajePapiro.HelpGuideLink_Click", Erl)
End Sub

Private Sub Image1_Click()
    On Error Goto Image1_Click_Err
    Unload Me
    Exit Sub
Image1_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmMensajePapiro.Image1_Click", Erl)
End Sub


Private Sub newsLink_Click()
    On Error Goto newsLink_Click_Err
    Call OpenLink("https://steamcommunity.com/app/1956740/allnews/")
    Exit Sub
newsLink_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmMensajePapiro.newsLink_Click", Erl)
End Sub

Private Sub OpenLink(link As String)
    On Error Goto OpenLink_Err
    ShellExecute ByVal 0&, "open", _
        link, _
        vbNullString, vbNullString, _
        vbMaximizedFocus
    Exit Sub
OpenLink_Err:
    Call TraceError(Err.Number, Err.Description, "frmMensajePapiro.OpenLink", Erl)
End Sub

Private Sub PatreonLink_Click(Index As Integer)
    On Error Goto PatreonLink_Click_Err
    Call OpenLink("https://www.patreon.com/nolandstudios")
    Exit Sub
PatreonLink_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmMensajePapiro.PatreonLink_Click", Erl)
End Sub

Private Sub Stream1_Click()
    On Error Goto Stream1_Click_Err
    Call OpenLink("https://www.twitch.tv/argentumonlineoficial")
    Exit Sub
Stream1_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmMensajePapiro.Stream1_Click", Erl)
End Sub

Private Sub Stream2_Click()
    On Error Goto Stream2_Click_Err
    Call OpenLink("https://www.twitch.tv/Plus_1986")
    Exit Sub
Stream2_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmMensajePapiro.Stream2_Click", Erl)
End Sub
