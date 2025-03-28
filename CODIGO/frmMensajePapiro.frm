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
      Left            =   5175
      TabIndex        =   9
      Top             =   6000
      Width           =   1095
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
      Height          =   735
      Left            =   3285
      TabIndex        =   8
      Top             =   6000
      Width           =   1695
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
      Height          =   435
      Index           =   2
      Left            =   5310
      TabIndex        =   7
      Top             =   5150
      Width           =   1785
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
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   6
      Top             =   5150
      Width           =   1860
   End
   Begin VB.Label GameRulesLink 
      BackStyle       =   0  'Transparent
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
      Height          =   150
      Index           =   0
      Left            =   2950
      TabIndex        =   5
      Top             =   4480
      Width           =   4000
   End
   Begin VB.Label HelpGuideLink 
      BackStyle       =   0  'Transparent
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
      Height          =   150
      Left            =   2950
      TabIndex        =   4
      Top             =   4110
      Width           =   4000
   End
   Begin VB.Label BasicInfoLink 
      BackStyle       =   0  'Transparent
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
      Height          =   150
      Left            =   2950
      TabIndex        =   3
      Top             =   3730
      Width           =   4000
   End
   Begin VB.Label eventNews 
      BackStyle       =   0  'Transparent
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
      Height          =   150
      Left            =   2950
      TabIndex        =   2
      Top             =   3360
      Width           =   4000
   End
   Begin VB.Label newsLink 
      BackStyle       =   0  'Transparent
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
      Height          =   150
      Left            =   2950
      TabIndex        =   1
      Top             =   2940
      Width           =   4000
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
    Call OpenLink("https://www.argentumonline.com.ar/wiki")
End Sub

Private Sub CafecitoLink_Click(Index As Integer)
    Call OpenLink("https://cafecito.app/nolandstudios")
End Sub

Private Sub eventNews_Click()
    Call OpenLink("https://discord.com/channels/761213868352471040/1225542315507056762")
End Sub

Private Sub Form_Load()

    Me.Picture = LoadInterface("board.bmp")
    MakeFormTransparent Me, vbBlack    'Set the Form "transparent by color."

End Sub

Private Sub Form_LostFocus()
    Unload Me
End Sub

Private Sub GameRulesLink_Click(Index As Integer)
    Call OpenLink("https://www.argentumonline.com.ar/reglamento")
End Sub

Private Sub HelpGuideLink_Click()
    Call OpenLink("https://www.argentumonline.com.ar/wiki/guia-general/skills-clases")
End Sub

Private Sub Image1_Click()
    Unload Me
End Sub

Private Sub newsLink_Click()
    Call OpenLink("https://steamcommunity.com/app/1956740/allnews/")
End Sub

Private Sub OpenLink(link As String)
    ShellExecute ByVal 0&, "open", _
        link, _
        vbNullString, vbNullString, _
        vbMaximizedFocus
End Sub

Private Sub PatreonLink_Click(Index As Integer)
    Call OpenLink("https://www.patreon.com/nolandstudios")
End Sub
Private Sub Stream1_Click()
    Call OpenLink("https://www.twitch.tv/argentumonlineoficial")
End Sub

Private Sub Stream2_Click()
    Call OpenLink("https://www.twitch.tv/Plus_1986")
End Sub
Private Sub Stream2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Stream2.MousePointer = 99 ' Cursor personalizado
    Stream2.MouseIcon = LoadResPicture("104", vbResCursor) 'Carga el cursor de la mano
End Sub

Private Sub Stream1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Stream1.MousePointer = 99 ' Cursor personalizado
    Stream1.MouseIcon = LoadResPicture("104", vbResCursor) 'Carga el cursor de la mano
End Sub
Private Sub PatreonLink_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    PatreonLink(Index).MousePointer = 99 ' Cursor personalizado
    PatreonLink(Index).MouseIcon = LoadResPicture("104", vbResCursor) 'Carga el cursor de la mano
End Sub
Private Sub newsLink_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    newsLink.MousePointer = 99 ' Cursor personalizado
    newsLink.MouseIcon = LoadResPicture("104", vbResCursor) 'Carga el cursor de la mano
End Sub
Private Sub HelpGuideLink_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HelpGuideLink.MousePointer = 99 ' Cursor personalizado
    HelpGuideLink.MouseIcon = LoadResPicture("104", vbResCursor) 'Carga el cursor de la mano
End Sub
Private Sub GameRulesLink_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    GameRulesLink(Index).MousePointer = 99 ' Cursor personalizado
    GameRulesLink(Index).MouseIcon = LoadResPicture("104", vbResCursor) 'Carga el cursor de la mano
End Sub
Private Sub eventNews_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    eventNews.MousePointer = 99 ' Cursor personalizado
    eventNews.MouseIcon = LoadResPicture("104", vbResCursor) 'Carga el cursor de la mano
End Sub
Private Sub BasicInfoLink_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    BasicInfoLink.MousePointer = 99 ' Cursor personalizado
    BasicInfoLink.MouseIcon = LoadResPicture("104", vbResCursor) 'Carga el cursor de la mano
End Sub
Private Sub CafecitoLink_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    CafecitoLink(Index).MousePointer = 99 ' Cursor personalizado
    CafecitoLink(Index).MouseIcon = LoadResPicture("104", vbResCursor) 'Carga el cursor de la mano
End Sub
